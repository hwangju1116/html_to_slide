import re
from typing import Any, Dict, Optional, Tuple
from pptx.dml.color import RGBColor


def hex_to_rgb_tuple(
    hex_str: Optional[str],
    default: Tuple[int, int, int] = (30, 41, 59),
) -> Tuple[int, int, int]:
    """Converts a #RRGGBB or #RGB hex string to an (r, g, b) integer tuple.

    Args:
        hex_str: Hex color string (e.g. '#FFFFFF', '#fff', '1E293B').
        default: Fallback (r, g, b) tuple if parsing fails.

    Returns:
        (r, g, b) tuple where each component is 0-255.
    """
    if not hex_str or not isinstance(hex_str, str):
        return default
    hex_clean = hex_str.strip().lstrip("#")
    if len(hex_clean) == 3:
        hex_clean = "".join([c * 2 for c in hex_clean])
    if len(hex_clean) != 6:
        return default
    try:
        return (
            int(hex_clean[0:2], 16),
            int(hex_clean[2:4], 16),
            int(hex_clean[4:6], 16),
        )
    except ValueError:
        return default


def hex_to_pptx_color(
    hex_str: Optional[str],
    default: Tuple[int, int, int] = (30, 41, 59),
) -> RGBColor:
    """Converts a hex string to a python-pptx RGBColor object."""
    r, g, b = hex_to_rgb_tuple(hex_str, default=default)
    return RGBColor(r, g, b)


def extract_root_css_vars(html_content: str) -> Dict[str, str]:
    """Extracts all :root CSS custom properties (--var: val) from HTML."""
    var_map = {}
    matches = re.findall(r"(--[a-zA-Z0-9_-]+)\s*:\s*([^;}\n]+)", html_content)
    for k, v in matches:
        var_map[k.strip()] = v.strip()
    return var_map


def parse_color_value(
    val: Optional[str],
    default: Optional[RGBColor] = RGBColor(17, 17, 21),
    bg_blend_rgb: Tuple[int, int, int] = (255, 255, 255),
) -> Optional[RGBColor]:
    """Parses HEX string (#RRGGBB) or rgba(r,g,b,a) with configurable canvas alpha blending into RGBColor."""
    if not val or not isinstance(val, str):
        return default
    val = val.strip()
    if val.startswith("#"):
        h = val.lstrip("#")
        if len(h) == 3:
            h = "".join(c * 2 for c in h)
        if len(h) == 6:
            try:
                return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
            except ValueError:
                return default
    m_rgba = re.match(
        r"rgba?\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)(?:\s*,\s*([0-9.]+))?\s*\)", val
    )
    if m_rgba:
        try:
            r, g, b = int(m_rgba.group(1)), int(m_rgba.group(2)), int(m_rgba.group(3))
            a = float(m_rgba.group(4)) if m_rgba.group(4) else 1.0
            br = int(round(r * a + bg_blend_rgb[0] * (1.0 - a)))
            bg = int(round(g * a + bg_blend_rgb[1] * (1.0 - a)))
            bb = int(round(b * a + bg_blend_rgb[2] * (1.0 - a)))
            return RGBColor(min(255, max(0, br)), min(255, max(0, bg)), min(255, max(0, bb)))
        except Exception:
            return default
    return default


def resolve_style_color(
    style_str: str,
    css_vars_raw: Dict[str, str],
    default: RGBColor = RGBColor(17, 17, 21),
) -> RGBColor:
    """Resolves color from inline CSS style string using extracted :root CSS variables."""
    if not style_str:
        return default
    m_var = re.search(r"color:\s*var\((--[a-zA-Z0-9_-]+)\)", style_str)
    if m_var and m_var.group(1) in css_vars_raw:
        return parse_color_value(css_vars_raw[m_var.group(1)], default)
    m_hex = re.search(r"color:\s*(#[0-9a-fA-F]{3,8}|rgba?\([^)]+\))", style_str)
    if m_hex:
        return parse_color_value(m_hex.group(1), default)
    return default


def extract_global_design_tokens(
    first_image_path: str = "",
    raw_html: Optional[str] = None,
    first_slide_geometry: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """Extracts or derives unified design tokens (brand colors, font family, typography hierarchy) for the presentation deck."""
    tokens = {
        "font_name": "Pretendard",
        "brand_color_hex": "#2563EB",
        "text_main_hex": "#111115",
        "text_muted_hex": "#64748B",
        "bg_color_hex": "#FFFFFF",
        "badge_bg_hex": "#2563EB",
        "content_title_size_pt": 24.0,
        "content_title_left_in": 0.8,
        "content_title_top_in": 0.95,
        "badge_left_in": 0.8,
        "badge_top_in": 0.48,
        "footer_top_in": 6.75,
    }

    if raw_html:
        css_vars = extract_root_css_vars(raw_html)
        brand_found = False
        for brand_key in (
            "--brand-color",
            "--brand",
            "--app-primary",
            "--primary",
            "--accent",
            "--accent-color",
        ):
            if brand_key in css_vars and css_vars[brand_key].startswith("#"):
                tokens["brand_color_hex"] = css_vars[brand_key]
                tokens["badge_bg_hex"] = css_vars[brand_key]
                brand_found = True
                break
        if not brand_found:
            for k, v in css_vars.items():
                if v.startswith("#") and not any(
                    bg_tok in k.lower()
                    for bg_tok in ("bg", "background", "text", "muted", "border", "surface", "card")
                ):
                    rgb = hex_to_rgb_tuple(v, (0, 0, 0))
                    if max(rgb) - min(rgb) >= 80:
                        tokens["brand_color_hex"] = v
                        tokens["badge_bg_hex"] = v
                        break

        bg_found = False
        for bg_key in (
            "--app-background",
            "--background",
            "--bg-dark",
            "--bg",
            "--slide-bg",
            "--canvas-bg",
        ):
            if bg_key in css_vars and css_vars[bg_key].startswith("#"):
                tokens["bg_color_hex"] = css_vars[bg_key]
                bg_found = True
                break
        if not bg_found:
            for k, v in css_vars.items():
                if v.startswith("#") and ("bg" in k.lower() or "background" in k.lower()):
                    tokens["bg_color_hex"] = v
                    break

        for fg_key in ("--app-foreground", "--text-main", "--foreground", "--text-primary"):
            if fg_key in css_vars and css_vars[fg_key].startswith("#"):
                tokens["text_main_hex"] = css_vars[fg_key]
                break

        for muted_key in (
            "--app-muted-foreground",
            "--app-muted",
            "--text-muted",
            "--muted",
            "--text-secondary",
        ):
            if muted_key in css_vars and css_vars[muted_key].startswith("#"):
                tokens["text_muted_hex"] = css_vars[muted_key]
                break

        font_match = re.search(
            r"(?:body|html|\.slide)\s*\{[^}]*font-family\s*:\s*([^;\}]+)",
            raw_html,
            re.IGNORECASE | re.DOTALL,
        )
        if not font_match:
            font_match = re.search(
                r"font-family\s*:\s*([^;\}]+)", raw_html, re.IGNORECASE
            )
        if font_match:
            raw_families = [
                f.strip().strip("'\"")
                for f in font_match.group(1).split(",")
                if f.strip()
            ]
            generic_families = {
                "sans-serif",
                "serif",
                "monospace",
                "system-ui",
                "-apple-system",
                "blinkmacsystemfont",
            }
            for fam in raw_families:
                if fam.lower() not in generic_families and not fam.startswith("var("):
                    tokens["font_name"] = fam
                    break

    if first_slide_geometry and first_slide_geometry.get("slideBgColor"):
        measured_bg = parse_color_value(first_slide_geometry["slideBgColor"], default=None)
        if measured_bg is not None:
            tokens["bg_color_hex"] = f"#{measured_bg[0]:02X}{measured_bg[1]:02X}{measured_bg[2]:02X}"

    bg_tuple = hex_to_rgb_tuple(tokens.get("bg_color_hex", "#FFFFFF"), (255, 255, 255))
    is_dark = (bg_tuple[0] * 0.299 + bg_tuple[1] * 0.587 + bg_tuple[2] * 0.114) < 128
    if is_dark:
        if tokens["text_main_hex"] == "#111115":
            tokens["text_main_hex"] = "#F8FAFC"
        if tokens["text_muted_hex"] == "#64748B":
            tokens["text_muted_hex"] = "#94A3B8"
        if tokens["brand_color_hex"] == "#2563EB":
            tokens["brand_color_hex"] = "#3B82F6"
            tokens["badge_bg_hex"] = "#3B82F6"

    return tokens
