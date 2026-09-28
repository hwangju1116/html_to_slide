from typing import Optional, Tuple
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

