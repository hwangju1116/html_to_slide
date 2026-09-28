import asyncio
import base64
import concurrent.futures
import json
import os
import re
import socket
import subprocess
import tempfile
import time
from typing import Any, Dict, List, Optional
import urllib.request
import lxml.html
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt
from pydantic import BaseModel, Field
import websockets
from google import genai
from google.genai import types
from app.genai_client import DEFAULT_GEMINI_MODEL, get_genai_client

FONTS_DIR = os.path.join(
    os.path.dirname(os.path.dirname(__file__)), "assets", "fonts"
)
FALLBACK_FONTS_DIR = os.path.expanduser("~/.local/share/fonts")


class TextRun(BaseModel):
  text: str = Field(description="Sub-string text segment")
  is_bold: bool = Field(default=False, description="Whether this segment is bold")
  color_hex: Optional[str] = Field(default=None, description="HEX color e.g. #2563EB")
  is_code: bool = Field(default=False, description="Whether this segment is inline code or monospace")


class SlideNativeElement(BaseModel):
  element_type: str = Field(
      description=(
          "One of: 'CONTAINER', 'SUB_CARD', 'HEADER_BADGE', 'TEXT_BOX', 'ICON_BADGE', "
          "'STAT_METRIC', 'FLOW_NODE', 'FLOW_ARROW', 'BRIDGE_CONNECTOR', 'PIPELINE_BANNER', "
          "'PIPELINE_STEP', 'CODE_BLOCK', 'CALLOUT_BOX', 'FOOTNOTE', 'TABLE', 'DIVIDER', 'IMAGE'"
      )
  )
  text: Optional[str] = Field(
      default="", description="Verbatim text content with lines and bullet indents"
  )
  text_runs: Optional[List[TextRun]] = Field(
      default=None, description="Detailed inline text runs with rich styles"
  )
  # Normalized Coordinates (0.0 to 100.0%)
  left_pct: float = Field(description="Left X percentage of slide width (0-100)")
  top_pct: float = Field(description="Top Y percentage of slide height (0-100)")
  width_pct: float = Field(description="Width percentage of slide width (0-100)")
  height_pct: float = Field(description="Height percentage of slide height (0-100)")
  # Typography
  font_size_pt: Optional[float] = Field(default=14.0, description="Font size in points")
  is_bold: Optional[bool] = Field(default=False)
  font_color_hex: Optional[str] = Field(
      default="#0F172A", description="HEX text color, e.g. #1E293B"
  )
  alignment: Optional[str] = Field(
      default="LEFT", description="'LEFT', 'CENTER', 'RIGHT'"
  )
  # Container / Shape styling
  bg_color_hex: Optional[str] = Field(
      default=None, description="Fill background HEX color, e.g. #F8FAFC"
  )
  border_color_hex: Optional[str] = Field(
      default=None, description="Border stroke HEX color, e.g. #E2E8F0"
  )
  border_width_pt: Optional[float] = Field(default=1.0)
  border_radius_pt: Optional[float] = Field(default=0.0)
  shape_type: Optional[str] = Field(
      default="ROUNDED_RECTANGLE",
      description="PowerPoint shape: 'ROUNDED_RECTANGLE', 'RECTANGLE', 'RIGHT_ARROW', 'LINE', 'OVAL', 'TABLE'",
  )
  icon_symbol: Optional[str] = Field(
      default=None, description="Emoji or symbol e.g. ✅, ⚠️, 💡, ➔, ──►, 🔒, ⚡, 🏦, 🚫"
  )
  accent_bar_color_hex: Optional[str] = Field(
      default=None, description="Top accent bar color if applicable"
  )
  z_index: Optional[int] = Field(
      default=1,
      description="Z-index rendering order: 0=background, 1=container, 2=sub-card, 3=text/badge/arrow",
  )
  parent_id: Optional[str] = Field(
      default=None, description="ID of parent container if nested"
  )
  # Sub-card / Pipeline / List attributes
  sub_items: Optional[List[str]] = Field(
      default=None, description="Bullet items or nested points inside a card/sub-card"
  )
  callout_type: Optional[str] = Field(
      default=None, description="'SUMMARY', 'WARNING', 'NOTE', 'KEY_TAKEAWAY'"
  )
  # Table Matrix (if element_type == 'TABLE')
  table_matrix: Optional[List[List[str]]] = Field(
      default=None, description="2D matrix of table cell strings"
  )
  table_header_bg_hex: Optional[str] = Field(default="#F1F5F9")


class SlideNativeDecomposition(BaseModel):
  slide_background_hex: str = Field(
      default="#FFFFFF", description="Canvas background HEX color"
  )
  elements: List[SlideNativeElement] = Field(
      description="Ordered list of native shapes, containers, textboxes, and tables"
  )


from app.color_utils import (
    hex_to_pptx_color as _hex_to_pptx_color,
    hex_to_rgb_tuple as _hex_to_rgb_tuple,
)


def _find_free_port() -> int:
  """Finds an available TCP port dynamically."""
  with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
    s.bind(("", 0))
    s.listen(1)
    port = s.getsockname()[1]
    return port


def _get_font_path(font_filename: str) -> str:
  """Resolves the font path from package assets or local share."""
  p1 = os.path.join(FONTS_DIR, font_filename)
  if os.path.exists(p1):
    return p1
  p2 = os.path.join(FALLBACK_FONTS_DIR, font_filename)
  if os.path.exists(p2):
    return p2
  return p1


def _tag_slide_ids_in_html(html_str: str) -> tuple[str, List[str]]:
  """Ensures all slide elements have explicit unique id='slide-N' attributes."""
  # Match slide containers without falsely matching inner elements like slide-head, slide-title, slide-content
  pattern = re.compile(
      r'(<(?:div|section|article)[^>]*class=["\'][^"\']*(?<![\w-])(?:slide|slide-view|page)(?![\w-])[^"\']*["\'][^>]*)>',
      re.IGNORECASE,
  )
  slide_count = 0
  found_ids = []

  def repl(m):
    nonlocal slide_count
    slide_count += 1
    tag_open = m.group(1)
    id_match = re.search(r'id=["\']([^"\']+)["\']', tag_open, re.IGNORECASE)
    if id_match:
      sid = id_match.group(1)
    else:
      sid = f"slide-{slide_count}"
      tag_open = f'{tag_open} id="{sid}"'
    found_ids.append(sid)
    return f"{tag_open}>"

  tagged_html = pattern.sub(repl, html_str)
  return tagged_html, found_ids


def _extract_slide_ids(html_content: str) -> List[str]:
  """Dynamically extracts or generates slide IDs from HTML."""
  # 1. Tag and extract all authentic slide elements (div, section, article)
  _, auto_ids = _tag_slide_ids_in_html(html_content)
  if auto_ids:
    return list(dict.fromkeys(auto_ids))

  # 2. Look for id="slide-\d+"
  slide_ids = re.findall(r'id=["\'](slide-\d+)["\']', html_content, re.IGNORECASE)
  if slide_ids:
    return list(dict.fromkeys(slide_ids))

  # 3. Look for <section id="...">
  section_ids = re.findall(
      r'<section[^>]*id=["\']([^"\']+)["\']', html_content, re.IGNORECASE
  )
  if section_ids:
    return list(dict.fromkeys(section_ids))

  return ["slide-1"]


def _prepare_slide_html(
    raw_html: str, target_slide_id: str, slide_index: int, base_dir: str
) -> str:
  """Prepares clean, self-contained HTML for a single slide with local fonts and asset URLs."""
  html = raw_html

  # 0. Clean markdown code fences if wrapped
  if "```html" in html:
    html = html.split("```html", 1)[1].split("```", 1)[0].strip()
  elif "```" in html and "<html" in html:
    html = html.split("```", 1)[1].split("```", 1)[0].strip()

  # 1. Clean out ALL external blocking links and rewrite remote tailwind scripts to local asset
  html = re.sub(
      r'<link[^>]*href=["\']https?://[^"\']*["\'][^>]*>',
      "",
      html,
      flags=re.IGNORECASE,
  )
  html = re.sub(
      r'<link[^>]*rel=["\']preconnect["\'][^>]*>', "", html, flags=re.IGNORECASE
  )
  local_tailwind = os.path.abspath(
      os.path.join(os.path.dirname(__file__), "..", "assets", "tailwindcss.min.js")
  )
  if os.path.exists(local_tailwind):
    html = re.sub(
        r'<script[^>]*src=["\']https?://[^"\']*(?:tailwindcss|tailwind)[^"\']*["\'][^>]*>\s*</script>',
        f'<script src="file://{local_tailwind}"></script>',
        html,
        flags=re.IGNORECASE,
    )
  # Remove any remaining external http/https script tags to avoid network timeouts in headless Chrome
  html = re.sub(
      r'<script[^>]*src=["\']https?://[^"\']*["\'][^>]*>\s*</script>',
      "",
      html,
      flags=re.IGNORECASE,
  )

  # 2. Fix relative image and svg src paths to absolute file:// URLs
  def fix_src(match):
    src = match.group(1)
    if (
        not src.startswith("http://")
        and not src.startswith("https://")
        and not src.startswith("data:")
        and not src.startswith("file://")
    ):
      abs_path = os.path.abspath(os.path.join(base_dir, src))
      return f'src="file://{abs_path}"'
    return match.group(0)

  html = re.sub(r'src=["\']([^"\']+)["\']', fix_src, html)

  # 3. Resolve fonts
  regular_font = _get_font_path("NotoSansKR-Regular.otf")
  medium_font = _get_font_path("NotoSansKR-Medium.otf")
  bold_font = _get_font_path("NotoSansKR-Bold.otf")
  black_font = _get_font_path("NotoSansKR-Black.otf")

  # 4. Normalize slide IDs across diverse deck formats
  html, _ = _tag_slide_ids_in_html(html)

  # Detect multi-slide tabs or multiple slides targeting
  has_target_id = bool(
      re.search(rf'id=["\']{re.escape(target_slide_id)}["\']', html, re.IGNORECASE)
  )

  if has_target_id:
    # 1. Remove existing active class from all slide elements to prevent dual active rendering
    html = re.sub(
        r'(class=["\'][^"\']*\b)active(\b[^"\']*["\'])',
        r'\1\2',
        html,
    )
    # 2. Add active class to target_slide_id element so its intended CSS styles apply naturally
    def _add_active_to_target(match):
      full_tag = match.group(0)
      if "class=" in full_tag:
        return re.sub(r'class=["\']([^"\']*)["\']', r'class="\1 active"', full_tag)
      else:
        return full_tag.replace(f'id="{target_slide_id}"', f'id="{target_slide_id}" class="active"')

    html = re.sub(
        rf'<[^>]+id=["\']{re.escape(target_slide_id)}["\'][^>]*>',
        _add_active_to_target,
        html,
    )

    # 3. Synchronize data-active attribute for data-active driven presentations
    html = re.sub(r'data-active=["\'][^"\']*["\']', 'data-active="false"', html)
    html = re.sub(
        rf'(<[^>]+id=["\']{re.escape(target_slide_id)}["\'][^>]*?)data-active=["\'][^"\']*["\']',
        r'\1data-active="true"',
        html,
    )
    def _ensure_data_active(m):
      tag = m.group(0)
      if "data-active=" not in tag:
        return tag[:-1] + ' data-active="true">'
      return tag

    html = re.sub(
        rf'<[^>]+id=["\']{re.escape(target_slide_id)}["\'][^>]*>',
        _ensure_data_active,
        html,
    )

  multi_slide_css = ""
  if has_target_id:
    multi_slide_css = f"""
    section.slide:not(#{target_slide_id}),
    div.slide:not(#{target_slide_id}),
    .slide-view:not(#{target_slide_id}),
    .slide:not(#{target_slide_id}),
    .deck-container > section:not(#{target_slide_id}),
    .preview-container > .slide-view:not(#{target_slide_id}) {{
        display: none !important;
    }}
    #{target_slide_id},
    section#{target_slide_id},
    div#{target_slide_id},
    #{target_slide_id}[data-active="false"],
    #{target_slide_id}:not(.active) {{
        display: flex !important;
        visibility: visible !important;
        opacity: 1 !important;
        position: relative !important;
    }}
    #{target_slide_id} * {{
        visibility: visible !important;
    }}
    #{target_slide_id}.slide, #{target_slide_id}.slide-view {{
        display: flex !important;
    }}
    """

  font_face_css = ""
  if os.path.exists(regular_font):
    font_face_css = f"""
    @font-face {{
        font-family: 'Pretendard';
        font-weight: 400;
        src: local('Pretendard'), url('file://{regular_font}') format('opentype');
    }}
    @font-face {{
        font-family: 'Pretendard';
        font-weight: 700;
        src: local('Pretendard Bold'), url('file://{bold_font}') format('opentype');
    }}
    @font-face {{
        font-family: 'Noto Sans KR';
        font-weight: 400;
        src: local('Noto Sans KR'), url('file://{regular_font}') format('opentype');
    }}
    @font-face {{
        font-family: 'Noto Sans KR';
        font-weight: 500;
        src: local('Noto Sans KR Medium'), url('file://{medium_font}') format('opentype');
    }}
    @font-face {{
        font-family: 'Noto Sans KR';
        font-weight: 700;
        src: local('Noto Sans KR Bold'), url('file://{bold_font}') format('opentype');
    }}
    @font-face {{
        font-family: 'Noto Sans KR';
        font-weight: 900;
        src: local('Noto Sans KR Black'), url('file://{black_font}') format('opentype');
    }}
    """

  custom_css = f"""
    {font_face_css}

    *, *::before, *::after {{
        animation-duration: 0s !important;
        animation-delay: 0s !important;
        transition-duration: 0s !important;
        transition-delay: 0s !important;
    }}

    * {{ box-sizing: border-box; margin: 0; padding: 0; }}
    html, body {{
        width: 1920px !important;
        height: 1080px !important;
        margin: 0 !important;
        padding: 0 !important;
        overflow: hidden !important;
        font-family: 'Pretendard', 'Noto Sans KR', 'Noto Sans Korean', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif !important;
        -webkit-font-smoothing: antialiased;
    }}
    header:not(.slide-header), footer:not(.slide-footer), nav, .chrome, .tabs, .stage-controls, .controls-footer, #dots-container, #prev-btn, #next-btn, #slide-counter, #btn-ratio, #btn-notes, #counter {{
        display: none !important;
    }}
    main, .stage, .deck-container, .preview-container {{
        width: 1920px !important;
        height: 1080px !important;
        max-width: none !important;
        max-height: none !important;
        min-width: 1920px !important;
        min-height: 1080px !important;
        margin: 0 !important;
        padding: 0 !important;
        position: relative !important;
        display: block !important;
    }}
    .frame, #frame, .slide-frame {{
        width: 1920px !important;
        height: 1080px !important;
        max-width: 1920px !important;
        max-height: 1080px !important;
        min-width: 1920px !important;
        min-height: 1080px !important;
        margin: 0 !important;
        padding: 0 !important;
        border-radius: 0 !important;
        border: none !important;
        box-shadow: none !important;
        transform: none !important;
        position: absolute !important;
        top: 0 !important;
        left: 0 !important;
        display: block !important;
    }}
    .slide-canvas, .slide-container, #{target_slide_id} {{
        width: 1920px !important;
        height: 1080px !important;
        max-width: 1920px !important;
        max-height: 1080px !important;
        margin: 0 !important;
        border-radius: 0 !important;
        border: none !important;
        box-shadow: none !important;
        transform: none !important;
        box-sizing: border-box !important;
        position: absolute !important;
        top: 0 !important;
        left: 0 !important;
    }}
    {multi_slide_css}
    """

  # Inject custom CSS
  if "</head>" in html:
    html = html.replace("</head>", f"<style>{custom_css}</style>\n</head>")
  else:
    html = f"<head><style>{custom_css}</style></head>\n" + html

  # Override JS slide state if present (robust across let, const, var and naming conventions)
  html = re.sub(
      r"(let|var|const)\s+(currentSlide|currentSlideIndex|activeIndex|activeSlide|current|slideIdx|idx)\s*=\s*\d+;",
      r"\1 \2 = " + str(slide_index) + ";",
      html,
  )

  return html


JS_SLIDE_GEOMETRY_EXTRACTOR = """
(() => {
    const targetId = '%SLIDE_ID%';
    const s = document.getElementById(targetId) || document.querySelector('.slide[data-active="true"]') || document.querySelector('.slide.active') || document.querySelector('.slide');
    if (!s) return null;

    try {
        s.style.opacity = '1';
        s.style.visibility = 'visible';
        s.style.display = 'flex';
    } catch (e) {}

    const sRect = s.getBoundingClientRect();
    const slideW = sRect.width > 100 ? sRect.width : 1920;
    const slideH = sRect.height > 100 ? sRect.height : 1080;
    const scaleX = 13.333333 / slideW;
    const scaleY = 7.5 / slideH;

    const toIn = (r) => {
        const relLeft = r.left - sRect.left;
        const relTop = r.top - sRect.top;
        const inLeft = Math.max(0, relLeft * scaleX);
        const inTop = Math.max(0, relTop * scaleY);
        const inWidth = Math.min(13.333333 - inLeft, Math.max(0, r.width * scaleX));
        const inHeight = Math.min(7.5 - inTop, Math.max(0, r.height * scaleY));
        return {
            left: Number(inLeft.toFixed(3)),
            top: Number(inTop.toFixed(3)),
            width: Number(inWidth.toFixed(3)),
            height: Number(inHeight.toFixed(3)),
        };
    };

    const getStyles = (el) => {
        const cs = window.getComputedStyle(el);
        const radius = parseFloat(cs.borderRadius) || 0;
        const borderWidth = parseFloat(cs.borderWidth) || 0;
        const fontSize = parseFloat(cs.fontSize) || 14;
        const fontWeight = parseInt(cs.fontWeight) || (cs.fontWeight === 'bold' ? 700 : 400);
        return {
            color: cs.color,
            backgroundColor: cs.backgroundColor,
            fontSizePt: Number((fontSize * 0.75).toFixed(1)),
            fontWeight: fontWeight,
            isBold: fontWeight >= 600,
            borderRadius: radius,
            isRounded: radius >= 4,
            borderTopColor: cs.borderTopColor,
            borderTopWidth: parseFloat(cs.borderTopWidth) || 0,
            borderColor: cs.borderColor,
            borderWidth: borderWidth,
            textAlign: cs.textAlign,
            display: cs.display
        };
    };

    const sStyles = getStyles(s);
    const bodyStyles = getStyles(document.body);
    let slideBg = sStyles.backgroundColor;
    if (!slideBg || slideBg === 'transparent' || slideBg === 'rgba(0, 0, 0, 0)') {
        slideBg = bodyStyles.backgroundColor;
    }
    if (!slideBg || slideBg === 'transparent' || slideBg === 'rgba(0, 0, 0, 0)') {
        slideBg = 'rgb(255, 255, 255)';
    }

    // Universal Heading & Title Discovery
    const headings = Array.from(s.querySelectorAll('h1, h2, h3, h4, h5, h6')).filter(el => {
        if (el.closest('header') || el.closest('footer') || el.closest('.foot')) return false;
        const cs = window.getComputedStyle(el);
        return cs.display !== 'none' && cs.visibility !== 'hidden';
    });

    let mainH = headings.find(h => h.tagName.toLowerCase() === 'h1');
    if (!mainH && headings.length > 0) mainH = headings[0];
    if (!mainH) {
        const prominent = Array.from(s.querySelectorAll('div, p, span')).filter(el => {
            const cs = window.getComputedStyle(el);
            const fs = parseFloat(cs.fontSize) || 0;
            const fw = parseInt(cs.fontWeight) || 400;
            const r = el.getBoundingClientRect();
            return fs >= 22 && fw >= 600 && r.top < 400 && el.innerText.trim().length > 0;
        });
        if (prominent.length > 0) mainH = prominent[0];
    }

    // Extract ALL Header Badges / Pills / Eyebrows
    const headerBadges = [];
    const seenBadges = new Set();
    s.querySelectorAll('.slide-head .eyebrow, .slide-head .badge-pill, .slide-head [class*="tag"], .slide-head [class*="badge"], .slide-head [class*="pill"], .slide-tag, [class*="eyebrow"]').forEach(el => {
        const txt = el.innerText.trim();
        if (txt && !seenBadges.has(txt)) {
            seenBadges.add(txt);
            headerBadges.push({
                text: txt,
                rect: toIn(el.getBoundingClientRect()),
                styles: getStyles(el)
            });
        }
    });

    // Subtitle / Premise Discovery
    let subDesc = s.querySelector('p.subtitle, p.premise, p.lead, p.cover-subtitle, h2.subtitle, [class*="subtitle"], [class*="premise"], [class*="lead"], p.text-slate-400');
    if (!subDesc && mainH) {
        let next = mainH.closest('.slide-head') ? mainH.closest('.slide-head').nextElementSibling : mainH.nextElementSibling;
        while (next) {
            if (next.tagName.toLowerCase() === 'p') {
                subDesc = next;
                break;
            }
            if (next.querySelector('p')) {
                subDesc = next.querySelector('p');
                break;
            }
            next = next.nextElementSibling;
        }
    }

    const num = s.querySelector('.slide-number, [class*="slide-counter"], [class*="foot__n"]');

    // Descriptions: ONLY non-subtitle cover/intro paragraphs to prevent duplicate rendering
    const desc = [];
    s.querySelectorAll('p.cover-desc, p.intro-desc').forEach(el => {
        if (el !== subDesc) {
            const txt = el.innerText.trim();
            if (txt) {
                desc.push({
                    text: txt,
                    rect: toIn(el.getBoundingClientRect()),
                    styles: getStyles(el)
                });
            }
        }
    });

    // Universal Container Discovery: Panels, Columns, Cards, Pillars, Bridges
    const isTopContainer = (el) => {
        if (el === s || el.contains(s) || el.querySelector('table')) return false;
        if (el.closest('.slide-head') || el.closest('header') || el.closest('.foot') || el.closest('footer')) return false;
        if (mainH && (el === mainH || el.contains(mainH))) return false;
        if (subDesc && (el === subDesc || el.contains(subDesc))) return false;

        const cls = el.className || '';
        if (typeof cls !== 'string') return false;

        const r = el.getBoundingClientRect();
        if (r.width < 50 || r.height < 25 || r.width > slideW * 0.98 || r.height > slideH * 0.98) return false;

        const cs = window.getComputedStyle(el);
        const hasBorder = parseFloat(cs.borderWidth) > 0 || parseFloat(cs.borderTopWidth) > 0;
        const hasRadius = parseFloat(cs.borderRadius) >= 4;
        const hasBg = cs.backgroundColor !== 'rgba(0, 0, 0, 0)' && cs.backgroundColor !== 'transparent' && cs.backgroundColor !== slideBg;
        const hasShadow = cs.boxShadow && cs.boxShadow !== 'none';
        const hasCardClass = /\\b(panel|card|box|pillar|item|runtime|sub-card|column|col|bridge-badge|bridge-col|f-item|serving-box|attach-card|oss-card|template-box)\\b/i.test(cls);

        return (hasBorder || hasRadius || hasBg || hasShadow || hasCardClass);
    };

    let allCandidateContainers = Array.from(s.querySelectorAll('*')).filter(isTopContainer);
    let topContainers = allCandidateContainers.filter(c => !allCandidateContainers.some(p => p !== c && p.contains(c) && isTopContainer(p)));
    if (topContainers.length === 0) topContainers = allCandidateContainers;

    const cards = [];
    topContainers.forEach(c => {
        if (c.querySelector('table')) return;
        const r = toIn(c.getBoundingClientRect());
        const st = getStyles(c);

        // Check Bridge connector
        const bridgeBadgeEl = c.querySelector('.bridge-badge') || (c.classList.contains('bridge-badge') ? c : null);
        const isBridge = (c.className && c.className.includes('bridge')) || !!bridgeBadgeEl;
        let bridgeData = null;
        if (bridgeBadgeEl) {
            const bArrow = bridgeBadgeEl.querySelector('[class*="arrow"]');
            const bTxt = bridgeBadgeEl.querySelector('[class*="txt"]');
            bridgeData = {
                arrow: bArrow ? bArrow.innerText.trim() : '➔',
                text: bTxt ? bTxt.innerText.trim().replace(/\\n/g, ' ') : bridgeBadgeEl.innerText.trim().replace(/\\n/g, ' '),
                rect: toIn(bridgeBadgeEl.getBoundingClientRect()),
                styles: getStyles(bridgeBadgeEl)
            };
        }

        // Check Attach Card specialized structure
        const isAttachCard = (c.className && c.className.includes('attach-card')) || !!c.querySelector('.attach-card__mech');
        let attachData = null;
        if (isAttachCard) {
            const tgtEl = c.querySelector('.attach-card__target, [class*="target"]');
            const modEl = c.querySelector('.attach-card__mode, [class*="mode"]');
            const mchEl = c.querySelector('.attach-card__mech, [class*="mech"]');
            const pEl = c.querySelector('p');
            attachData = {
                target: tgtEl ? tgtEl.innerText.trim() : '',
                mode: modEl ? modEl.innerText.trim() : '',
                mech: mchEl ? mchEl.innerText.trim() : '',
                desc: pEl ? pEl.innerText.trim() : ''
            };
        }

        const tagEl = c.querySelector('.panel__tag, [class*="panel__tag"], span[class*="tag"], .ma-pillar__num');
        const tagData = tagEl ? {
            text: tagEl.innerText.trim(),
            rect: toIn(tagEl.getBoundingClientRect()),
            styles: getStyles(tagEl)
        } : null;

        let titleEl = c.querySelector('.panel__title, h2, h3, h4, [class*="title"], [class*="name"]');
        if (!titleEl && !isAttachCard) {
            const boldLeaves = Array.from(c.querySelectorAll('div, span, p')).filter(el => {
                if (el.children.length > 0) return false;
                const cs = window.getComputedStyle(el);
                const fw = parseInt(cs.fontWeight) || 400;
                const fs = parseFloat(cs.fontSize) || 0;
                return (fw >= 700 || fs >= 16) && el.innerText.trim().length > 0 && el.innerText.trim().length < 60;
            });
            if (boldLeaves.length > 0) titleEl = boldLeaves[0];
        }
        let cardTitle = titleEl ? titleEl.innerText.trim() : null;
        if (isAttachCard && attachData && attachData.target) {
            cardTitle = attachData.target;
        }
        const titleRect = titleEl ? toIn(titleEl.getBoundingClientRect()) : null;
        const titleStyles = titleEl ? getStyles(titleEl) : null;

        const descEl = c.querySelector('.panel__desc, [class*="desc"], .oss-card__pos');
        let cardDesc = descEl && descEl !== titleEl ? {
            text: descEl.innerText.trim(),
            rect: toIn(descEl.getBoundingClientRect()),
            styles: getStyles(descEl)
        } : null;
        if (isAttachCard && attachData && attachData.desc) {
            cardDesc = {
                text: attachData.desc,
                rect: r,
                styles: st
            };
        }

        // Pipeline steps
        const pipeStepEls = Array.from(c.querySelectorAll('.ma-pipe-step, [class*="pipe-step"], [class*="step-item"]'));
        const pipeSteps = pipeStepEls.map(st => {
            const h = st.querySelector('h3, h4, h5, [class*="title"], strong');
            const p = st.querySelector('p, span:last-child');
            return {
                title: h ? h.innerText.trim() : st.innerText.trim(),
                desc: p && p !== h ? p.innerText.trim() : '',
                rect: toIn(st.getBoundingClientRect()),
                styles: getStyles(st)
            };
        });

        // Sub-cards
        const subCardEls = Array.from(c.querySelectorAll('.sub-card, .runtime-card, [class*="sub-card"], [class*="runtime-card"]')).filter(sc => {
            const cls = sc.className || '';
            if (typeof cls !== 'string') return false;
            return !cls.includes('__head') && !cls.includes('__badge') && !cls.includes('__mech') && !cls.includes('__target') && !cls.includes('__mode') && !cls.includes('-grid') && !cls.includes('-layout');
        });
        const subCards = subCardEls.map(sc => {
            const scHead = sc.querySelector('.sub-card__head, .runtime-card__head, [class*="head"], h4, h5');
            const scBadge = sc.querySelector('.runtime-card__badge, [class*="badge"], [class*="mode"]');
            const headTxt = scHead ? scHead.innerText.trim() : '';
            const badgeTxt = scBadge ? scBadge.innerText.trim() : '';
            const items = Array.from(sc.querySelectorAll('li, p')).filter(el => {
                if (scHead && (el === scHead || scHead.contains(el))) return false;
                if (scBadge && (el === scBadge || scBadge.contains(el))) return false;
                return true;
            }).map(el => el.innerText.trim()).filter(Boolean);
            return {
                rect: toIn(sc.getBoundingClientRect()),
                head: headTxt,
                badge: badgeTxt,
                items: items,
                styles: getStyles(sc)
            };
        });

        // Inline flows
        const flowEls = Array.from(c.querySelectorAll('.inline-flow, .flow, [class*="inline-flow"]'));
        const flows = flowEls.map(fl => {
            const nodes = Array.from(fl.children).map(ch => ({
                text: ch.innerText.trim(),
                isArrow: ch.innerText.includes('──►') || ch.innerText.includes('➔') || ch.innerText.includes('→'),
                isHighlight: (ch.className || '').includes('hl') || (ch.className || '').includes('node--hl'),
                rect: toIn(ch.getBoundingClientRect()),
                styles: getStyles(ch)
            }));
            return { rect: toIn(fl.getBoundingClientRect()), nodes: nodes };
        });

        // Feature items
        const fItemEls = Array.from(c.querySelectorAll('.f-item, [class*="f-item"]')).filter(fi => {
            const cls = fi.className || '';
            return typeof cls === 'string' && !cls.includes('__icon') && !cls.includes('__text') && !cls.includes('-list');
        });
        const fItems = fItemEls.map(fi => {
            const iconEl = fi.querySelector('.f-item__icon, [class*="icon"], span:first-child');
            const textEl = fi.querySelector('div, p') || (iconEl ? Array.from(fi.children).find(ch => ch !== iconEl) : fi);
            return {
                icon: iconEl ? iconEl.innerText.trim() : '',
                iconRect: iconEl ? toIn(iconEl.getBoundingClientRect()) : null,
                iconStyles: iconEl ? getStyles(iconEl) : null,
                text: textEl ? textEl.innerText.trim() : (iconEl ? fi.innerText.replace(iconEl.innerText, '').trim() : fi.innerText.trim()),
                textRect: textEl ? toIn(textEl.getBoundingClientRect()) : toIn(fi.getBoundingClientRect()),
                textStyles: textEl ? getStyles(textEl) : getStyles(fi),
                rect: toIn(fi.getBoundingClientRect())
            };
        });

        const seenTexts = new Set();
        if (cardTitle) seenTexts.add(cardTitle);
        if (tagData) seenTexts.add(tagData.text);
        if (cardDesc) seenTexts.add(cardDesc.text);
        if (attachData) {
            if (attachData.target) seenTexts.add(attachData.target);
            if (attachData.mode) seenTexts.add(attachData.mode);
            if (attachData.mech) seenTexts.add(attachData.mech);
            if (attachData.desc) seenTexts.add(attachData.desc);
            const ah = c.querySelector('.attach-card__head');
            if (ah) seenTexts.add(ah.innerText.trim());
        }
        pipeSteps.forEach(ps => { if (ps.title) seenTexts.add(ps.title); if (ps.desc) seenTexts.add(ps.desc); });
        subCards.forEach(sc => {
            if (sc.head) seenTexts.add(sc.head);
            if (sc.badge) seenTexts.add(sc.badge);
            sc.items.forEach(it => seenTexts.add(it));
        });
        fItems.forEach(fi => {
            if (fi.icon) seenTexts.add(fi.icon);
            if (fi.text) seenTexts.add(fi.text);
        });

        // Block-level code and example blocks
        const isBlockCode = (cb) => {
            if (cb.tagName.toLowerCase() === 'pre') return true;
            const cls = cb.className || '';
            if (typeof cls === 'string' && (cls.includes('__code') || cls.includes('__ex') || cls.includes('code-block') || cls.includes('spec') || cls.includes('template-box__code'))) return true;
            const cs = window.getComputedStyle(cb);
            const isMono = cs.fontFamily.includes('mono') || cls.includes('mono');
            const hasMultipleLines = cb.innerText.trim().includes('\\n') || cb.offsetHeight >= 36;
            return isMono && hasMultipleLines;
        };

        const codeBlocks = Array.from(c.querySelectorAll('pre, [class*="code"], [class*="spec"], [class*="__ex"]')).filter(cb => {
            if (cb === c || (titleEl && titleEl.contains(cb)) || (tagEl && tagEl.contains(cb))) return false;
            return isBlockCode(cb) && cb.innerText.trim().length > 0;
        }).map(cb => ({
            text: cb.innerText.trim(),
            rect: toIn(cb.getBoundingClientRect()),
            styles: getStyles(cb)
        }));
        codeBlocks.forEach(cb => seenTexts.add(cb.text));

        // Paragraphs / list items
        const paras = [];
        if (!isAttachCard) {
            const directLis = Array.from(c.querySelectorAll('li')).filter(li => !subCardEls.some(sc => sc.contains(li)));
            if (directLis.length > 0) {
                directLis.forEach(li => {
                    const txt = li.innerText.trim();
                    if (txt && !seenTexts.has(txt)) {
                        seenTexts.add(txt);
                        paras.push({ text: txt, rect: toIn(li.getBoundingClientRect()), styles: getStyles(li) });
                    }
                });
            }
            c.querySelectorAll('p, div').forEach(el => {
                if (el === titleEl || (titleEl && titleEl.contains(el))) return;
                if (descEl && (el === descEl || descEl.contains(el))) return;
                if (tagEl && (el === tagEl || tagEl.contains(el))) return;
                if (subCardEls.some(sc => sc.contains(el))) return;
                if (flowEls.some(fl => fl.contains(el))) return;
                if (fItemEls.some(fi => fi.contains(el))) return;
                if (pipeStepEls.some(ps => ps.contains(el))) return;
                if (codeBlocks.some(cb => cb.text === el.innerText.trim())) return;
                if (el.querySelector('p, div, table, pre')) return;
                const txt = el.innerText.trim();
                if (txt && !seenTexts.has(txt)) {
                    seenTexts.add(txt);
                    paras.push({ text: txt, rect: toIn(el.getBoundingClientRect()), styles: getStyles(el) });
                }
            });
        }

        const badges = Array.from(c.querySelectorAll('span.badge, [class*="badge"], [class*="tag"], [class*="lang"]')).filter(b => {
            if (b === tagEl || (tagEl && tagEl.contains(b))) return false;
            if (subCardEls.some(sc => sc.contains(b))) return false;
            if (isAttachCard) return false;
            return true;
        }).map(b => ({
            text: b.innerText.trim(),
            rect: toIn(b.getBoundingClientRect()),
            styles: getStyles(b)
        }));

        cards.push({
            rect: r,
            styles: st,
            isRounded: st.isRounded,
            borderRadius: st.borderRadius,
            hasTopAccent: st.borderTopWidth >= 4,
            topAccentColor: st.borderTopColor,
            isBridge: isBridge,
            bridge: bridgeData,
            isAttachCard: isAttachCard,
            attachData: attachData,
            tag: tagData,
            title: cardTitle,
            titleRect: titleRect,
            titleStyles: titleStyles,
            desc: cardDesc,
            pipelineSteps: pipeSteps,
            subCards: subCards,
            flows: flows,
            fItems: fItems,
            badges: badges,
            codeBlocks: codeBlocks,
            paragraphs: paras
        });
    });

    // Tables
    const tables = [];
    s.querySelectorAll('table').forEach(tbl => {
        const parentCard = tbl.closest('.info-card, .card, .panel, [class*="rounded"]');
        const cardTitle = parentCard && parentCard.querySelector('h3, h4, [class*="font-bold"], [class*="title"]') ? parentCard.querySelector('h3, h4, [class*="font-bold"], [class*="title"]').innerText.trim() : '';
        const tRect = toIn(tbl.getBoundingClientRect());
        const headers = [];
        tbl.querySelectorAll('thead th, tr:first-child th').forEach(th => {
            const st = getStyles(th);
            headers.push({
                text: th.innerText.trim(),
                color: st.color,
                bgColor: st.backgroundColor !== 'transparent' && st.backgroundColor !== 'rgba(0, 0, 0, 0)' ? st.backgroundColor : null,
                fontWeight: st.fontWeight
            });
        });
        const rows = [];
        tbl.querySelectorAll('tbody tr, tr:not(:first-child)').forEach(tr => {
            if (tr.querySelector('th') && !tr.querySelector('td') && tr.parentElement.tagName === 'THEAD') return;
            const trSt = getStyles(tr);
            const rowCells = [];
            tr.querySelectorAll('td, th').forEach(td => {
                const mb = td.querySelector('.math-badge, [class*="badge"]');
                const st = getStyles(td);
                let cellBg = st.backgroundColor;
                if (!cellBg || cellBg === 'transparent' || cellBg === 'rgba(0, 0, 0, 0)') {
                    cellBg = trSt.backgroundColor;
                }
                if (cellBg === 'transparent' || cellBg === 'rgba(0, 0, 0, 0)') {
                    cellBg = null;
                }
                rowCells.push({
                    text: td.innerText.trim(),
                    color: st.color,
                    bgColor: cellBg,
                    is_bold: parseInt(st.fontWeight) >= 600,
                    has_badge: !!mb,
                    badge_text: mb ? mb.innerText.trim() : null
                });
            });
            if (rowCells.length) rows.push(rowCells);
        });
        tables.push({
            rect: tRect,
            cardTitle: cardTitle,
            inCard: !!parentCard,
            headers: headers,
            rows: rows,
            num_cols: Math.max(headers.length, ...rows.map(r => r.length), 0),
            num_rows: (headers.length ? 1 : 0) + rows.length
        });
    });

    // Highlight / Callout boxes (strictly exclude wrappers like div.foot that contain .foot__pocket)
    const hlBoxes = [];
    const seenHl = new Set();
    s.querySelectorAll('.foot__pocket, .highlight-box, [class*="highlight-box"], [class*="callout"]').forEach(hl => {
        const txt = hl.innerText ? hl.innerText.trim() : '';
        if (!txt || seenHl.has(txt)) return;
        seenHl.add(txt);
        hlBoxes.push({
            rect: toIn(hl.getBoundingClientRect()),
            styles: getStyles(hl),
            text: txt
        });
    });
    if (hlBoxes.length === 0) {
        s.querySelectorAll('*').forEach(hl => {
            if (hl.closest('table') || hl.querySelector('table')) return;
            if (hl.querySelector('.foot__pocket, [class*="highlight-box"], [class*="callout"]')) return;
            const txt = hl.innerText ? hl.innerText.trim() : '';
            if (!txt || seenHl.has(txt)) return;
            const isCallout = txt.startsWith('💡') || txt.startsWith('📌') || txt.startsWith('★') || 
                              txt.startsWith('⚠️') || txt.startsWith('ℹ️') || txt.includes('핵심 메시지') ||
                              txt.includes('핵심 요약') || txt.includes('참고') || txt.includes('결론');
            if (isCallout && hl.offsetHeight < 320 && hl.offsetWidth > 150) {
                seenHl.add(txt);
                hlBoxes.push({
                    rect: toIn(hl.getBoundingClientRect()),
                    styles: getStyles(hl),
                    text: txt
                });
            }
        });
    }

    // Footnotes / Bottom notes (strictly leaf elements)
    const footnotes = [];
    const seenFn = new Set();
    s.querySelectorAll('.foot__n, .footnote, [class*="footnote"], [class*="bottom-note"]').forEach(el => {
        const txt = el.innerText ? el.innerText.trim() : '';
        if (!txt || seenFn.has(txt)) return;
        seenFn.add(txt);
        footnotes.push({
            rect: toIn(el.getBoundingClientRect()),
            styles: getStyles(el),
            text: txt
        });
    });
    if (footnotes.length === 0) {
        s.querySelectorAll('p, div, span, small').forEach(el => {
            if (el.querySelector('p, div, table')) return;
            const txt = el.innerText ? el.innerText.trim() : '';
            if (!txt || seenFn.has(txt)) return;
            const isFootnote = txt.startsWith('*') || txt.startsWith('※') || txt.startsWith('†') || 
                               txt.startsWith('1)') || txt.startsWith('참조') || txt.startsWith('출처');
            const r = toIn(el.getBoundingClientRect());
            if (isFootnote && r.top > 3.5) {
                seenFn.add(txt);
                footnotes.push({
                    rect: r,
                    styles: getStyles(el),
                    text: txt
                });
            }
        });
    }

    return {
        id: targetId,
        slideBgColor: slideBg,
        headerBadges: headerBadges,
        tag: headerBadges.length > 0 ? headerBadges[0] : null,
        num: num ? { text: num.innerText.trim(), rect: toIn(num.getBoundingClientRect()), styles: getStyles(num) } : null,
        title: mainH ? { text: mainH.innerText.trim(), rect: toIn(mainH.getBoundingClientRect()), styles: getStyles(mainH) } : null,
        sub: subDesc ? { text: subDesc.innerText.trim(), rect: toIn(subDesc.getBoundingClientRect()), styles: getStyles(subDesc) } : null,
        desc: desc,
        cards: cards,
        tables: tables,
        highlightBoxes: hlBoxes,
        footnotes: footnotes
    };
})()
"""


async def _capture_slides_cdp(
    slide_items: List[Any],
    port: int,
    profile_dir: str,
    output_dir: str,
    scale_factor: int = 2,
) -> tuple[List[str], List[Optional[Dict[str, Any]]]]:
  """Connects to headless Chrome via CDP and captures high-res PNGs and exact DOM layout geometries."""
  captured_paths = []
  captured_geometries = []

  # Wait for CDP to respond
  for _ in range(30):
    try:
      with urllib.request.urlopen(
          f"http://127.0.0.1:{port}/json/version", timeout=1
      ) as resp:
        break
    except Exception:
      await asyncio.sleep(0.3)

  # Create a new tab
  req = urllib.request.Request(f"http://127.0.0.1:{port}/json/new", method="PUT")
  with urllib.request.urlopen(req, timeout=5) as resp:
    tab_info = json.loads(resp.read().decode())

  ws_url = tab_info["webSocketDebuggerUrl"]
  target_id = tab_info["id"]

  async with websockets.connect(ws_url, max_size=50 * 1024 * 1024) as ws:
    msg_id = 0

    async def send_recv(method, params=None, timeout=12.0):
      nonlocal msg_id
      msg_id += 1
      curr_id = msg_id
      payload = {"id": curr_id, "method": method}
      if params:
        payload["params"] = params
      await ws.send(json.dumps(payload))
      start_t = time.time()
      while True:
        elapsed = time.time() - start_t
        remaining = max(1.0, timeout - elapsed)
        try:
          raw = await asyncio.wait_for(ws.recv(), timeout=remaining)
        except asyncio.TimeoutError:
          print(f"[CDP Warning] Timeout waiting for {method} (id={curr_id})")
          return {}
        msg = json.loads(raw)
        if msg.get("id") == curr_id:
          if "error" in msg:
            print(f"[CDP Error in {method}]: {msg['error']}")
            return {}
          return msg.get("result", {})

    await send_recv("Page.enable")
    await send_recv(
        "Emulation.setDeviceMetricsOverride",
        {
            "width": 1920,
            "height": 1080,
            "deviceScaleFactor": scale_factor,
            "mobile": False,
        },
    )

    for idx, item in enumerate(slide_items, start=1):
      if isinstance(item, (tuple, list)):
        sid, slide_file = item[0], item[1]
      else:
        sid, slide_file = f"slide-{idx}", item

      file_url = f"file://{os.path.abspath(slide_file)}"
      await send_recv("Page.navigate", {"url": file_url}, timeout=8.0)

      # Wait for DOM to be interactive / complete
      try:
        await send_recv(
            "Runtime.evaluate",
            {
                "expression": (
                    "new Promise(r => {"
                    "  if (document.readyState === 'complete' || document.readyState === 'interactive') r(true);"
                    "  else window.addEventListener('DOMContentLoaded', () => r(true), {once: true});"
                    "  setTimeout(() => r(false), 1200);"
                    "})"
                ),
                "awaitPromise": True,
                "returnByValue": True,
            },
            timeout=3.0,
        )
      except Exception:
        pass

      # Wait for local font files (Pretendard/Noto Sans KR) with bounded race
      try:
        await send_recv(
            "Runtime.evaluate",
            {
                "expression": (
                    "Promise.race(["
                    "  document.fonts ? document.fonts.ready : Promise.resolve(),"
                    "  new Promise(r => setTimeout(r, 350))"
                    "])"
                ),
                "awaitPromise": True,
                "returnByValue": True,
            },
            timeout=2.0,
        )
      except Exception:
        await asyncio.sleep(0.2)

      # Brief stabilization buffer for DOM reflow
      await asyncio.sleep(0.1)

      # Extract exact browser-measured DOM geometry directly from rendered layout
      geom = None
      try:
        js = JS_SLIDE_GEOMETRY_EXTRACTOR.replace("%SLIDE_ID%", str(sid))
        layout_res = await send_recv("Runtime.evaluate", {"expression": js, "returnByValue": True})
        geom = layout_res.get("result", {}).get("value")
      except Exception as geom_err:
        print(f"[CDP Layout Warning for {sid}]: {geom_err}")
      captured_geometries.append(geom)

      res = await send_recv("Page.captureScreenshot", {"format": "png"})
      img_bytes = base64.b64decode(res["data"])

      out_png = os.path.join(output_dir, f"slide-{idx}.png")
      with open(out_png, "wb") as f_out:
        f_out.write(img_bytes)
      captured_paths.append(out_png)

  try:
    urllib.request.urlopen(f"http://127.0.0.1:{port}/json/close/{target_id}")
  except Exception:
    pass

  return captured_paths, captured_geometries


def _find_chrome_binary(custom_path: Optional[str] = None) -> str:
  """Locates the available Chrome or Chromium binary executable."""
  candidates = [
      custom_path,
      os.environ.get("CHROME_BIN"),
      "/usr/bin/google-chrome-stable",
      "/opt/google/chrome/chrome",
      "/usr/bin/google-chrome",
      "/usr/bin/chromium-browser",
      "/usr/bin/chromium",
      "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
  ]
  for c in candidates:
    if c and os.path.exists(c):
      return c
  return "/usr/bin/google-chrome-stable"


from app.genai_client import DEFAULT_GEMINI_MODEL, get_genai_client


def decompose_slide_image_with_vision(
    image_path: str,
    model_name: str = DEFAULT_GEMINI_MODEL,
) -> Optional[SlideNativeDecomposition]:
  """Decomposes a rendered 16:9 slide screenshot into 100% native slide elements using Gemini Vision."""
  try:
    from google import genai
    from google.genai import types

    with open(image_path, "rb") as f:
      image_bytes = f.read()

    client = get_genai_client()

    prompt = """
    You are an expert Vision-to-Native-Slide Reverse Engineer.
    Analyze this 16:9 widescreen slide image and decompose it into 100% editable native PowerPoint slide elements.

    [STRICT MANDATES]
    1. ZERO-OMISSION OCR: Transcribe 100% of all text, titles, subtitles, bullets, table cells, metric numbers, badges, and footnotes verbatim without dropping or summarizing anything.
    2. NORMALIZED COORDINATES: Provide exact bounding box percentages (left_pct, top_pct, width_pct, height_pct from 0.0 to 100.0) for each element so they align precisely with the screenshot visual layout.
    3. ELEMENT TYPES:
       - 'CONTAINER': Background cards, rounded surface tiles, banners. (Specify bg_color_hex, border_color_hex, border_radius_pt).
       - 'TEXT_BOX': Independent titles, paragraphs, bullet lists, metric callouts. (Specify font_size_pt, is_bold, font_color_hex, alignment).
       - 'TABLE': Tabular matrices (provide table_matrix as 2D list of row strings, table_header_bg_hex).
       - 'DIVIDER': Thin horizontal/vertical separating lines.
       - 'BADGE': Small category tags/pills.
    4. DESIGN TOKENS: Extract exact HEX colors for background canvas, text colors, container surfaces, and border strokes.
    """

    response = client.models.generate_content(
        model=model_name,
        contents=[
            types.Part.from_bytes(data=image_bytes, mime_type="image/png"),
            prompt,
        ],
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=SlideNativeDecomposition,
            temperature=0.1,
        ),
    )

    if response.text:
      return SlideNativeDecomposition.model_validate_json(response.text)
  except Exception as e:
    print(f"[Vision Decomposition Error] {type(e).__name__}: {e}")
    return None


def extract_root_css_vars(html_content: str) -> Dict[str, str]:
  """Extracts all :root CSS custom properties (--var: val) from HTML."""
  var_map = {}
  matches = re.findall(r"(--[a-zA-Z0-9_-]+)\s*:\s*([^;}\n]+)", html_content)
  for k, v in matches:
    var_map[k.strip()] = v.strip()
  return var_map


def parse_color_value(
    val: str,
    default: RGBColor = RGBColor(17, 17, 21),
    bg_blend_rgb: tuple[int, int, int] = (255, 255, 255),
) -> RGBColor:
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
  m_rgba = re.match(r"rgba?\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)(?:\s*,\s*([0-9.]+))?\s*\)", val)
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


def resolve_style_color(style_str: str, css_vars_raw: Dict[str, str], default: RGBColor = RGBColor(17, 17, 21)) -> RGBColor:
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


def parse_slide_semantic_data(slide_elem, css_vars_raw: Dict[str, str]) -> Dict[str, Any]:
  """Performs deep semantic DOM extraction of slide headers, tables, cell alert colors, badges, and cards."""
  slide_id = slide_elem.get("id", "slide-unknown")
  tag_nodes = slide_elem.xpath('.//*[contains(@class, "slide-tag") or contains(@class, "tag") or contains(@class, "badge") or contains(@class, "eyebrow") or contains(@class, "pill")]')
  num_nodes = slide_elem.xpath('.//*[contains(@class, "slide-number") or contains(@class, "slide-counter") or contains(@class, "foot__n")]')
  tag = tag_nodes[0].text_content().strip() if tag_nodes else ""
  number = num_nodes[0].text_content().strip() if num_nodes else ""

  h_nodes = slide_elem.xpath('.//h1 | .//h2')
  sub_nodes = slide_elem.xpath('.//p[contains(@class, "subtitle") or contains(@class, "premise") or contains(@class, "lead")]')
  title = h_nodes[0].text_content().strip() if h_nodes else ""
  subtitle = sub_nodes[0].text_content().strip() if sub_nodes else ""

  # Tables with cell-level style and run parsing
  tables_data = []
  for tbl in slide_elem.xpath('.//table'):
    headers = [th.text_content().strip() for th in tbl.xpath('.//thead//th | .//tr[1]//th')]
    rows = []
    for tr in tbl.xpath('.//tbody//tr | .//tr[position() > 1]'):
      row_cells = []
      for td in tr.xpath('.//td | .//th'):
        style = td.get("style", "")
        text = td.text_content().strip()
        color_rgb = resolve_style_color(style, css_vars_raw)
        is_parent_bold = "font-weight: 700" in style or "font-weight: 800" in style
        has_badge = bool(td.xpath('.//span[contains(@class, "math-badge")]'))
        badge_text = td.xpath('.//span[contains(@class, "math-badge")]/text()')
        badge_str = badge_text[0].strip() if badge_text else ""

        child_nodes = td.xpath('child::node()')
        runs = []
        if any(getattr(n, "tag", None) in ("b", "strong", "span") for n in child_nodes):
          for n in child_nodes:
            if isinstance(n, str):
              if n:
                runs.append({"text": n, "bold": is_parent_bold})
            elif getattr(n, "tag", None) in ("b", "strong"):
              runs.append({"text": n.text_content(), "bold": True})
            else:
              runs.append({"text": n.text_content(), "bold": is_parent_bold})

        hex_color = f"#{color_rgb[0]:02X}{color_rgb[1]:02X}{color_rgb[2]:02X}"
        row_cells.append({
            "text": text,
            "color_hex": hex_color,
            "color_rgb": (color_rgb[0], color_rgb[1], color_rgb[2]),
            "is_bold": is_parent_bold or bool(td.xpath('.//b | .//strong')),
            "has_badge": has_badge,
            "badge_text": badge_str,
            "runs": runs,
        })
      if row_cells:
        rows.append(row_cells)

    parent_card = tbl.xpath('ancestor::div[contains(@class, "info-card") or contains(@class, "card") or contains(@class, "panel")]')
    card_title_nodes = parent_card[0].xpath('.//h3/text() | .//h4/text()') if parent_card else []
    card_title = card_title_nodes[0].strip() if card_title_nodes else ""

    tables_data.append({
        "headers": headers,
        "rows": rows,
        "num_cols": max(len(headers), max((len(r) for r in rows), default=0)),
        "num_rows": (1 if headers else 0) + len(rows),
        "in_card": bool(parent_card),
        "card_title": card_title,
    })

  # Cards / Visual Containers
  cards_data = []
  card_candidates = slide_elem.xpath(
      './/*[contains(@class, "info-card") or contains(@class, "card") or contains(@class, "panel") or contains(@class, "pillar") or contains(@class, "box") or contains(@class, "runtime") or contains(@class, "attach-card") or contains(@class, "oss-card")]'
  )
  card_elems = [c for c in card_candidates if not any(p != c and c in p.iterdescendants() for p in card_candidates)]
  for c in card_elems:
    if c.xpath('.//table'):
      continue
    c_headings = c.xpath('.//h1 | .//h2 | .//h3 | .//h4 | .//*[contains(@class, "title") or contains(@class, "head") or contains(@class, "name")]')
    c_title = c_headings[0].text_content().strip() if c_headings else ""
    paras = [p.text_content().strip() for p in c.xpath('.//p | .//li | .//*[contains(@class, "f-item")]') if p.text_content().strip()]
    badges = [b.text_content().strip() for b in c.xpath('.//*[contains(@class, "badge") or contains(@class, "tag") or contains(@class, "lang")]')]
    cls = c.get("class", "")
    accent = "purple" if "purple-accent" in cls else ("blue" if "blue-accent" in cls else ("green" if "green-accent" in cls else ("red" if "red-accent" in cls else "none")))
    cards_data.append({
        "title": c_title,
        "paragraphs": paras,
        "badges": badges,
        "accent": accent,
    })

  # Highlight / Callout boxes
  hl_boxes = []
  for hl in slide_elem.xpath('.//*[contains(@class, "highlight-box") or contains(@class, "foot__pocket") or contains(@class, "callout") or contains(@class, "alert")]'):
    txt = hl.text_content().strip()
    if txt:
      hl_boxes.append({
          "title": txt[:40],
          "body": txt,
          "color": "blue",
      })

  # Math Badges
  math_badges = [b.text_content().strip() for b in slide_elem.xpath('.//span[contains(@class, "math-badge")]')]
  raw_html = lxml.html.tostring(slide_elem, encoding="unicode", pretty_print=True)

  return {
      "slide_id": slide_id,
      "tag": tag,
      "number": number,
      "title": title,
      "subtitle": subtitle,
      "tables": tables_data,
      "cards": cards_data,
      "highlight_boxes": hl_boxes,
      "math_badges": math_badges,
      "raw_html": raw_html,
  }


def build_styled_native_table(
    slide,
    table_data: Dict[str, Any],
    left: Inches,
    top: Inches,
    width: Inches,
    height: Inches,
    card_border_rgb: Optional[RGBColor] = None,
    font_name: str = "Pretendard",
    col_ratios: Optional[List[float]] = None,
):
  """Builds a pixel-faithful native PowerPoint table with theme-adaptive styling, exact row colors, and contrast preservation."""
  is_dark = table_data.get("is_dark", False)
  slide_bg = table_data.get("slide_bg_rgb", (15, 23, 42) if is_dark else (255, 255, 255))
  blend_rgb = (slide_bg[0], slide_bg[1], slide_bg[2])

  if card_border_rgb is None:
    card_border_rgb = RGBColor(51, 65, 85) if is_dark else RGBColor(233, 219, 237)

  # Guarantee table does not overflow slide boundaries
  if left < Inches(0.4):
    left = Inches(0.4)
  max_avail_w = Inches(13.333) - left - Inches(0.4)
  if width > max_avail_w:
    width = max_avail_w

  if table_data.get("in_card") or table_data.get("inCard"):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card_bg = table_data.get("card_bg_rgb") or (RGBColor(30, 41, 59) if is_dark else RGBColor(255, 255, 255))
    card.fill.fore_color.rgb = card_bg
    card.line.color.rgb = card_border_rgb
    card.line.width = Pt(1.0)

  card_title = table_data.get("card_title")
  if card_title:
    tx = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(0.12), width - Inches(0.4), Inches(0.38))
    p = tx.text_frame.paragraphs[0]
    p.text = card_title
    p.font.name = font_name
    p.font.size = Pt(13.0)
    p.font.bold = True
    p.font.color.rgb = RGBColor(147, 197, 253) if is_dark else RGBColor(37, 99, 235)

  t_top = top + (Inches(0.52) if card_title else Inches(0.12))
  t_height = height - (Inches(0.64) if card_title else Inches(0.24))
  t_left = left + Inches(0.12)
  t_width = width - Inches(0.24)

  num_rows = table_data["num_rows"]
  num_cols = table_data["num_cols"]
  table_shape = slide.shapes.add_table(num_rows, num_cols, t_left, t_top, t_width, t_height)
  tbl = table_shape.table

  if col_ratios and len(col_ratios) == num_cols:
    total_r = sum(col_ratios)
    for c_idx, r in enumerate(col_ratios):
      tbl.columns[c_idx].width = int(t_width * (r / total_r))
  else:
    eq_w = int(t_width / num_cols)
    for c_idx in range(num_cols):
      tbl.columns[c_idx].width = eq_w

  if table_data.get("headers"):
    for c_idx, h in enumerate(table_data["headers"]):
      h_text = h["text"] if isinstance(h, dict) else str(h)
      h_bg = h.get("bgColor") if isinstance(h, dict) else None
      h_col = h.get("color") if isinstance(h, dict) else None

      if h_bg:
        header_bg_rgb = parse_color_value(h_bg, default=RGBColor(15, 23, 42) if is_dark else RGBColor(241, 245, 249), bg_blend_rgb=blend_rgb)
      elif is_dark:
        header_bg_rgb = RGBColor(15, 23, 42)
      else:
        header_bg_rgb = RGBColor(241, 245, 249)

      if h_col:
        header_fg_rgb = parse_color_value(h_col, default=RGBColor(148, 163, 184) if is_dark else RGBColor(30, 41, 59))
      elif is_dark:
        header_fg_rgb = RGBColor(148, 163, 184)
      else:
        header_fg_rgb = RGBColor(30, 41, 59)

      cell = tbl.cell(0, c_idx)
      cell.text = h_text
      cell.fill.solid()
      cell.fill.fore_color.rgb = header_bg_rgb
      cell.vertical_anchor = MSO_ANCHOR.MIDDLE
      cell.margin_left = Inches(0.12)
      cell.margin_right = Inches(0.12)
      cell.margin_top = Inches(0.08)
      cell.margin_bottom = Inches(0.08)
      p = cell.text_frame.paragraphs[0]
      p.font.name = font_name
      p.font.size = Pt(11.0)
      p.font.bold = True
      p.font.color.rgb = header_fg_rgb

  for r_idx, row in enumerate(table_data["rows"]):
    row_num = r_idx + (1 if table_data.get("headers") else 0)
    for c_idx, cell_data in enumerate(row):
      if c_idx >= num_cols:
        break
      cell = tbl.cell(row_num, c_idx)
      cell.vertical_anchor = MSO_ANCHOR.MIDDLE
      cell.margin_left = Inches(0.12)
      cell.margin_right = Inches(0.12)
      cell.margin_top = Inches(0.08)
      cell.margin_bottom = Inches(0.08)
      cell.fill.solid()

      # Dynamic cell background resolution
      cell_bg = cell_data.get("bgColor")
      if cell_bg:
        row_bg_rgb = parse_color_value(cell_bg, default=RGBColor(30, 41, 59) if is_dark else RGBColor(255, 255, 255), bg_blend_rgb=blend_rgb)
      elif is_dark:
        row_bg_rgb = RGBColor(30, 41, 59) if r_idx % 2 == 0 else RGBColor(22, 30, 46)
      else:
        row_bg_rgb = RGBColor(255, 255, 255)
      cell.fill.fore_color.rgb = row_bg_rgb

      tf = cell.text_frame
      tf.word_wrap = True
      p = tf.paragraphs[0]
      p.font.name = font_name
      p.font.size = Pt(11.5)

      # Color and contrast safety
      default_cell_fg = RGBColor(241, 245, 249) if is_dark else RGBColor(17, 17, 21)
      if "color" in cell_data:
        col_rgb = parse_color_value(cell_data["color"], default=default_cell_fg)
      elif "color_rgb" in cell_data:
        col_rgb = RGBColor(*cell_data["color_rgb"])
      else:
        col_rgb = default_cell_fg

      # Contrast check against row_bg_rgb
      bg_lum = (row_bg_rgb[0] * 0.299 + row_bg_rgb[1] * 0.587 + row_bg_rgb[2] * 0.114)
      fg_lum = (col_rgb[0] * 0.299 + col_rgb[1] * 0.587 + col_rgb[2] * 0.114)
      if bg_lum < 120 and fg_lum < 90:
        col_rgb = RGBColor(241, 245, 249)
      elif bg_lum >= 150 and fg_lum > 180:
        col_rgb = RGBColor(17, 17, 21)

      if cell_data.get("has_badge"):
        p.text = cell_data.get("badge_text") or cell_data["text"]
        p.font.bold = True
        p.font.color.rgb = RGBColor(96, 165, 250) if is_dark else RGBColor(37, 99, 235)
        p.alignment = PP_ALIGN.CENTER
      elif cell_data.get("runs") and len(cell_data["runs"]) > 1:
        p.text = ""
        for r in cell_data["runs"]:
          run = p.add_run()
          run.text = r["text"]
          run.font.name = font_name
          run.font.size = Pt(11.5)
          run.font.bold = r["bold"]
          run.font.color.rgb = col_rgb
      else:
        p.text = cell_data["text"]
        p.font.bold = cell_data["is_bold"]
        p.font.color.rgb = col_rgb

  return table_shape


def apply_semantic_styles_to_table(
    table_shape,
    table_data: Dict[str, Any],
    font_name: str = "Pretendard",
    is_dark: bool = False,
    slide_bg_rgb: tuple[int, int, int] = (15, 23, 42),
) -> None:
  """Guarantees 100% pixel-faithful styling on any native table with theme-adaptive colors and contrast preservation."""
  tbl = table_shape.table
  blend_rgb = (slide_bg_rgb[0], slide_bg_rgb[1], slide_bg_rgb[2])

  # 1. Header styling
  if table_data.get("headers") and len(tbl.rows) > 0:
    for c_idx, h in enumerate(table_data["headers"]):
      if c_idx >= len(tbl.rows[0].cells):
        break
      c = tbl.rows[0].cells[c_idx]
      c.fill.solid()
      h_bg = h.get("bgColor") if isinstance(h, dict) else None
      h_col = h.get("color") if isinstance(h, dict) else None
      if h_bg:
        h_bg_rgb = parse_color_value(h_bg, default=RGBColor(15, 23, 42) if is_dark else RGBColor(241, 245, 249), bg_blend_rgb=blend_rgb)
      else:
        h_bg_rgb = RGBColor(15, 23, 42) if is_dark else RGBColor(241, 245, 249)
      c.fill.fore_color.rgb = h_bg_rgb

      h_fg_rgb = parse_color_value(h_col, default=RGBColor(148, 163, 184) if is_dark else RGBColor(30, 41, 59))
      for p in c.text_frame.paragraphs:
        p.font.name = font_name
        p.font.size = Pt(11.0)
        p.font.bold = True
        p.font.color.rgb = h_fg_rgb
        for r in p.runs:
          r.font.name = font_name
          r.font.size = Pt(11.0)
          r.font.bold = True
          r.font.color.rgb = h_fg_rgb

  # 2. Content rows styling: Theme-adaptive fills, high-contrast readable text
  for r_idx, row_data in enumerate(table_data.get("rows", [])):
    row_num = r_idx + (1 if table_data.get("headers") else 0)
    if row_num >= len(tbl.rows):
      break
    for c_idx, cell_data in enumerate(row_data):
      if c_idx >= len(tbl.columns):
        break
      cell = tbl.cell(row_num, c_idx)
      cell.fill.solid()

      cell_bg = cell_data.get("bgColor")
      if cell_bg:
        row_bg_rgb = parse_color_value(cell_bg, default=RGBColor(30, 41, 59) if is_dark else RGBColor(255, 255, 255), bg_blend_rgb=blend_rgb)
      elif is_dark:
        row_bg_rgb = RGBColor(30, 41, 59) if r_idx % 2 == 0 else RGBColor(22, 30, 46)
      else:
        row_bg_rgb = RGBColor(255, 255, 255) if r_idx % 2 == 0 else RGBColor(248, 250, 252)
      cell.fill.fore_color.rgb = row_bg_rgb

      # Contrast safety
      default_cell_fg = RGBColor(241, 245, 249) if is_dark else RGBColor(17, 17, 21)
      if "color" in cell_data:
        col_rgb = parse_color_value(cell_data["color"], default=default_cell_fg)
      elif "color_rgb" in cell_data:
        col_rgb = RGBColor(*cell_data["color_rgb"])
      else:
        col_rgb = default_cell_fg

      bg_lum = (row_bg_rgb[0] * 0.299 + row_bg_rgb[1] * 0.587 + row_bg_rgb[2] * 0.114)
      fg_lum = (col_rgb[0] * 0.299 + col_rgb[1] * 0.587 + col_rgb[2] * 0.114)
      if bg_lum < 120 and fg_lum < 90:
        col_rgb = RGBColor(241, 245, 249)
      elif bg_lum >= 150 and fg_lum > 180:
        col_rgb = RGBColor(17, 17, 21)

      if not cell.text.strip() and cell_data.get("text"):
        cell.text = cell_data["text"]

      for p in cell.text_frame.paragraphs:
        p.font.name = font_name
        p.font.size = Pt(11.5)
        if cell_data.get("has_badge"):
          p.font.bold = True
          p.font.color.rgb = RGBColor(96, 165, 250) if is_dark else RGBColor(37, 99, 235)
          p.alignment = PP_ALIGN.CENTER
        for r in p.runs:
          r.font.name = font_name
          r.font.size = Pt(11.5)
          if cell_data.get("has_badge"):
            r.font.bold = True
            r.font.color.rgb = RGBColor(96, 165, 250) if is_dark else RGBColor(37, 99, 235)
          else:
            r.font.color.rgb = col_rgb
            if cell_data.get("is_bold"):
              r.font.bold = True


def ensure_slide_canvas_background(
    slide,
    prs: Optional[Presentation] = None,
    bg_color: Optional[RGBColor] = None,
    is_dark: bool = False,
) -> None:
  """Ensures the slide has a guaranteed full-bleed 16:9 canvas background rectangle covering (0, 0, 13.333, 7.5), and removes any duplicate backgrounds."""
  if bg_color is None:
    bg_color = RGBColor(15, 23, 42) if is_dark else RGBColor(255, 255, 255)

  bg_shapes = [
      s for s in slide.shapes
      if s.shape_type == 1 and s.left.inches <= 0.1 and s.top.inches <= 0.1 and s.width.inches >= 13.0 and s.height.inches >= 7.0
  ]
  if len(bg_shapes) > 1:
    for dup in bg_shapes[1:]:
      try:
        slide.shapes._spTree.remove(dup._element)
      except Exception:
        pass

  if bg_shapes:
    full_bleed_bg = bg_shapes[0]
    full_bleed_bg.fill.solid()
    full_bleed_bg.fill.fore_color.rgb = bg_color
    full_bleed_bg.line.fill.background()
  else:
    w = prs.slide_width if prs else Inches(13.333333)
    h = prs.slide_height if prs else Inches(7.5)
    bg_shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), w, h)
    bg_shape.fill.solid()
    bg_shape.fill.fore_color.rgb = bg_color
    bg_shape.line.fill.background()
    try:
      spTree = slide.shapes._spTree
      spTree.insert(2, bg_shape._element)
    except Exception:
      pass


def build_slide_from_geometry(
    slide,
    geom: Dict[str, Any],
    font_name: str = "Pretendard",
    brand_color_rgb: Optional[RGBColor] = None,
) -> None:
  """Deterministically constructs a pixel-faithful native PowerPoint slide from browser-measured DOM layout geometry."""
  if not geom:
    return

  if brand_color_rgb is None:
    brand_color_rgb = RGBColor(37, 99, 235)

  # 1. Canvas Background
  bg_color = parse_color_value(geom.get("slideBgColor"), RGBColor(255, 255, 255))
  is_dark = (bg_color[0] * 0.299 + bg_color[1] * 0.587 + bg_color[2] * 0.114) < 128
  ensure_slide_canvas_background(slide, prs=None, bg_color=bg_color, is_dark=is_dark)
  blend_rgb = (bg_color[0], bg_color[1], bg_color[2])

  # 2. Header Badges (both eyebrow and badge-pill)
  for b_data in geom.get("headerBadges", []):
    r = b_data["rect"]
    b_txt = b_data["text"]
    is_pill = r["left"] > 7.0 or "pill" in str(b_data.get("styles", {})) or "Slide #" in b_txt
    bg_rgb = parse_color_value(b_data["styles"]["backgroundColor"], brand_color_rgb if not is_pill else RGBColor(241, 245, 249), bg_blend_rgb=blend_rgb)
    fg_rgb = parse_color_value(b_data["styles"]["color"], RGBColor(255, 255, 255) if not is_pill else brand_color_rgb)
    b_shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(r["left"]),
        Inches(r["top"]),
        Inches(max(0.8, r["width"])),
        Inches(max(0.24, r["height"])),
    )
    b_shape.fill.solid()
    b_shape.fill.fore_color.rgb = bg_rgb
    b_border = parse_color_value(b_data["styles"].get("borderColor"), None)
    if b_border:
      b_shape.line.color.rgb = b_border
      b_shape.line.width = Pt(1.0)
    else:
      b_shape.line.fill.background()
    tf = b_shape.text_frame
    tf.word_wrap = False
    tf.margin_left = Inches(0.08)
    tf.margin_right = Inches(0.08)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    p = tf.paragraphs[0]
    p.text = b_txt
    p.font.name = font_name
    p.font.size = Pt(b_data["styles"].get("fontSizePt", 10.0))
    p.font.bold = True
    p.font.color.rgb = fg_rgb
    p.alignment = PP_ALIGN.CENTER if is_pill else PP_ALIGN.LEFT

  # 3. Slide Number
  if geom.get("num") and geom["num"]["text"]:
    n_data = geom["num"]
    r = n_data["rect"]
    num_box = slide.shapes.add_textbox(
        Inches(min(12.2, r["left"])),
        Inches(r["top"]),
        Inches(max(0.8, r["width"])),
        Inches(max(0.25, r["height"])),
    )
    p = num_box.text_frame.paragraphs[0]
    p.text = n_data["text"]
    p.font.name = font_name
    p.font.size = Pt(10.5)
    p.font.bold = True
    p.font.color.rgb = parse_color_value(n_data["styles"].get("color") if n_data.get("styles") else None, RGBColor(148, 163, 184) if is_dark else RGBColor(100, 116, 139))

  # 4. Master Title
  if geom.get("title") and geom["title"]["text"]:
    t_data = geom["title"]
    r = t_data["rect"]
    st = t_data["styles"]
    is_center = st.get("textAlign") == "center"
    t_left = Inches(r["left"]) if is_center else Inches(max(0.6, r["left"]))
    t_top = Inches(r["top"])
    t_w = Inches(r["width"]) if is_center else Inches(min(12.0, r["width"]))
    t_h = Inches(max(0.45, r["height"]))
    t_box = slide.shapes.add_textbox(t_left, t_top, t_w, t_h)
    tf = t_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = t_data["text"]
    p.font.name = font_name
    p.font.size = Pt(st.get("fontSizePt", 24.0))
    p.font.bold = True
    p.font.color.rgb = parse_color_value(st.get("color"), RGBColor(248, 250, 252) if is_dark else brand_color_rgb)
    if is_center:
      p.alignment = PP_ALIGN.CENTER

  # 5. Master Subtitle
  if geom.get("sub") and geom["sub"]["text"]:
    s_data = geom["sub"]
    r = s_data["rect"]
    st = s_data["styles"]
    s_left = Inches(max(0.6, r["left"]))
    s_top = Inches(r["top"])
    s_w = Inches(min(12.0, r["width"]))
    s_h = Inches(max(0.30, r["height"]))
    s_box = slide.shapes.add_textbox(s_left, s_top, s_w, s_h)
    tf = s_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = s_data["text"]
    p.font.name = font_name
    p.font.size = Pt(st.get("fontSizePt", 13.0))
    p.font.color.rgb = parse_color_value(st.get("color"), RGBColor(148, 163, 184) if is_dark else RGBColor(100, 116, 139))

  # 6. Description / Intro Paragraphs (Cover & Non-card slides)
  for d in geom.get("desc", []):
    r = d["rect"]
    st = d["styles"]
    d_box = slide.shapes.add_textbox(Inches(r["left"]), Inches(r["top"]), Inches(r["width"]), Inches(r["height"]))
    tf = d_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = d["text"]
    p.font.name = font_name
    p.font.size = Pt(st.get("fontSizePt", 12.0))
    p.font.color.rgb = parse_color_value(st.get("color"), RGBColor(203, 213, 225) if is_dark else RGBColor(74, 77, 82))

  # 7. Native Cards & Layout Containers
  for card in geom.get("cards", []):
    r = card["rect"]
    c_left = Inches(r["left"])
    c_top = Inches(r["top"])
    c_w = Inches(r["width"])
    c_h = Inches(r["height"])

    # Handle Bridge Connector
    if card.get("isBridge") and card.get("bridge"):
      b_info = card["bridge"]
      br = b_info["rect"]
      b_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(br["left"]), Inches(br["top"]), Inches(br["width"]), Inches(br["height"]))
      b_shape.fill.solid()
      b_shape.fill.fore_color.rgb = parse_color_value(b_info["styles"]["backgroundColor"], brand_color_rgb)
      b_shape.line.fill.background()
      tf = b_shape.text_frame
      tf.word_wrap = True
      tf.margin_left = Inches(0.05)
      tf.margin_right = Inches(0.05)
      p0 = tf.paragraphs[0]
      p0.text = b_info.get("arrow", "➔")
      p0.font.name = font_name
      p0.font.size = Pt(15.0)
      p0.font.bold = True
      p0.font.color.rgb = RGBColor(255, 255, 255)
      p0.alignment = PP_ALIGN.CENTER
      p1 = tf.add_paragraph()
      p1.text = b_info.get("text", "")
      p1.font.name = font_name
      p1.font.size = Pt(9.0)
      p1.font.bold = True
      p1.font.color.rgb = RGBColor(255, 255, 255)
      p1.alignment = PP_ALIGN.CENTER
      continue

    # Normal Container / Card
    c_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, c_left, c_top, c_w, c_h)
    c_shape.fill.solid()
    c_bg = parse_color_value(card["styles"]["backgroundColor"], RGBColor(30, 41, 59) if is_dark else RGBColor(255, 255, 255), bg_blend_rgb=blend_rgb)
    c_border = parse_color_value(card["styles"].get("borderColor"), RGBColor(51, 65, 85) if is_dark else RGBColor(226, 232, 240))
    c_shape.fill.fore_color.rgb = c_bg
    c_shape.line.color.rgb = c_border
    c_shape.line.width = Pt(1.0)

    # Handle Attach Card
    if card.get("isAttachCard") and card.get("attachData"):
      ad = card["attachData"]
      tf = c_shape.text_frame
      tf.word_wrap = True
      tf.margin_left = Inches(0.18)
      tf.margin_right = Inches(0.18)
      tf.margin_top = Inches(0.14)
      tf.margin_bottom = Inches(0.14)
      p0 = tf.paragraphs[0]
      p0.text = ad["target"] + (f"  [{ad['mode']}]" if ad.get("mode") else "")
      p0.font.name = font_name
      p0.font.size = Pt(12.5)
      p0.font.bold = True
      p0.font.color.rgb = brand_color_rgb
      if ad.get("mech"):
        p1 = tf.add_paragraph()
        p1.text = ad["mech"]
        p1.font.name = "Consolas"
        p1.font.size = Pt(9.0)
        p1.font.bold = True
        p1.font.color.rgb = RGBColor(24, 107, 242)
      if ad.get("desc"):
        p2 = tf.add_paragraph()
        p2.text = ad["desc"]
        p2.font.name = font_name
        p2.font.size = Pt(10.0)
        p2.font.color.rgb = RGBColor(74, 85, 104)
      continue

    # Top Accent Capsule Pill (guaranteed height 0.06in, fully inset, zero text collision)
    curr_top = c_top
    if card.get("hasTopAccent"):
      bar_color = parse_color_value(card.get("topAccentColor"), brand_color_rgb)
      accent_bar = slide.shapes.add_shape(
          MSO_SHAPE.ROUNDED_RECTANGLE,
          c_left + Inches(0.20),
          c_top + Inches(0.08),
          c_w - Inches(0.40),
          Inches(0.06),
      )
      accent_bar.fill.solid()
      accent_bar.fill.fore_color.rgb = bar_color
      accent_bar.line.fill.background()
      try:
        accent_bar.adjustments[0] = 0.5
      except Exception:
        pass
      curr_top = c_top + Inches(0.18)
    else:
      curr_top = c_top + Inches(0.12)

    # Pipeline Steps container: do NOT draw outer card title
    has_pipeline = len(card.get("pipelineSteps", [])) > 0

    # Card Tag
    if card.get("tag") and card["tag"]["text"]:
      tg = card["tag"]
      tg_r = tg["rect"]
      tg_bg = parse_color_value(tg["styles"]["backgroundColor"], brand_color_rgb, bg_blend_rgb=blend_rgb)
      tg_fg = parse_color_value(tg["styles"]["color"], RGBColor(255, 255, 255))
      tg_shape = slide.shapes.add_shape(
          MSO_SHAPE.ROUNDED_RECTANGLE,
          c_left + Inches(0.18),
          curr_top,
          Inches(max(0.8, tg_r["width"])),
          Inches(max(0.22, tg_r["height"])),
      )
      tg_shape.fill.solid()
      tg_shape.fill.fore_color.rgb = tg_bg
      tg_shape.line.fill.background()
      tf = tg_shape.text_frame
      tf.word_wrap = False
      tf.margin_left = Inches(0.06)
      tf.margin_right = Inches(0.06)
      p = tf.paragraphs[0]
      p.text = tg["text"]
      p.font.name = font_name
      p.font.size = Pt(tg["styles"].get("fontSizePt", 9.5))
      p.font.bold = True
      p.font.color.rgb = tg_fg
      curr_top += Inches(max(0.24, tg_r["height"])) + Inches(0.06)

    # Badges
    for b in card.get("badges", []):
      br = b.get("rect", {})
      b_txt = b.get("text", "")
      if b_txt and br:
        b_w_in = max(0.60, br.get("width", 1.0))
        b_w = Inches(b_w_in)
        b_h = Inches(max(0.22, br.get("height", 0.24)))
        b_left = Inches(br.get("left", r.get("left", 0.0) + r.get("width", 1.0) - b_w_in - 0.15))
        b_top = Inches(br.get("top", r.get("top", 0.0) + 0.18))
        b_shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            b_left,
            b_top,
            b_w,
            b_h,
        )
        b_shape.fill.solid()
        b_bg = parse_color_value(b.get("styles", {}).get("backgroundColor"), RGBColor(241, 245, 249), bg_blend_rgb=blend_rgb)
        b_shape.fill.fore_color.rgb = b_bg
        b_border = parse_color_value(b.get("styles", {}).get("borderColor"), RGBColor(226, 232, 240))
        b_shape.line.color.rgb = b_border
        b_shape.line.width = Pt(1.0)
        tf = b_shape.text_frame
        tf.word_wrap = False
        tf.margin_left = Inches(0.06)
        tf.margin_right = Inches(0.06)
        tf.margin_top = Inches(0.02)
        tf.margin_bottom = Inches(0.02)
        p = tf.paragraphs[0]
        p.text = b_txt
        p.font.name = font_name
        p.font.size = Pt(b.get("styles", {}).get("fontSizePt", 9.5))
        p.font.bold = True
        p.font.color.rgb = parse_color_value(b.get("styles", {}).get("color"), RGBColor(78, 86, 95))
        p.alignment = PP_ALIGN.CENTER

    # Card Title (only if NOT pipeline banner)
    if card.get("title") and card["title"] != "➔" and not has_pipeline:
      b_first_left = min((b["rect"]["left"] for b in card.get("badges", []) if b.get("rect")), default=None)
      if b_first_left is not None:
        title_w = max(Inches(1.0), Inches(b_first_left - (r.get("left", 0.0) + 0.18) - 0.08))
      else:
        title_w = c_w - Inches(0.36)
      t_box = slide.shapes.add_textbox(
          c_left + Inches(0.18),
          curr_top,
          title_w,
          Inches(0.36),
      )
      tf = t_box.text_frame
      tf.word_wrap = True
      p = tf.paragraphs[0]
      p.text = card["title"]
      p.font.name = font_name
      p.font.size = Pt(card.get("titleStyles", {}).get("fontSizePt", 13.5))
      p.font.bold = True
      p.font.color.rgb = parse_color_value(card.get("titleStyles", {}).get("color") if card.get("titleStyles") else None, RGBColor(96, 165, 250) if is_dark else brand_color_rgb)
      curr_top += Inches(0.38)

    # Card Description
    if card.get("desc") and card["desc"]["text"] and not has_pipeline:
      d_box = slide.shapes.add_textbox(c_left + Inches(0.18), curr_top, c_w - Inches(0.36), Inches(0.36))
      tf = d_box.text_frame
      tf.word_wrap = True
      p = tf.paragraphs[0]
      p.text = card["desc"]["text"]
      p.font.name = font_name
      p.font.size = Pt(card["desc"]["styles"].get("fontSizePt", 10.5))
      p.font.color.rgb = parse_color_value(card["desc"]["styles"].get("color"), RGBColor(148, 163, 184) if is_dark else RGBColor(100, 116, 139))
      curr_top += Inches(0.40)

    # Pipeline Steps
    for ps in card.get("pipelineSteps", []):
      psr = ps["rect"]
      ps_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(psr["left"]), Inches(psr["top"]), Inches(psr["width"]), Inches(psr["height"]))
      ps_shape.fill.solid()
      ps_bg = parse_color_value(ps["styles"]["backgroundColor"], RGBColor(255, 255, 255), bg_blend_rgb=blend_rgb)
      ps_shape.fill.fore_color.rgb = ps_bg
      ps_border = parse_color_value(ps["styles"].get("borderColor"), RGBColor(217, 226, 236))
      ps_shape.line.color.rgb = ps_border
      ps_shape.line.width = Pt(1.0)
      tf = ps_shape.text_frame
      tf.word_wrap = True
      tf.margin_left = Inches(0.10)
      tf.margin_right = Inches(0.10)
      p0 = tf.paragraphs[0]
      p0.text = ps["title"]
      p0.font.name = font_name
      p0.font.size = Pt(11.0)
      p0.font.bold = True
      p0.font.color.rgb = parse_color_value(ps["styles"].get("color"), brand_color_rgb)
      if ps.get("desc"):
        p1 = tf.add_paragraph()
        p1.text = ps["desc"]
        p1.font.name = font_name
        p1.font.size = Pt(9.5)
        p1.font.color.rgb = RGBColor(100, 116, 139)

    # Inline Flows
    for fl in card.get("flows", []):
      for node in fl.get("nodes", []):
        nr = node["rect"]
        if node.get("isArrow"):
          arr_box = slide.shapes.add_textbox(Inches(nr["left"]), Inches(nr["top"]), Inches(nr["width"]), Inches(nr["height"]))
          tf = arr_box.text_frame
          tf.word_wrap = False
          p = tf.paragraphs[0]
          p.text = node["text"]
          p.font.name = font_name
          p.font.size = Pt(12.0)
          p.font.bold = True
          p.font.color.rgb = brand_color_rgb
          p.alignment = PP_ALIGN.CENTER
        else:
          n_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(nr["left"]), Inches(nr["top"]), Inches(nr["width"]), Inches(nr["height"]))
          n_shape.fill.solid()
          n_bg = parse_color_value(node["styles"]["backgroundColor"], RGBColor(241, 245, 249) if not node.get("isHighlight") else RGBColor(238, 242, 255), bg_blend_rgb=blend_rgb)
          n_shape.fill.fore_color.rgb = n_bg
          n_border = parse_color_value(node["styles"].get("borderColor"), brand_color_rgb if node.get("isHighlight") else RGBColor(203, 213, 225))
          n_shape.line.color.rgb = n_border
          n_shape.line.width = Pt(1.0)
          tf = n_shape.text_frame
          tf.word_wrap = True
          tf.margin_left = Inches(0.06)
          tf.margin_right = Inches(0.06)
          p = tf.paragraphs[0]
          p.text = node["text"]
          p.font.name = font_name
          p.font.size = Pt(node["styles"].get("fontSizePt", 9.5))
          p.font.bold = True
          p.font.color.rgb = parse_color_value(node["styles"].get("color"), brand_color_rgb if node.get("isHighlight") else RGBColor(30, 41, 59))
          p.alignment = PP_ALIGN.CENTER

    # Feature Items
    for fi in card.get("fItems", []):
      fir = fi["rect"]
      ic_r = fi.get("iconRect")
      tx_r = fi.get("textRect")
      ic_w = 0.0
      ic_left = 0.0
      if fi.get("icon") and ic_r:
        ic_w = max(0.38, ic_r.get("width", 0.40))
        ic_h = max(0.22, ic_r.get("height", 0.24))
        ic_left = ic_r.get("left", fir.get("left", r.get("left", 0.0) + 0.18))
        ic_top = ic_r.get("top", fir.get("top", r.get("top", 0.0) + 0.18))
        ic_shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(ic_left),
            Inches(ic_top),
            Inches(ic_w),
            Inches(ic_h),
        )
        ic_shape.fill.solid()
        ic_bg = parse_color_value(
            fi.get("iconStyles", {}).get("backgroundColor") if fi.get("iconStyles") else None,
            brand_color_rgb,
            bg_blend_rgb=blend_rgb,
        )
        ic_shape.fill.fore_color.rgb = ic_bg
        ic_shape.line.fill.background()
        tf = ic_shape.text_frame
        tf.word_wrap = False
        tf.margin_left = Inches(0.04)
        tf.margin_right = Inches(0.04)
        tf.margin_top = Inches(0.02)
        tf.margin_bottom = Inches(0.02)
        p = tf.paragraphs[0]
        p.text = fi["icon"]
        p.font.name = font_name
        p.font.size = Pt(fi.get("iconStyles", {}).get("fontSizePt", 9.0) if fi.get("iconStyles") else 9.0)
        p.font.bold = True
        p.font.color.rgb = parse_color_value(
            fi.get("iconStyles", {}).get("color") if fi.get("iconStyles") else None,
            RGBColor(255, 255, 255),
        )
        p.alignment = PP_ALIGN.CENTER

      if fi.get("text"):
        card_r = r.get("left", 0.0) + r.get("width", 4.0) - 0.15
        if ic_w > 0:
          min_tx_left = ic_left + ic_w + 0.08
          tx_left = max(tx_r.get("left", min_tx_left) if tx_r else min_tx_left, min_tx_left)
        else:
          tx_left = tx_r.get("left", fir.get("left", r.get("left", 0.0) + 0.18)) if tx_r else fir.get("left", r.get("left", 0.0) + 0.18)

        tx_top = tx_r.get("top", fir.get("top", r.get("top", 0.0) + 0.18)) if tx_r else fir.get("top", r.get("top", 0.0) + 0.18)
        tx_w = max(0.8, card_r - tx_left)
        tx_h = max(0.24, tx_r.get("height", 0.24) if tx_r else 0.24)

        tx_box = slide.shapes.add_textbox(Inches(tx_left), Inches(tx_top), Inches(tx_w), Inches(tx_h))
        tf = tx_box.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.02)
        tf.margin_right = Inches(0.02)
        tf.margin_top = Inches(0.02)
        tf.margin_bottom = Inches(0.02)
        p = tf.paragraphs[0]
        p.text = fi["text"]
        p.font.name = font_name
        p.font.size = Pt(fi.get("textStyles", {}).get("fontSizePt", 10.0) if fi.get("textStyles") else 10.0)
        p.font.color.rgb = parse_color_value(
            fi.get("textStyles", {}).get("color") if fi.get("textStyles") else None,
            RGBColor(17, 17, 21),
        )

    # Sub-Cards
    for sc in card.get("subCards", []):
      scr = sc["rect"]
      sc_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(scr["left"]), Inches(scr["top"]), Inches(scr["width"]), Inches(scr["height"]))
      sc_shape.fill.solid()
      sc_bg = parse_color_value(sc["styles"]["backgroundColor"], RGBColor(248, 250, 252), bg_blend_rgb=blend_rgb)
      sc_shape.fill.fore_color.rgb = sc_bg
      sc_border = parse_color_value(sc["styles"].get("borderColor"), RGBColor(226, 232, 240))
      sc_shape.line.color.rgb = sc_border
      sc_shape.line.width = Pt(1.0)
      tf = sc_shape.text_frame
      tf.word_wrap = True
      tf.margin_left = Inches(0.12)
      tf.margin_right = Inches(0.12)
      tf.margin_top = Inches(0.08)
      tf.margin_bottom = Inches(0.08)
      p_idx = 0
      if sc.get("head"):
        p0 = tf.paragraphs[0]
        p0.text = sc["head"] + (f"  [{sc['badge']}]" if sc.get("badge") else "")
        p0.font.name = font_name
        p0.font.size = Pt(11.0)
        p0.font.bold = True
        p0.font.color.rgb = parse_color_value(sc["styles"].get("color"), brand_color_rgb)
        p_idx += 1
      for it in sc.get("items", []):
        p_it = tf.paragraphs[0] if p_idx == 0 else tf.add_paragraph()
        p_it.text = it if it.startswith(("•", "-", "1", "2")) else f"• {it}"
        p_it.font.name = font_name
        p_it.font.size = Pt(9.5)
        p_it.font.color.rgb = RGBColor(30, 41, 59)
        p_idx += 1

    # Code Blocks & Spec/Example Blocks
    for cb in card.get("codeBlocks", []):
      cbr = cb["rect"]
      cb_h = Inches(cbr["height"]) if cbr.get("height", 0) > 0.3 else Inches(0.40)
      cb_box = slide.shapes.add_shape(
          MSO_SHAPE.ROUNDED_RECTANGLE,
          Inches(cbr["left"]),
          Inches(cbr["top"]),
          Inches(cbr["width"]),
          cb_h,
      )
      cb_box.fill.solid()
      cb_bg = parse_color_value(cb["styles"]["backgroundColor"], RGBColor(241, 245, 249), bg_blend_rgb=blend_rgb)
      cb_box.fill.fore_color.rgb = cb_bg
      cb_border = parse_color_value(cb["styles"].get("borderColor"), RGBColor(226, 232, 240))
      cb_box.line.color.rgb = cb_border
      cb_box.line.width = Pt(1.0)
      tf = cb_box.text_frame
      tf.word_wrap = True
      tf.margin_left = Inches(0.08)
      tf.margin_right = Inches(0.08)
      tf.margin_top = Inches(0.04)
      tf.margin_bottom = Inches(0.04)
      for p_idx, line in enumerate(cb["text"].split("\n")):
        p = tf.paragraphs[0] if p_idx == 0 else tf.add_paragraph()
        p.text = line
        p.font.name = "Consolas"
        p.font.size = Pt(cb["styles"].get("fontSizePt", 9.0))
        p.font.color.rgb = parse_color_value(cb["styles"].get("color"), RGBColor(30, 41, 59))

    # Remaining Paragraphs & List Items
    paras = card.get("paragraphs", [])
    if paras and not card.get("subCards") and not card.get("pipelineSteps"):
      p_box = slide.shapes.add_textbox(
          c_left + Inches(0.18),
          curr_top,
          c_w - Inches(0.36),
          max(Inches(0.5), c_h - (curr_top - c_top) - Inches(0.10)),
      )
      tf = p_box.text_frame
      tf.word_wrap = True
      tf.margin_top = Inches(0.02)
      tf.margin_bottom = Inches(0.02)
      for p_idx, para in enumerate(paras):
        p = tf.paragraphs[0] if p_idx == 0 else tf.add_paragraph()
        p.text = para["text"] if para["text"].startswith(("•", "-", "1", "2")) else f"• {para['text']}"
        p.font.name = font_name
        is_sub = len(para["text"]) > 15 and not para["text"].startswith(("1", "2", "3", "•", "✔", "❌", "💬", "📦", "🔑", "🔄"))
        p.font.size = Pt(10.5 if is_sub else 11.5)
        p.font.bold = not is_sub
        p_col = parse_color_value(para["styles"]["color"], RGBColor(226, 232, 240) if is_dark else RGBColor(17, 17, 21))
        p.font.color.rgb = p_col


def extract_global_design_tokens(
    first_image_path: str = "",
    raw_html: Optional[str] = None,
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
    if "--brand-color" in css_vars:
      tokens["brand_color_hex"] = css_vars["--brand-color"]
      tokens["badge_bg_hex"] = css_vars["--brand-color"]
    elif "--brand" in css_vars:
      tokens["brand_color_hex"] = css_vars["--brand"]
      tokens["badge_bg_hex"] = css_vars["--brand"]
    elif "--app-primary" in css_vars:
      tokens["brand_color_hex"] = css_vars["--app-primary"]
      tokens["badge_bg_hex"] = css_vars["--app-primary"]
    elif "--primary" in css_vars:
      tokens["brand_color_hex"] = css_vars["--primary"]
      tokens["badge_bg_hex"] = css_vars["--primary"]
    elif "--accent" in css_vars:
      tokens["brand_color_hex"] = css_vars["--accent"]
      tokens["badge_bg_hex"] = css_vars["--accent"]

    if "--app-background" in css_vars:
      tokens["bg_color_hex"] = css_vars["--app-background"]
    elif "--background" in css_vars:
      tokens["bg_color_hex"] = css_vars["--background"]

    if "--app-foreground" in css_vars:
      tokens["text_main_hex"] = css_vars["--app-foreground"]
    elif "--text-main" in css_vars:
      tokens["text_main_hex"] = css_vars["--text-main"]
    elif "--foreground" in css_vars:
      tokens["text_main_hex"] = css_vars["--foreground"]

    if "--app-muted-foreground" in css_vars:
      tokens["text_muted_hex"] = css_vars["--app-muted-foreground"]
    elif "--app-muted" in css_vars:
      tokens["text_muted_hex"] = css_vars["--app-muted"]
    elif "--text-muted" in css_vars:
      tokens["text_muted_hex"] = css_vars["--text-muted"]
    elif "--muted" in css_vars:
      tokens["text_muted_hex"] = css_vars["--muted"]

    # Detect dark theme and guarantee readable text defaults
    bg_tuple = _hex_to_rgb_tuple(tokens.get("bg_color_hex", "#FFFFFF"), (255, 255, 255))
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


def generate_native_slide_builder_code(
    image_path: str,
    slide_dom_data: Optional[Dict[str, Any]] = None,
    slide_geometry: Optional[Dict[str, Any]] = None,
    model_name: str = DEFAULT_GEMINI_MODEL,
    max_retries: int = 2,
    design_tokens: Optional[Dict[str, Any]] = None,
    slide_index: Optional[int] = None,
    total_slides: Optional[int] = None,
) -> Optional[str]:
  """Uses Gemini Vision fusing 16:9 screenshot and semantic HTML DOM to generate 100% native slide builder code."""
  from app.slide_fidelity_critic import audit_slide_fidelity, run_code_on_test_slide
  try:
    with open(image_path, "rb") as f:
      image_bytes = f.read()

    client = get_genai_client()

    tokens = design_tokens or extract_global_design_tokens()
    b_r, b_g, b_b = _hex_to_rgb_tuple(tokens.get("brand_color_hex", "#2563EB"), (37, 99, 235))
    t_r, t_g, t_b = _hex_to_rgb_tuple(tokens.get("text_main_hex", "#111115"), (17, 17, 21))
    m_r, m_g, m_b = _hex_to_rgb_tuple(tokens.get("text_muted_hex", "#64748B"), (100, 116, 139))
    bg_r, bg_g, bg_b = _hex_to_rgb_tuple(tokens.get("bg_color_hex", "#FFFFFF"), (255, 255, 255))
    is_dark = (bg_r * 0.299 + bg_g * 0.587 + bg_b * 0.114) < 128
    font_family = tokens.get("font_name", "Pretendard")
    slide_label = f"Slide {slide_index} of {total_slides}" if slide_index and total_slides else "Slide"

    # Measured DOM layout coordinates snippet
    measured_layout_snippet = ""
    if slide_geometry:
      geom_clean = {k: v for k, v in slide_geometry.items() if v}
      geom_json = json.dumps(geom_clean, ensure_ascii=False, indent=2)
      measured_layout_snippet = f"""
[EXACT MEASURED BROWSER LAYOUT GEOMETRY (PIXEL-PERFECT INCHES)]
Use these exact coordinates and styling values measured directly from Chrome:
{geom_json}
"""

    # Semantic context if parsed DOM is available
    semantic_context = ""
    has_tables = False
    table_snippet = ""
    if slide_dom_data:
      has_tables = len(slide_dom_data.get("tables", [])) > 0
      semantic_json = json.dumps(
          {k: v for k, v in slide_dom_data.items() if k != "raw_html"},
          ensure_ascii=False,
          indent=2,
      )
      semantic_context = f"""
[PARSED SEMANTIC HTML TRUTH]
{semantic_json}

[SLIDE ORIGINAL HTML SOURCE]
```html
{slide_dom_data.get('raw_html', '')}
```
"""
      if has_tables:
        t_data_json = json.dumps(slide_dom_data["tables"][0], ensure_ascii=False)
        is_slide_dark = (tokens.get("text_main_hex") == "#F8FAFC") or (_hex_to_rgb_tuple(tokens.get("bg_color_hex", "#FFFFFF"))[0] < 128)
        if is_slide_dark:
          table_style_rules = """- Dark theme header background `RGBColor(15, 23, 42)` with muted light text `RGBColor(148, 163, 184)`.
- Alternating dark row fills (`RGBColor(30, 41, 59)` and `RGBColor(22, 30, 46)`).
- Subtotal accent blue row fill (`RGBColor(28, 48, 79)`) with light blue text `RGBColor(147, 197, 253)`.
- Total row deep navy fill (`RGBColor(15, 23, 42)`) with bright white bold text `RGBColor(255, 255, 255)`.
- High contrast readable cell text: default body text `RGBColor(248, 250, 252)`. Never use dark text on dark cells!"""
        else:
          table_style_rules = f"""- Light header background `RGBColor(241, 245, 249)` with bold accent text `RGBColor({b_r}, {b_g}, {b_b})`.
- Alternating row fills (`#FFFFFF` and `#F8FAFC`).
- Semantic cell alert colors:
  * Red alert text (`#D93025`): `RGBColor(217, 48, 37)`
  * Gold alert text (`#B48000`): `RGBColor(180, 128, 0)`
  * Green alert text (`#008744`): `RGBColor(0, 135, 68)`
  * Parameter keys in Column 1: `bold = True`"""
        table_snippet = f"""
[TABLE BUILDER INSTRUCTIONS - 100% PRECISE DATA & STYLING]
This slide contains a data table. You can use the built-in helper function available in scope:
`build_styled_native_table(slide, table_data, left, top, width, height, col_ratios=None)`
Example:
```python
t_data = {t_data_json}
build_styled_native_table(slide, t_data, Inches(0.8), Inches(1.95), Inches(11.733), Inches(4.5))
```
Or you may build the table natively using `slide.shapes.add_table`, adhering strictly to:
{table_style_rules}
"""

    # Build structured Slide Manifest for the slide-building AI
    manifest_lines = []
    manifest_lines.append(f"Canvas: 16:9 Widescreen (13.333\" x 7.5\"), Background = RGBColor({bg_r}, {bg_g}, {bg_b}) (HEX: {tokens.get('bg_color_hex', '#FFFFFF')}, is_dark={is_dark})")
    manifest_lines.append(f"Global Font Family: {font_family}")
    
    if slide_geometry:
      header_badges = slide_geometry.get("headerBadges", [])
      if header_badges:
        for hb_idx, hb in enumerate(header_badges):
          r = hb.get("rect", {})
          st = hb.get("styles", {})
          manifest_lines.append(f"Header Category Pill/Badge #{hb_idx+1}: \"{hb.get('text')}\" at (left={r.get('left')}, top={r.get('top')}, width={r.get('width')}, height={r.get('height')}), fontSize={st.get('fontSizePt', 10.0)}pt, color={st.get('color')}, bg={st.get('backgroundColor')}")
      elif slide_geometry.get("tag"):
        t = slide_geometry["tag"]
        r = t.get("rect", {})
        st = t.get("styles", {})
        manifest_lines.append(f"Header Category Pill/Badge: \"{t.get('text')}\" at (left={r.get('left')}, top={r.get('top')}, width={r.get('width')}, height={r.get('height')}), fontSize={st.get('fontSizePt', 10.0)}pt, color={st.get('color')}, bg={st.get('backgroundColor')}")
      
      if slide_geometry.get("title"):
        t = slide_geometry["title"]
        r = t.get("rect", {})
        st = t.get("styles", {})
        manifest_lines.append(f"Main Slide Title: \"{t.get('text')}\" at (left={r.get('left')}, top={r.get('top')}, width={r.get('width')}, height={r.get('height')}), fontSize={st.get('fontSizePt', 24.0)}pt, bold=True, color={st.get('color')}")
      
      if slide_geometry.get("sub"):
        s = slide_geometry["sub"]
        r = s.get("rect", {})
        st = s.get("styles", {})
        manifest_lines.append(f"Subtitle / Description: \"{s.get('text')}\" at (left={r.get('left')}, top={r.get('top')}, width={r.get('width')}, height={r.get('height')}), fontSize={st.get('fontSizePt', 13.0)}pt, color={st.get('color')}")
      
      cards = slide_geometry.get("cards", [])
      if cards:
        manifest_lines.append(f"\nVisual Layout Containers ({len(cards)} containers detected):")
        for c_idx, c in enumerate(cards):
          r = c.get("rect", {})
          st = c.get("styles", {})
          is_rnd = st.get("isRounded", st.get("borderRadius", 0) >= 6)
          c_title = c.get("title") or "Container"
          manifest_lines.append(f"  Container #{c_idx+1}: [{c_title}] at (left={r.get('left')}, top={r.get('top')}, width={r.get('width')}, height={r.get('height')})")
          manifest_lines.append(f"    - Shape: {'MSO_SHAPE.ROUNDED_RECTANGLE' if is_rnd else 'MSO_SHAPE.RECTANGLE'} (borderRadius={st.get('borderRadius', 0)}px)")
          manifest_lines.append(f"    - Fill: {st.get('backgroundColor')}, Border: {st.get('borderColor')}")
          if c.get("hasTopAccent"):
            manifest_lines.append(f"    - Top Accent Bar: color={c.get('topAccentColor')}")
          if c.get("isBridge") and c.get("bridge"):
            manifest_lines.append(f"    - Bridge Connector Banner: \"{c['bridge'].get('text')}\" (arrow={c['bridge'].get('arrow')})")
          if c.get("isAttachCard") and c.get("attachData"):
            ad = c["attachData"]
            manifest_lines.append(f"    - Attach Card Details: target=\"{ad.get('target')}\", mode=\"{ad.get('mode')}\", mech=\"{ad.get('mech')}\", desc=\"{ad.get('desc')}\"")
          if c.get("pipelineSteps"):
            manifest_lines.append(f"    - Pipeline Steps ({len(c['pipelineSteps'])} steps):")
            for ps in c["pipelineSteps"]:
              manifest_lines.append(f"      * Step {ps.get('stepNum')}: \"{ps.get('title')}\" - {ps.get('desc')}")
          if c.get("subCards"):
            manifest_lines.append(f"    - Nested Sub-Cards ({len(c['subCards'])} cards):")
            for sc in c["subCards"]:
              manifest_lines.append(f"      * [{sc.get('head', '')}] badge='{sc.get('badge', '')}', items={sc.get('items', [])}")
          if c.get("flows"):
            manifest_lines.append(f"    - Inline Flows:")
            for fl in c["flows"]:
              node_texts = [n.get('text', '') if isinstance(n, dict) else str(n) for n in fl.get('nodes', [])]
              manifest_lines.append(f"      * Sequence: {' -> '.join(node_texts)} (arrow: {fl.get('arrow')})")
          if c.get("fItems"):
            manifest_lines.append(f"    - Feature Items with Icon Badges:")
            for fi in c["fItems"]:
              manifest_lines.append(f"      * [{fi.get('icon')}] \"{fi.get('text')}\"")
          for b in c.get("badges", []):
            manifest_lines.append(f"    - Inner Badge: \"{b.get('text')}\"")
          for p in c.get("paragraphs", []):
            manifest_lines.append(f"    - Text Item: \"{p.get('text')}\" (color={p.get('styles', {}).get('color')})")
          for cb in c.get("codeBlocks", []):
            manifest_lines.append(f"    - Code / Calculation Block: \"{cb.get('text')}\"")
      
      tables = slide_geometry.get("tables", [])
      if tables:
        manifest_lines.append(f"\nData Tables ({len(tables)} detected):")
        for t_idx, tbl in enumerate(tables):
          r = tbl.get("rect", {})
          manifest_lines.append(f"  Table #{t_idx+1}: {tbl.get('num_rows')} rows x {tbl.get('num_cols')} columns at (left={r.get('left')}, top={r.get('top')}, width={r.get('width')}, height={r.get('height')})")
      
      hl = slide_geometry.get("highlightBoxes", [])
      if hl:
        manifest_lines.append(f"\nCallout / Highlight Boxes ({len(hl)} detected):")
        for h_idx, h in enumerate(hl):
          r = h.get("rect", {})
          manifest_lines.append(f"  Callout #{h_idx+1}: \"{h.get('text')}\" at (left={r.get('left')}, top={r.get('top')}, width={r.get('width')}, height={r.get('height')})")
          manifest_lines.append(f"    - Fill: {h.get('styles', {}).get('backgroundColor')}, Border: {h.get('styles', {}).get('borderColor')}")
      
      fn = slide_geometry.get("footnotes", [])
      if fn:
        manifest_lines.append(f"\nFootnotes ({len(fn)} detected):")
        for f_idx, f in enumerate(fn):
          r = f.get("rect", {})
          manifest_lines.append(f"  Footnote #{f_idx+1}: \"{f.get('text')}\" at (left={r.get('left')}, top={r.get('top')})")

    manifest_context = "\n".join(manifest_lines)

    base_prompt = f"""
You are an expert PowerPoint Engineer specializing in high-fidelity native PPTX recreation.
Analyze this 16:9 widescreen slide image ({slide_label}) and the semantic HTML structure & style manifest below.
Generate a Python function `def build_slide(prs, slide):` using `python-pptx` that reconstructs this slide into 100% native, fully editable shapes, text boxes, and tables.

[AVAILABLE IN EXECUTION SCOPE]
- `prs`, `slide`, `Inches`, `Pt`, `RGBColor`, `MSO_SHAPE`, `PP_ALIGN`, `MSO_ANCHOR`
- `build_styled_native_table(slide, table_data, left, top, width, height, col_ratios=None)`

[AUTHENTIC SLIDE STRUCTURE & STYLE MANIFEST (MEASURED GROUND TRUTH)]
{manifest_context}

[STRICT INSTRUCTIONS FOR THE SLIDE BUILDER AI]
1. ZERO HARDCODING (DYNAMIC RECONSTRUCTION):
   - DO NOT assume any fixed master grid or fixed template coordinates.
   - Faithfully reconstruct the exact layout, containers, and elements defined in the MANIFEST above.
   - If there are 4 metric cards, build 4 metric cards. If there are 2 comparison containers, build 2 comparison containers. If there is a table and 3 bottom cards, build them exactly in that order!
2. EXACT SHAPES & CORNER RADII (MSO_SHAPE.ROUNDED_RECTANGLE):
   - Any container or box with `isRounded=True` or `borderRadius >= 6px` MUST be created using `MSO_SHAPE.ROUNDED_RECTANGLE`. Never create sharp rectangular boxes where rounded cards exist!
   - For cards with top accent colors, add an inner rounded capsule pill (`MSO_SHAPE.ROUNDED_RECTANGLE` with `adjustments[0] = 0.5`, height ~0.06 in, inset inside the card: `left = card_left + Inches(0.20)`, `top = card_top + Inches(0.08)`, `width = card_width - Inches(0.40)`).
3. 100% FAITHFUL STYLING & HIGH CONTRAST:
   - Canvas background MUST be full-bleed 16:9 rectangle: `slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))` with fill `RGBColor({bg_r}, {bg_g}, {bg_b})` and no border line.
   - For every container and text run, apply the exact RGB colors from the manifest.
   - Dark theme rule: On dark containers (e.g. #0F172A, #1E293B), all text MUST be high contrast (white or light tint like RGBColor(248, 250, 252) or light slate RGBColor(148, 163, 184)). NEVER use dark text on dark cards!
4. ZERO OMISSION MANDATE (100% Verbatim):
   - Transcribe 100% of all Korean and English titles, metrics, card items, bullet points, nested spec/calculation blocks, callout boxes (💡핵심 요약, 📌참고, 💡결론), and bottom footnotes (* ...) verbatim!
   - Every single bullet point must be preserved without skipping or summarizing.
5. ZERO CLIPPING & DYNAMIC HEIGHT BUDGETING:
   - Set `tf.word_wrap = True` and give text boxes adequate height so text never overflows or gets clipped.
   - When a table is present, place bottom cards below the table without overlapping.
{table_snippet}
{measured_layout_snippet}
{semantic_context}

Output ONLY executable Python code defining:
def build_slide(prs, slide):
    # implementation
"""

    current_prompt = base_prompt
    best_code = None

    for attempt in range(1, max_retries + 1):
      code = ""
      for api_try in range(3):
        try:
          response = client.models.generate_content(
              model=model_name,
              contents=[
                  types.Part.from_bytes(data=image_bytes, mime_type="image/png"),
                  current_prompt,
              ],
              config=types.GenerateContentConfig(
                  thinking_config=types.ThinkingConfig(thinking_level=types.ThinkingLevel.MEDIUM),
                  temperature=0.1,
              ),
          )
          code = response.text or ""
          if code:
            break
        except Exception as api_err:
          print(f"[Vision CodeGen API Try {api_try + 1}] Transient error: {api_err}")
          if api_try < 2:
            import time
            time.sleep(3 * (api_try + 1))
      if not code:
        continue
      best_code = code

      # 1. Structural runtime safety check
      success, test_slide, test_prs, exec_err = run_code_on_test_slide(code)
      if not success:
        print(f"[Slide CodeGen Attempt {attempt}] Execution error: {exec_err}")
        current_prompt = (
            base_prompt
            + f"\n\n[CRITICAL ERROR IN PREVIOUS ATTEMPT]:\nExecution failed with error: {exec_err}\nFix syntax, imports, and variables, and regenerate complete code."
        )
        continue

      # 2. Slide Fidelity & Zero-Omission Critic Audit
      report = audit_slide_fidelity(
          original_image_input=image_bytes,
          generated_code=code,
          slide=test_slide,
          model_name=model_name,
      )

      if report.passed and report.score >= 7:
        print(f"[Slide Fidelity Audit PASSED on Attempt {attempt}] Score: {report.score}/10")
        return code
      else:
        print(
            f"[Slide Fidelity Audit REJECTED on Attempt {attempt}] Score: {report.score}/10."
            f" Issues: {report.design_and_layout_issues}, Missing: {report.missing_or_truncated_text}"
        )
        current_prompt = (
            base_prompt
            + "\n\n[PREVIOUS ATTEMPT REJECTED BY FIDELITY AUDITOR - CORRECTIVE DIRECTIVES]:\n"
            + f"- Score: {report.score}/10 (Must be >= 7)\n"
            + (f"- Missing/Truncated text: {', '.join(report.missing_or_truncated_text)}\n" if report.missing_or_truncated_text else "")
            + (f"- Design/Layout defects: {', '.join(report.design_and_layout_issues)}\n" if report.design_and_layout_issues else "")
            + f"- Instructions: {report.corrective_instructions}\n"
            + "Please fix ALL issues and regenerate the complete, flawless Python code."
        )

    return best_code
  except Exception as e:
    print(f"[Vision CodeGen Error] {type(e).__name__}: {e}")
    return None


def ensure_slide_typography_consistency(
    slide,
    default_font: str = "Pretendard",
    default_color: Optional[RGBColor] = None,
    is_dark: bool = False,
) -> None:
  """Guarantees 100% typography consistency across all paragraphs, runs, and table cells on the slide."""
  if default_color is None:
    default_color = RGBColor(241, 245, 249) if is_dark else RGBColor(17, 17, 21)

  for shape in slide.shapes:
    if shape.has_text_frame:
      tf = shape.text_frame
      tf.word_wrap = True
      for p in tf.paragraphs:
        if not p.font.name:
          p.font.name = default_font
        if not p.font.size:
          p.font.size = Pt(11.0)
        try:
          if not p.font.color or p.font.color.type != 1:
            p.font.color.rgb = default_color
        except Exception:
          pass
        for run in p.runs:
          if not run.font.name:
            run.font.name = default_font
          if not run.font.size:
            run.font.size = p.font.size
    elif shape.has_table:
      for row in shape.table.rows:
        for cell in row.cells:
          cell.text_frame.word_wrap = True
          for p in cell.text_frame.paragraphs:
            if not p.font.name:
              p.font.name = default_font
            if not p.font.size:
              p.font.size = Pt(10.0)
            try:
              if not p.font.color or p.font.color.type != 1:
                p.font.color.rgb = default_color
            except Exception:
              pass
            for run in p.runs:
              if not run.font.name:
                run.font.name = default_font
              if not run.font.size:
                run.font.size = p.font.size


def refine_card_accent_bars(slide) -> None:
  """Refines any card top accent bars into sleek rounded capsule pills inset inside cards, and guarantees zero collision with text."""
  cards = [sh for sh in slide.shapes if sh.height.inches > 0.8 and sh.width.inches > 1.5 and sh.shape_type == 1]
  thin_bars = [sh for sh in slide.shapes if 0.03 <= sh.height.inches <= 0.20 and sh.width.inches > 1.2 and sh.top.inches < 6.5]
  for b in thin_bars:
    for c in cards:
      if abs(c.left.inches - b.left.inches) < 0.6 and abs(c.top.inches - b.top.inches) < 0.5:
        inset_x = Inches(0.24)
        inset_y = Inches(0.06)
        b.left = c.left + inset_x
        b.top = c.top + inset_y
        b.width = max(Inches(0.5), c.width - (inset_x * 2))
        b.height = Inches(0.06)
        try:
          b.adjustments[0] = 0.5  # Fully rounded capsule pill
        except Exception:
          pass
        try:
          b.line.fill.background()
        except Exception:
          pass

        # Guarantee zero text collision: ensure text boxes inside card start below accent bar
        for tb in slide.shapes:
          if tb.has_text_frame and not tb.has_table and tb != c and tb != b:
            if c.left.inches - 0.1 <= tb.left.inches <= c.left.inches + c.width.inches + 0.1:
              if c.top.inches - 0.05 <= tb.top.inches < c.top.inches + 0.18:
                tb.top = c.top + Inches(0.18)
        break


def audit_and_resolve_slide_collisions(slide) -> None:
  """Audits and guarantees zero shape collisions, deduplicates full-bleed backgrounds and text, and prevents badges from overlapping text."""
  # 1. Deduplicate full bleed backgrounds
  bg_shapes = [
      s for s in slide.shapes
      if s.shape_type == 1 and s.left.inches <= 0.1 and s.top.inches <= 0.1 and s.width.inches >= 13.0 and s.height.inches >= 7.0
  ]
  if len(bg_shapes) > 1:
    for dup in bg_shapes[1:]:
      try:
        slide.shapes._spTree.remove(dup._element)
      except Exception:
        pass

  # 2. Deduplicate exact duplicate text frames at identical positions
  seen_signatures = set()
  for s in list(slide.shapes):
    if s.has_text_frame and not s.has_table:
      txt = s.text_frame.text.strip()
      if txt and len(txt) > 5:
        sig = (round(s.left.inches, 2), round(s.top.inches, 2), txt)
        if sig in seen_signatures:
          try:
            slide.shapes._spTree.remove(s._element)
          except Exception:
            pass
          continue
        seen_signatures.add(sig)

  # 3. Resolve horizontal badge vs text collisions
  shapes = [
      s for s in slide.shapes
      if not (s.left.inches <= 0.1 and s.top.inches <= 0.1 and s.width.inches >= 13.0)
  ]

  for i in range(len(shapes)):
    s1 = shapes[i]
    l1, t1, w1, h1 = s1.left.inches, s1.top.inches, s1.width.inches, s1.height.inches
    r1, b1 = l1 + w1, t1 + h1

    for j in range(len(shapes)):
      if i == j:
        continue
      s2 = shapes[j]
      l2, t2, w2, h2 = s2.left.inches, s2.top.inches, s2.width.inches, s2.height.inches
      r2, b2 = l2 + w2, t2 + h2

      # Skip if one contains the other (e.g. card container containing text or badge)
      if l1 <= l2 + 0.05 and t1 <= t2 + 0.05 and r1 >= r2 - 0.05 and b1 >= b2 - 0.05:
        continue
      if l2 <= l1 + 0.05 and t2 <= t1 + 0.05 and r2 >= r1 - 0.05 and b2 >= b1 - 0.05:
        continue

      # Check overlap in both dimensions
      x_overlap = min(r1, r2) - max(l1, l2)
      y_overlap = min(b1, b2) - max(t1, t2)

      if x_overlap > 0.02 and y_overlap > 0.05:
        # Case A: S1 is AUTO_SHAPE badge/pill to the left of S2 TEXT_BOX
        if s1.shape_type == 1 and s2.shape_type == 17 and l1 < l2:
          new_left = r1 + 0.08
          shift = new_left - l2
          s2.left = Inches(new_left)
          s2.width = Inches(max(0.6, w2 - shift))
          l2, r2 = new_left, new_left + max(0.6, w2 - shift)

        # Case B: S1 is TEXT_BOX to the left of S2 AUTO_SHAPE badge
        elif s1.shape_type == 17 and s2.shape_type == 1 and l1 < l2:
          new_w = max(0.6, l2 - l1 - 0.08)
          s1.width = Inches(new_w)
          w1, r1 = new_w, l1 + new_w

        # Case C: S1 and S2 are both TEXT_BOX in the same row
        elif s1.shape_type == 17 and s2.shape_type == 17 and l1 < l2:
          new_w = max(0.6, l2 - l1 - 0.08)
          s1.width = Inches(new_w)
          w1, r1 = new_w, l1 + new_w


def build_native_slide_from_code(
    slide,
    prs: Presentation,
    code_str: Optional[str],
    fallback_img_path: str,
    design_tokens: Optional[Dict[str, Any]] = None,
    slide_dom_data: Optional[Dict[str, Any]] = None,
    slide_geometry: Optional[Dict[str, Any]] = None,
):
  """Executes generated python-pptx builder code to populate native slide, with geometry native builder and 2-tier fallback."""
  font_family = (design_tokens or {}).get("font_name", "Pretendard")
  text_color_hex = (design_tokens or {}).get("text_main_hex", "#111115")
  default_text_color = _hex_to_pptx_color(text_color_hex, default=(17, 17, 21))
  brand_color = _hex_to_pptx_color((design_tokens or {}).get("brand_color_hex"), default=(37, 99, 235))

  # Authoritative slide background & theme detection
  geom_bg = (slide_geometry or {}).get("slideBgColor")
  token_bg = (design_tokens or {}).get("bg_color_hex")
  slide_bg_rgb = parse_color_value(geom_bg or token_bg, default=RGBColor(255, 255, 255))
  is_dark = (slide_bg_rgb[0] * 0.299 + slide_bg_rgb[1] * 0.587 + slide_bg_rgb[2] * 0.114) < 128
  if default_text_color == RGBColor(17, 17, 21) and is_dark:
    default_text_color = RGBColor(241, 245, 249)

  # Pre-guarantee full-bleed 16:9 canvas background shape
  ensure_slide_canvas_background(slide, prs=prs, bg_color=slide_bg_rgb, is_dark=is_dark)

  if not code_str:
    if slide_geometry:
      build_slide_from_geometry(slide, slide_geometry, font_name=font_family, brand_color_rgb=brand_color)
      ensure_slide_canvas_background(slide, prs=prs, bg_color=slide_bg_rgb, is_dark=is_dark)
      ensure_slide_typography_consistency(slide, default_font=font_family, default_color=default_text_color, is_dark=is_dark)
      refine_card_accent_bars(slide)
      audit_and_resolve_slide_collisions(slide)
      return

    decomp = decompose_slide_image_with_vision(fallback_img_path)
    if decomp and decomp.elements:
      build_native_pptx_slide(slide, decomp, prs, fallback_img_path)
      ensure_slide_canvas_background(slide, prs=prs, bg_color=slide_bg_rgb, is_dark=is_dark)
      ensure_slide_typography_consistency(slide, default_font=font_family, default_color=default_text_color, is_dark=is_dark)
      return

    slide.shapes.add_picture(
        fallback_img_path,
        Inches(0),
        Inches(0),
        width=prs.slide_width,
        height=prs.slide_height,
    )
    return

  clean_code = code_str
  if "```python" in clean_code:
    clean_code = clean_code.split("```python", 1)[1].split("```", 1)[0].strip()
  elif "```" in clean_code:
    clean_code = clean_code.split("```", 1)[1].split("```", 1)[0].strip()

  scope = {
      "prs": prs,
      "slide": slide,
      "Presentation": Presentation,
      "Inches": Inches,
      "Pt": Pt,
      "RGBColor": RGBColor,
      "MSO_SHAPE": MSO_SHAPE,
      "PP_ALIGN": PP_ALIGN,
      "MSO_ANCHOR": MSO_ANCHOR,
      "build_styled_native_table": build_styled_native_table,
  }

  try:
    exec(clean_code, scope)
    if "build_slide" in scope:
      func = scope["build_slide"]
      import inspect
      sig = inspect.signature(func)
      if len(sig.parameters) == 1:
        func(slide)
      else:
        func(prs, slide)

      # Ensure canvas background exists even if code wiped or drew over it
      ensure_slide_canvas_background(slide, prs=prs, bg_color=slide_bg_rgb, is_dark=is_dark)

      # Safety Net: Table construction and styling
      tbl_dom = (slide_geometry.get("tables", [{}])[0] if slide_geometry and slide_geometry.get("tables")
                 else (slide_dom_data.get("tables", [{}])[0] if slide_dom_data and slide_dom_data.get("tables") else None))
      if tbl_dom:
        has_table_shape = any(s.has_table for s in slide.shapes)
        if not has_table_shape and tbl_dom.get("rows"):
          print("[Safety Net] Constructing table via deterministic builder because it was missing in executed code.")
          tbl_r = tbl_dom.get("rect", {})
          t_left = Inches(tbl_r.get("left", 0.6))
          t_top = Inches(tbl_r.get("top", 1.5))
          t_w = Inches(min(tbl_r.get("width", 12.0), 12.2))
          t_h = Inches(tbl_r.get("height", 3.0))
          build_styled_native_table(slide, tbl_dom, t_left, t_top, t_w, t_h, font_name=font_family)
        else:
          for s in slide.shapes:
            if s.has_table:
              apply_semantic_styles_to_table(s, tbl_dom, font_name=font_family, is_dark=is_dark, slide_bg_rgb=(slide_bg_rgb[0], slide_bg_rgb[1], slide_bg_rgb[2]))

      # Safety Net: Zero Omission for Footnotes, Callouts, and Bottom Cards
      if slide_geometry:
        slide_text_corpus = " ".join([
            p.text for s in slide.shapes if s.has_text_frame for p in s.text_frame.paragraphs
        ] + [
            c.text for s in slide.shapes if s.has_table for row in s.table.rows for c in row.cells
        ]).lower()

        # Check highlight boxes
        for hl in slide_geometry.get("highlightBoxes", []):
          hl_snippet = hl["text"][:15].lower()
          if hl_snippet not in slide_text_corpus:
            print(f"[Safety Net] Appending omitted highlight box: {hl['text'][:30]}...")
            hl_r = hl.get("rect", {})
            hl_w = Inches(hl_r.get("width", 11.733))
            hl_l = Inches(hl_r.get("left", 0.8))
            hl_t = Inches(hl_r.get("top", 6.5))
            hl_h = Inches(max(0.40, hl_r.get("height", 0.42)))
            hl_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, hl_l, hl_t, hl_w, hl_h)
            hl_shape.fill.solid()
            hl_bg = parse_color_value(hl["styles"]["backgroundColor"], RGBColor(30, 41, 59) if is_dark else RGBColor(241, 245, 249), bg_blend_rgb=(slide_bg_rgb[0], slide_bg_rgb[1], slide_bg_rgb[2]))
            hl_border = parse_color_value(hl["styles"]["borderColor"], RGBColor(59, 130, 246) if is_dark else brand_color)
            hl_shape.fill.fore_color.rgb = hl_bg
            hl_shape.line.color.rgb = hl_border
            tf = hl_shape.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = hl["text"]
            p.font.name = font_family
            p.font.size = Pt(11.0)
            p.font.bold = True
            p.font.color.rgb = parse_color_value(hl["styles"]["color"], RGBColor(147, 197, 253) if is_dark else brand_color)

        # Check footnotes
        for fn in slide_geometry.get("footnotes", []):
          fn_snippet = fn["text"][:15].lower()
          if fn_snippet not in slide_text_corpus:
            print(f"[Safety Net] Appending omitted footnote: {fn['text'][:30]}...")
            fn_r = fn.get("rect", {})
            fn_l = Inches(fn_r.get("left", 0.8))
            fn_t = Inches(fn_r.get("top", 7.0))
            fn_w = Inches(fn_r.get("width", 11.733))
            fn_h = Inches(max(0.25, fn_r.get("height", 0.25)))
            fn_box = slide.shapes.add_textbox(fn_l, fn_t, fn_w, fn_h)
            tf = fn_box.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = fn["text"]
            p.font.name = font_family
            p.font.size = Pt(9.5)
            p.font.color.rgb = parse_color_value(fn["styles"].get("color"), RGBColor(148, 163, 184) if is_dark else RGBColor(100, 116, 139))


      ensure_slide_typography_consistency(slide, default_font=font_family, default_color=default_text_color, is_dark=is_dark)
      refine_card_accent_bars(slide)
      audit_and_resolve_slide_collisions(slide)
    elif len(slide.shapes) > 0:
      ensure_slide_canvas_background(slide, prs=prs, bg_color=slide_bg_rgb, is_dark=is_dark)
      ensure_slide_typography_consistency(slide, default_font=font_family, default_color=default_text_color, is_dark=is_dark)
      refine_card_accent_bars(slide)
      audit_and_resolve_slide_collisions(slide)
    else:
      print("[Warn] 'build_slide' function not found and no shapes added.")
      if slide_geometry:
        build_slide_from_geometry(slide, slide_geometry, font_name=font_family, brand_color_rgb=brand_color)
        ensure_slide_canvas_background(slide, prs=prs, bg_color=slide_bg_rgb, is_dark=is_dark)
        ensure_slide_typography_consistency(slide, default_font=font_family, default_color=default_text_color, is_dark=is_dark)
        refine_card_accent_bars(slide)
        audit_and_resolve_slide_collisions(slide)
        return
      decomp = decompose_slide_image_with_vision(fallback_img_path)
      if decomp and decomp.elements:
        build_native_pptx_slide(slide, decomp, prs, fallback_img_path)
        ensure_slide_canvas_background(slide, prs=prs, bg_color=slide_bg_rgb, is_dark=is_dark)
        ensure_slide_typography_consistency(slide, default_font=font_family, default_color=default_text_color, is_dark=is_dark)
        audit_and_resolve_slide_collisions(slide)
      else:
        slide.shapes.add_picture(
            fallback_img_path,
            Inches(0),
            Inches(0),
            width=prs.slide_width,
            height=prs.slide_height,
        )
  except Exception as e:
    print(f"[Execution Error in generated slide code]: {e}")
    if slide_geometry:
      print("[Fallback] Building native slide from exact DOM geometry on execution error.")
      build_slide_from_geometry(slide, slide_geometry, font_name=font_family, brand_color_rgb=brand_color)
      ensure_slide_canvas_background(slide, prs=prs, bg_color=slide_bg_rgb, is_dark=is_dark)
      ensure_slide_typography_consistency(slide, default_font=font_family, default_color=default_text_color, is_dark=is_dark)
      refine_card_accent_bars(slide)
      audit_and_resolve_slide_collisions(slide)
      return
    # If DOM data has table, construct native table slide
    if slide_dom_data and slide_dom_data.get("tables"):
      print("[Fallback] Building deterministic native table slide on execution error.")
      t_data = slide_dom_data["tables"][0]
      build_styled_native_table(slide, t_data, Inches(0.8), Inches(1.95), Inches(11.733), Inches(4.5))
      ensure_slide_canvas_background(slide, prs=prs, bg_color=slide_bg_rgb, is_dark=is_dark)
      ensure_slide_typography_consistency(slide, default_font=font_family, default_color=default_text_color, is_dark=is_dark)
    else:
      decomp = decompose_slide_image_with_vision(fallback_img_path)
      if decomp and decomp.elements:
        build_native_pptx_slide(slide, decomp, prs, fallback_img_path)
        ensure_slide_canvas_background(slide, prs=prs, bg_color=slide_bg_rgb, is_dark=is_dark)
        ensure_slide_typography_consistency(slide, default_font=font_family, default_color=default_text_color, is_dark=is_dark)
      else:
        slide.shapes.add_picture(
            fallback_img_path,
            Inches(0),
            Inches(0),
            width=prs.slide_width,
            height=prs.slide_height,
        )


def build_native_pptx_slide(
    slide,
    decomposition: Optional[SlideNativeDecomposition],
    prs: Presentation,
    fallback_img_path: str,
):
  """Builds 100% native editable shapes, text boxes, and tables on a slide, or falls back to picture."""
  if not decomposition or not decomposition.elements:
    slide.shapes.add_picture(
        fallback_img_path,
        Inches(0),
        Inches(0),
        width=prs.slide_width,
        height=prs.slide_height,
    )
    return

  slide_w = prs.slide_width
  slide_h = prs.slide_height

  # 1. Slide Background
  bg_color = _hex_to_pptx_color(
      decomposition.slide_background_hex, default=(255, 255, 255)
  )
  bg_shape = slide.shapes.add_shape(
      MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), slide_w, slide_h
  )
  bg_shape.fill.solid()
  bg_shape.fill.fore_color.rgb = bg_color
  bg_shape.line.fill.background()

  def _in_x(pct: float) -> Inches:
    return Inches(pct / 100.0 * 13.333333)

  def _in_y(pct: float) -> Inches:
    return Inches(pct / 100.0 * 7.5)

  # 2. Containers / Cards (Background layers first)
  for elem in decomposition.elements:
    if elem.element_type == "CONTAINER":
      left = _in_x(elem.left_pct)
      top = _in_y(elem.top_pct)
      width = _in_x(elem.width_pct)
      height = _in_y(elem.height_pct)

      shape_type = (
          MSO_SHAPE.ROUNDED_RECTANGLE
          if (elem.border_radius_pt or 0) > 0
          else MSO_SHAPE.RECTANGLE
      )
      shape = slide.shapes.add_shape(shape_type, left, top, width, height)

      if elem.bg_color_hex:
        shape.fill.solid()
        shape.fill.fore_color.rgb = _hex_to_pptx_color(elem.bg_color_hex)
      else:
        shape.fill.background()

      if elem.border_color_hex:
        shape.line.color.rgb = _hex_to_pptx_color(elem.border_color_hex)
        shape.line.width = Pt(elem.border_width_pt or 1.0)
      else:
        shape.line.fill.background()

      if elem.accent_bar_color_hex:
        bar_w = width - Inches(0.40)
        if bar_w > Inches(0.5):
          bar_shape = slide.shapes.add_shape(
              MSO_SHAPE.ROUNDED_RECTANGLE,
              left + Inches(0.20),
              top + Inches(0.08),
              bar_w,
              Inches(0.06),
          )
          bar_shape.fill.solid()
          bar_shape.fill.fore_color.rgb = _hex_to_pptx_color(elem.accent_bar_color_hex)
          bar_shape.line.fill.background()

  # 3. Dividers
  for elem in decomposition.elements:
    if elem.element_type == "DIVIDER":
      left = _in_x(elem.left_pct)
      top = _in_y(elem.top_pct)
      width = _in_x(elem.width_pct)
      height = _in_y(max(elem.height_pct, 0.2))

      line_shape = slide.shapes.add_shape(
          MSO_SHAPE.RECTANGLE, left, top, width, height
      )
      line_color = _hex_to_pptx_color(elem.border_color_hex or "#E2E8F0")
      line_shape.fill.solid()
      line_shape.fill.fore_color.rgb = line_color
      line_shape.line.fill.background()

  # 4. Sub-cards, Code Blocks, Callouts, Pipeline Steps
  for elem in decomposition.elements:
    if elem.element_type in ("SUB_CARD", "CODE_BLOCK", "CALLOUT_BOX", "PIPELINE_STEP"):
      left = _in_x(elem.left_pct)
      top = _in_y(elem.top_pct)
      width = _in_x(elem.width_pct)
      height = _in_y(elem.height_pct)

      shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
      if elem.bg_color_hex:
        shape.fill.solid()
        shape.fill.fore_color.rgb = _hex_to_pptx_color(elem.bg_color_hex)
      else:
        shape.fill.background()

      if elem.border_color_hex:
        shape.line.color.rgb = _hex_to_pptx_color(elem.border_color_hex)
        shape.line.width = Pt(elem.border_width_pt or 1.0)
      else:
        shape.line.fill.background()

      if elem.text and elem.text.strip():
        tf = shape.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.08)
        tf.margin_right = Inches(0.08)
        tf.margin_top = Inches(0.04)
        tf.margin_bottom = Inches(0.04)
        for p_idx, line in enumerate(elem.text.split("\n")):
          p = tf.paragraphs[0] if p_idx == 0 else tf.add_paragraph()
          p.text = line
          p.font.name = "Consolas" if elem.element_type == "CODE_BLOCK" else "Noto Sans KR"
          p.font.size = Pt(elem.font_size_pt or (9.0 if elem.element_type == "CODE_BLOCK" else 10.5))
          p.font.bold = bool(elem.is_bold) or (p_idx == 0 and elem.element_type != "CODE_BLOCK")
          if elem.font_color_hex:
            p.font.color.rgb = _hex_to_pptx_color(elem.font_color_hex)

  # 5. Tables
  for elem in decomposition.elements:
    if elem.element_type == "TABLE" and elem.table_matrix:
      matrix = elem.table_matrix
      rows = len(matrix)
      cols = max(len(r) for r in matrix) if rows > 0 else 0
      if rows > 0 and cols > 0:
        left = _in_x(elem.left_pct)
        top = _in_y(elem.top_pct)
        width = _in_x(elem.width_pct)
        height = _in_y(elem.height_pct)

        table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
        tbl = table_shape.table
        for r_idx, row in enumerate(matrix):
          for c_idx, cell_val in enumerate(row):
            if c_idx < cols:
              cell = tbl.cell(r_idx, c_idx)
              cell.text = str(cell_val).strip()
              for p in cell.text_frame.paragraphs:
                p.font.name = "Noto Sans KR"
                p.font.size = Pt(11.0 if r_idx > 0 else 12.0)
                p.font.bold = (r_idx == 0)
                p.font.color.rgb = _hex_to_pptx_color(
                    "#1E293B" if r_idx == 0 else "#0F172A"
                )
              if r_idx == 0 and elem.table_header_bg_hex:
                cell.fill.solid()
                cell.fill.fore_color.rgb = _hex_to_pptx_color(
                    elem.table_header_bg_hex
                )

  # 6. Flow Nodes, Flow Arrows, Bridge Connectors, Badges
  for elem in decomposition.elements:
    if elem.element_type in ("FLOW_NODE", "FLOW_ARROW", "BRIDGE_CONNECTOR", "HEADER_BADGE", "ICON_BADGE"):
      left = _in_x(elem.left_pct)
      top = _in_y(elem.top_pct)
      width = _in_x(elem.width_pct)
      height = _in_y(elem.height_pct)

      shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
      if elem.bg_color_hex:
        shape.fill.solid()
        shape.fill.fore_color.rgb = _hex_to_pptx_color(elem.bg_color_hex)
      else:
        shape.fill.background()

      if elem.border_color_hex:
        shape.line.color.rgb = _hex_to_pptx_color(elem.border_color_hex)
        shape.line.width = Pt(elem.border_width_pt or 1.0)
      else:
        shape.line.fill.background()

      if elem.text and elem.text.strip():
        tf = shape.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.04)
        tf.margin_right = Inches(0.04)
        tf.margin_top = Inches(0.02)
        tf.margin_bottom = Inches(0.02)
        p = tf.paragraphs[0]
        p.text = elem.text.strip()
        p.font.name = "Noto Sans KR"
        p.font.size = Pt(elem.font_size_pt or 10.0)
        p.font.bold = bool(elem.is_bold)
        p.font.color.rgb = _hex_to_pptx_color(elem.font_color_hex or "#FFFFFF")
        if elem.alignment == "CENTER":
          p.alignment = PP_ALIGN.CENTER
        elif elem.alignment == "RIGHT":
          p.alignment = PP_ALIGN.RIGHT

  # 7. Text Boxes, Stat Metrics, and Footnotes
  for elem in decomposition.elements:
    if elem.element_type in ("TEXT_BOX", "STAT_METRIC", "FOOTNOTE") and elem.text and elem.text.strip():
      left = _in_x(elem.left_pct)
      top = _in_y(elem.top_pct)
      width = _in_x(elem.width_pct)
      height = _in_y(elem.height_pct)

      txBox = slide.shapes.add_textbox(left, top, width, height)
      tf = txBox.text_frame
      tf.word_wrap = True
      tf.margin_left = Inches(0.04)
      tf.margin_right = Inches(0.04)
      tf.margin_top = Inches(0.02)
      tf.margin_bottom = Inches(0.02)

      lines = [l for l in elem.text.split("\n") if l.strip()]
      for p_idx, line in enumerate(lines):
        p = tf.paragraphs[0] if p_idx == 0 else tf.add_paragraph()
        p.text = line.strip()
        p.font.name = "Noto Sans KR"
        p.font.size = Pt(elem.font_size_pt or (9.5 if elem.element_type == "FOOTNOTE" else 14.0))
        p.font.bold = bool(elem.is_bold)
        p.font.color.rgb = _hex_to_pptx_color(
            elem.font_color_hex or ("#64748B" if elem.element_type == "FOOTNOTE" else "#0F172A")
        )
        if elem.alignment == "CENTER":
          p.alignment = PP_ALIGN.CENTER
        elif elem.alignment == "RIGHT":
          p.alignment = PP_ALIGN.RIGHT
        else:
          p.alignment = PP_ALIGN.LEFT


def _run_async_in_thread(coro_fn, *args, **kwargs):
  """Safely executes an async coroutine even if invoked from within a running asyncio event loop (e.g. FastAPI / ADK)."""
  try:
    loop = asyncio.get_running_loop()
  except RuntimeError:
    loop = None

  if loop and loop.is_running():
    with concurrent.futures.ThreadPoolExecutor(max_workers=1) as executor:
      future = executor.submit(lambda: asyncio.run(coro_fn(*args, **kwargs)))
      return future.result()
  else:
    return asyncio.run(coro_fn(*args, **kwargs))


def capture_html_slides(
    html_input: str,
    output_dir: Optional[str] = None,
    target_slide_id: Optional[str] = None,
    base_dir: Optional[str] = None,
    scale_factor: int = 2,
    chrome_binary: Optional[str] = None,
) -> Dict[str, Any]:
  """Renders HTML presentation slides to 16:9 widescreen PNG images via headless Chrome CDP.

  Args:
      html_input: File path to .html or raw HTML string.
      output_dir: Directory where captured PNG images will be stored. Defaults
        to 'captures/'.
      target_slide_id: Optional slide ID or number to capture only that slide.
      base_dir: Base directory for resolving relative assets.
      scale_factor: Device scale factor (default 2 for Retina quality).
      chrome_binary: Path to Chrome executable.

  Returns:
      Dict with:
          - 'success': bool
          - 'slide_count': int
          - 'slide_ids': List[str]
          - 'image_paths': List[str] (absolute file paths)
          - 'image_data_uris': List[str] (data:image/png;base64,...)
          - 'output_dir': str
  """
  # 1. Determine whether html_input is a path or raw string
  if os.path.isfile(html_input):
    html_path = os.path.abspath(html_input)
    if not base_dir:
      base_dir = os.path.dirname(html_path)
    base_name = os.path.splitext(os.path.basename(html_path))[0]
    with open(html_path, "r", encoding="utf-8") as f:
      raw_html = f.read()
  else:
    raw_html = html_input
    base_name = "slide"
    if not base_dir:
      base_dir = os.getcwd()

  if not output_dir:
    output_dir = os.path.join(
        os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..")),
        "captures",
    )
  os.makedirs(output_dir, exist_ok=True)

  # 2. Extract slide IDs dynamically
  all_slide_ids = _extract_slide_ids(raw_html)
  if not all_slide_ids:
    raise ValueError("No slide elements detected in the provided HTML content.")

  if target_slide_id:
    target_clean = str(target_slide_id).strip().lower()
    matched = []
    for idx, sid in enumerate(all_slide_ids):
      sid_clean = sid.lower()
      num_match = re.search(r"\d+", sid_clean)
      num_str = num_match.group(0) if num_match else ""
      if (
          target_clean == sid_clean
          or target_clean == f"slide-{sid_clean}"
          or target_clean == num_str
          or target_clean == str(idx + 1)
      ):
        matched.append((idx + 1, sid))
    target_tuples = matched if matched else list(enumerate(all_slide_ids, start=1))
  else:
    target_tuples = list(enumerate(all_slide_ids, start=1))

  total_slides = len(target_tuples)
  slide_ids = [t[1] for t in target_tuples]

  with tempfile.TemporaryDirectory() as tmpdir:
    # 3. Generate individual temporary HTML files for each slide
    temp_slide_files = []
    for orig_idx, sid in target_tuples:
      slide_html = _prepare_slide_html(raw_html, sid, orig_idx, base_dir)
      temp_file = os.path.join(tmpdir, f"temp_{sid}.html")
      with open(temp_file, "w", encoding="utf-8") as f:
        f.write(slide_html)
      temp_slide_files.append((sid, temp_file))

    # 4. Start Chrome CDP on an available ephemeral port
    port = _find_free_port()
    profile_dir = os.path.join(tmpdir, "chrome_profile")
    os.makedirs(profile_dir, exist_ok=True)

    actual_chrome = _find_chrome_binary(chrome_binary)
    chrome_cmd = [
        actual_chrome,
        "--headless=new",
        f"--remote-debugging-port={port}",
        "--no-sandbox",
        "--disable-gpu",
        "--disable-dev-shm-usage",
        "--allow-file-access-from-files",
        "--disable-web-security",
        "--ignore-certificate-errors",
        f"--user-data-dir={profile_dir}",
        "--window-size=1920,1080",
        "about:blank",
    ]

    chrome_proc = subprocess.Popen(
        chrome_cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
    )

    try:
      captured_images, captured_geometries = _run_async_in_thread(
          _capture_slides_cdp,
          temp_slide_files,
          port,
          profile_dir,
          output_dir,
          scale_factor=scale_factor,
      )
    finally:
      try:
        chrome_proc.terminate()
        chrome_proc.wait(timeout=3)
      except Exception:
        try:
          chrome_proc.kill()
        except Exception:
          pass
      time.sleep(0.2)

  # Read Base64 for each captured image
  image_data_uris = []
  for img_p in captured_images:
    with open(img_p, "rb") as f_img:
      b64 = base64.b64encode(f_img.read()).decode("utf-8")
      image_data_uris.append(f"data:image/png;base64,{b64}")

  return {
      "success": True,
      "slide_count": total_slides,
      "slide_ids": slide_ids,
      "image_paths": captured_images,
      "slide_geometries": captured_geometries,
      "image_data_uris": image_data_uris,
      "output_dir": output_dir,
  }


def harmonize_deck_presentation_fidelity(
    prs: Presentation,
    default_font: str = "Pretendard",
    default_bg_hex: Optional[str] = None,
) -> None:
  """Guarantees 100% deck-wide typography consistency, canvas background fill, font floors, and non-destructive accent bar safety."""
  is_dark = False
  deck_bg_rgb = RGBColor(255, 255, 255)
  if default_bg_hex:
    bg_tuple = _hex_to_rgb_tuple(default_bg_hex, (255, 255, 255))
    deck_bg_rgb = RGBColor(*bg_tuple)
    is_dark = (bg_tuple[0] * 0.299 + bg_tuple[1] * 0.587 + bg_tuple[2] * 0.114) < 128
  elif len(prs.slides) > 0 and len(prs.slides[0].shapes) > 0:
    for sh in prs.slides[0].shapes:
      if sh.shape_type == 1 and sh.left.inches <= 0.1 and sh.top.inches <= 0.1 and sh.width.inches >= 13.0:
        try:
          col = sh.fill.fore_color.rgb
          deck_bg_rgb = col
          is_dark = (col[0] * 0.299 + col[1] * 0.587 + col[2] * 0.114) < 128
          break
        except Exception:
          pass

  default_color = RGBColor(241, 245, 249) if is_dark else RGBColor(17, 17, 21)

  for slide in prs.slides:
    ensure_slide_canvas_background(slide, prs=prs, bg_color=deck_bg_rgb, is_dark=is_dark)
    ensure_slide_typography_consistency(slide, default_font=default_font, default_color=default_color, is_dark=is_dark)
    refine_card_accent_bars(slide)
    audit_and_resolve_slide_collisions(slide)

    for s in slide.shapes:
      if s.has_text_frame and not s.has_table:
        for p in s.text_frame.paragraphs:
          if not p.font.name:
            p.font.name = default_font
          if not p.font.size:
            p.font.size = Pt(10.5)
          for r in p.runs:
            if not r.font.name:
              r.font.name = default_font
            if not r.font.size:
              r.font.size = p.font.size


def convert_html_to_pptx(
    html_input: str,
    output_pptx_path: Optional[str] = None,
    target_slide_id: Optional[str] = None,
    base_dir: Optional[str] = None,
    scale_factor: int = 2,
    native_mode: bool = True,
    chrome_binary: Optional[str] = None,
) -> Dict[str, Any]:
  """Dynamically converts any multi-slide HTML file or HTML string to a 16:9 widescreen PowerPoint presentation.

  Uses Chrome Headless to capture high-res screenshots, then leverages Gemini 3.8 Flash
  OCR, CSS Token resolution & Layout Decomposition to construct 100% native editable shapes, textboxes, and tables.

  Args:
      html_input: File path to .html file, or raw HTML markup string.
      output_pptx_path: Destination path for .pptx file. Defaults to [html_name].pptx.
      target_slide_id: Optional single slide ID or number (e.g. 'slide-3', '3') to convert only that slide.
      base_dir: Directory to resolve relative image/svg assets. Defaults to HTML file dir.
      scale_factor: Device scale factor for screenshot capture (default 2 for Retina).
      native_mode: If True, uses Gemini 3.8 Flash to decompose into 100% native editable PowerPoint shapes, text boxes, and tables.
      chrome_binary: Path to Chrome or Chromium executable.

  Returns:
      Dictionary with conversion metadata:
          - 'success': bool
          - 'slide_count': int
          - 'slide_ids': List[str]
          - 'output_pptx_path': str
          - 'image_paths': List[str]
          - 'file_size_bytes': int
  """
  # 1. Determine whether html_input is a path or raw string
  if os.path.isfile(html_input):
    html_path = os.path.abspath(html_input)
    if not base_dir:
      base_dir = os.path.dirname(html_path)
    if not output_pptx_path:
      base_name = os.path.splitext(os.path.basename(html_path))[0]
      output_pptx_path = os.path.join(base_dir, f"{base_name}.pptx")
    with open(html_path, "r", encoding="utf-8") as f:
      raw_html_content = f.read()
  else:
    raw_html_content = html_input
    if not base_dir:
      base_dir = os.getcwd()
    if not output_pptx_path:
      output_pptx_path = os.path.join(base_dir, "slide.pptx")

  # 2. Capture slides using headless Chrome
  with tempfile.TemporaryDirectory() as tmp_captures:
    cap_res = capture_html_slides(
        html_input=html_input,
        output_dir=tmp_captures,
        target_slide_id=target_slide_id,
        base_dir=base_dir,
        scale_factor=scale_factor,
        chrome_binary=chrome_binary,
    )
    all_captured_images = cap_res["image_paths"]
    all_captured_geometries = cap_res.get("slide_geometries", [])
    all_slide_ids = cap_res["slide_ids"]
    total_deck_slides = cap_res["slide_count"]

    # 3. Extract CSS custom properties and parse HTML DOM
    css_vars_raw = extract_root_css_vars(raw_html_content)
    tagged_html_content, _ = _tag_slide_ids_in_html(raw_html_content)
    try:
      doc = lxml.html.fromstring(tagged_html_content)
    except Exception:
      doc = None

    # Filter for target slide if specified
    active_indices = list(range(len(all_slide_ids)))
    if target_slide_id:
      target_clean = str(target_slide_id).strip().lower()
      matched = []
      for idx, sid in enumerate(all_slide_ids):
        sid_clean = sid.lower()
        num_match = re.search(r"\d+", sid_clean)
        num_str = num_match.group(0) if num_match else ""
        if (
            target_clean == sid_clean
            or target_clean == f"slide-{sid_clean}"
            or target_clean == num_str
            or target_clean == str(idx + 1)
        ):
          matched.append(idx)
      if matched:
        active_indices = matched

    captured_images = [all_captured_images[i] for i in active_indices]
    captured_geometries = [all_captured_geometries[i] if i < len(all_captured_geometries) else None for i in active_indices]
    slide_ids = [all_slide_ids[i] for i in active_indices]
    total_slides = len(captured_images)

    # Parse slide DOM semantic data for each active slide
    slide_dom_list = []
    for sid in slide_ids:
      s_dom = None
      if doc is not None:
        nodes = doc.xpath(f"//*[@id='{sid}']")
        if nodes:
          s_dom = parse_slide_semantic_data(nodes[0], css_vars_raw)
      slide_dom_list.append(s_dom)

    global_tokens = extract_global_design_tokens(
        first_image_path=captured_images[0] if captured_images else "",
        raw_html=raw_html_content,
    )

    # 4. Build 16:9 Widescreen PowerPoint Presentation
    prs = Presentation()
    prs.slide_width = Inches(13.333333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    slide_codes = [None] * len(captured_images)
    if native_mode and captured_images:
      max_workers = min(len(captured_images), 2)
      with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {
            executor.submit(
                generate_native_slide_builder_code,
                img_path,
                slide_dom_data=slide_dom_list[idx],
                slide_geometry=captured_geometries[idx],
                design_tokens=global_tokens,
                slide_index=active_indices[idx] + 1,
                total_slides=total_deck_slides,
            ): idx
            for idx, img_path in enumerate(captured_images)
        }
        for future in concurrent.futures.as_completed(futures):
          idx = futures[future]
          try:
            slide_codes[idx] = future.result()
          except Exception as e:
            print(f"[Semantic CodeGen Error for Slide {active_indices[idx]+1}]: {e}")
            slide_codes[idx] = None

    for idx, img_path in enumerate(captured_images):
      slide = prs.slides.add_slide(blank_layout)
      build_native_slide_from_code(
          slide,
          prs=prs,
          code_str=slide_codes[idx],
          fallback_img_path=img_path,
          design_tokens=global_tokens,
          slide_dom_data=slide_dom_list[idx],
          slide_geometry=captured_geometries[idx],
      )

    # 5. Harmonize deck-wide spatial anchors, table fidelity, and typography floors
    harmonize_deck_presentation_fidelity(
        prs,
        default_font=global_tokens.get("font_name", "Pretendard"),
        default_bg_hex=global_tokens.get("bg_color_hex"),
    )

    # Ensure target directory exists
    os.makedirs(os.path.dirname(os.path.abspath(output_pptx_path)), exist_ok=True)
    prs.save(output_pptx_path)

    file_size = os.path.getsize(output_pptx_path)

    return {
        "success": True,
        "slide_count": total_slides,
        "slide_ids": slide_ids,
        "output_pptx_path": output_pptx_path,
        "file_size_bytes": file_size,
    }


if __name__ == "__main__":
  import argparse

  parser = argparse.ArgumentParser(
      description="Dynamically convert an HTML slide deck to PPTX with Native Editable Vision Decomposition."
  )
  parser.add_argument(
      "input_html", help="Path to the input HTML presentation file"
  )
  parser.add_argument(
      "-o", "--output", help="Path to output PPTX file (optional)"
  )
  parser.add_argument(
      "--scale",
      type=int,
      default=2,
      help="Device scale factor (default: 2 for Retina)",
  )
  parser.add_argument(
      "--native",
      action="store_true",
      default=True,
      help="Use Gemini Vision to decompose into native editable shapes (default: True)",
  )

  args = parser.parse_args()
  result = convert_html_to_pptx(
      args.input_html,
      output_pptx_path=args.output,
      scale_factor=args.scale,
      native_mode=args.native,
  )
  print(f"[SUCCESS] Converted {result['slide_count']} slides to PPTX:")
  print(f"File: {result['output_pptx_path']} ({result['file_size_bytes']:,} bytes)")
