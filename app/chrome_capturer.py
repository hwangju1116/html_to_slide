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
import websockets

from app.js_geometry_extractor import JS_SLIDE_GEOMETRY_EXTRACTOR

FONTS_DIR = os.path.join(
    os.path.dirname(os.path.dirname(__file__)), "assets", "fonts"
)


FALLBACK_FONTS_DIR = os.path.expanduser("~/.local/share/fonts")


def _find_free_port() -> int:
  """Finds an available TCP port dynamically."""
  with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
    s.bind(("", 0))
    s.listen(1)
    port = s.getsockname()[1]
    return port


def _get_font_path(font_filename: str) -> str:
  """Resolves the font path from package assets, local user font cache, or container font cache."""
  candidates = [
      os.path.join(FONTS_DIR, font_filename),
      os.path.join(FALLBACK_FONTS_DIR, "pretendard", font_filename),
      os.path.join(FALLBACK_FONTS_DIR, font_filename),
      os.path.join("/usr/local/share/fonts/pretendard", font_filename),
  ]
  for candidate in candidates:
    if os.path.exists(candidate):
      return candidate
  return candidates[0]


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

  # 1. Clean out ALL external blocking links/imports and rewrite remote tailwind/chart.js scripts to local assets
  html = re.sub(
      r"@import\s+url\(\s*['\"]?https?://[^)]+\)\s*;?",
      "",
      html,
      flags=re.IGNORECASE,
  )
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
  local_chartjs = os.path.abspath(
      os.path.join(os.path.dirname(__file__), "..", "assets", "chart.min.js")
  )
  chart_shim = (
      f'<script src="file://{local_chartjs}"></script>\n'
      "<script>\n"
      "if (window.Chart) {\n"
      "  Chart.defaults.animation = false;\n"
      "  Chart.defaults.animations = { colors: false, x: false, y: false };\n"
      "  Chart.defaults.transitions = { active: { animation: { duration: 0 } }, resize: { animation: { duration: 0 } } };\n"
      "  const _OrigChart = window.Chart;\n"
      "  window.Chart = function(ctx, cfg) {\n"
      "    if (cfg) {\n"
      "      cfg.options = cfg.options || {};\n"
      "      cfg.options.animation = false;\n"
      "      cfg.options.animations = { colors: false, x: false, y: false };\n"
      "      const canvasEl = (ctx && ctx.canvas) ? ctx.canvas : (ctx instanceof HTMLElement ? ctx : null);\n"
      "      if (canvasEl) {\n"
      "        try { canvasEl.__chartConfig = JSON.parse(JSON.stringify(cfg)); } catch (e) { canvasEl.__chartConfig = cfg; }\n"
      "      }\n"
      "    }\n"
      "    return new _OrigChart(ctx, cfg);\n"
      "  };\n"
      "  Object.assign(window.Chart, _OrigChart);\n"
      "  window.Chart.prototype = _OrigChart.prototype;\n"
      "}\n"
      "</script>"
  )
  if os.path.exists(local_chartjs):
    html = re.sub(
        r'<script[^>]*src=["\']https?://[^"\']*chart(?:\.umd|\.min)?(?:\.js)?[^"\']*["\'][^>]*>\s*</script>',
        chart_shim,
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

  # 3. Resolve fonts (official Pretendard + Noto Sans KR fallbacks)
  regular_font = _get_font_path("NotoSansKR-Regular.otf")
  medium_font = _get_font_path("NotoSansKR-Medium.otf")
  bold_font = _get_font_path("NotoSansKR-Bold.otf")
  black_font = _get_font_path("NotoSansKR-Black.otf")

  pret_reg = _get_font_path("Pretendard-Regular.otf")
  pret_med = _get_font_path("Pretendard-Medium.otf")
  pret_semi = _get_font_path("Pretendard-SemiBold.otf")
  pret_bold = _get_font_path("Pretendard-Bold.otf")
  pret_xbold = _get_font_path("Pretendard-ExtraBold.otf")
  pret_black = _get_font_path("Pretendard-Black.otf")

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

  reg_url = f", url('file://{regular_font}') format('opentype')" if os.path.exists(regular_font) else ""
  med_url = f", url('file://{medium_font}') format('opentype')" if os.path.exists(medium_font) else reg_url
  bold_url = f", url('file://{bold_font}') format('opentype')" if os.path.exists(bold_font) else reg_url
  black_url = f", url('file://{black_font}') format('opentype')" if os.path.exists(black_font) else bold_url

  p_reg_url = f", url('file://{pret_reg}') format('opentype')" if os.path.exists(pret_reg) else reg_url
  p_med_url = f", url('file://{pret_med}') format('opentype')" if os.path.exists(pret_med) else med_url
  p_semi_url = f", url('file://{pret_semi}') format('opentype')" if os.path.exists(pret_semi) else p_med_url
  p_bold_url = f", url('file://{pret_bold}') format('opentype')" if os.path.exists(pret_bold) else bold_url
  p_xbold_url = f", url('file://{pret_xbold}') format('opentype')" if os.path.exists(pret_xbold) else p_bold_url
  p_black_url = f", url('file://{pret_black}') format('opentype')" if os.path.exists(pret_black) else black_url

  font_face_css = f"""
    @font-face {{
        font-family: 'Pretendard';
        font-weight: 400;
        src: local('Pretendard'), local('Pretendard Regular'){p_reg_url}, local('Noto Sans KR'), local('Noto Sans CJK KR'), local('Noto Sans Korean'), local('Apple SD Gothic Neo'), local('Malgun Gothic');
    }}
    @font-face {{
        font-family: 'Pretendard';
        font-weight: 500;
        src: local('Pretendard Medium'){p_med_url}, local('Noto Sans KR Medium'), local('Noto Sans CJK KR Medium'), local('Noto Sans Korean Medium'), local('Apple SD Gothic Neo'), local('Malgun Gothic');
    }}
    @font-face {{
        font-family: 'Pretendard';
        font-weight: 600;
        src: local('Pretendard SemiBold'){p_semi_url}, local('Pretendard Medium'), local('Noto Sans KR Medium'), local('Noto Sans CJK KR Medium'), local('Apple SD Gothic Neo Bold'), local('Malgun Gothic Bold');
    }}
    @font-face {{
        font-family: 'Pretendard';
        font-weight: 700;
        src: local('Pretendard Bold'){p_bold_url}, local('Noto Sans KR Bold'), local('Noto Sans CJK KR Bold'), local('Noto Sans Korean Bold'), local('Apple SD Gothic Neo Bold'), local('Malgun Gothic Bold');
    }}
    @font-face {{
        font-family: 'Pretendard';
        font-weight: 800;
        src: local('Pretendard ExtraBold'){p_xbold_url}, local('Pretendard Bold'), local('Noto Sans KR Bold'), local('Noto Sans CJK KR Bold');
    }}
    @font-face {{
        font-family: 'Pretendard';
        font-weight: 900;
        src: local('Pretendard Black'){p_black_url}, local('Noto Sans KR Black'), local('Noto Sans CJK KR Black');
    }}
    @font-face {{
        font-family: 'Noto Sans KR';
        font-weight: 400;
        src: local('Noto Sans KR'), local('Noto Sans CJK KR'), local('Noto Sans Korean'){reg_url}, local('Pretendard'), local('Apple SD Gothic Neo'), local('Malgun Gothic');
    }}
    @font-face {{
        font-family: 'Noto Sans KR';
        font-weight: 500;
        src: local('Noto Sans KR Medium'), local('Noto Sans CJK KR Medium'), local('Noto Sans Korean Medium'){med_url}, local('Pretendard Medium'), local('Apple SD Gothic Neo'), local('Malgun Gothic');
    }}
    @font-face {{
        font-family: 'Noto Sans KR';
        font-weight: 700;
        src: local('Noto Sans KR Bold'), local('Noto Sans CJK KR Bold'), local('Noto Sans Korean Bold'){bold_url}, local('Pretendard Bold'), local('Apple SD Gothic Neo Bold'), local('Malgun Gothic Bold');
    }}
    @font-face {{
        font-family: 'Noto Sans KR';
        font-weight: 900;
        src: local('Noto Sans KR Black'), local('Noto Sans CJK KR Black'), local('Noto Sans Korean Black'){black_url}, local('Pretendard Black'), local('Apple SD Gothic Neo Bold'), local('Malgun Gothic Bold');
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
        font-family: 'Pretendard', 'Noto Sans KR', 'Noto Sans CJK KR', 'Noto Sans Korean', 'Apple SD Gothic Neo', 'Malgun Gothic', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
        -webkit-font-smoothing: antialiased;
    }}
    header:not(.slide-header), footer:not(.slide-footer), nav, [role="navigation"], .chrome, .tabs, .stage-controls, .controls-footer, div:not(.slide):not(:has(#{target_slide_id})):has(> button), div:not(.slide):not(:has(#{target_slide_id})):has(> div > button) {{
        display: none !important;
    }}
    main, .stage, .deck-container, .preview-container, body > div:has(#{target_slide_id}) {{
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
    .frame, #frame, .slide-frame, div:has(> #{target_slide_id}), main:has(> #{target_slide_id}), section:has(> #{target_slide_id}) {{
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

  # Override JS slide state if present (0-based for currentSlide/idx)
  zero_idx = max(0, slide_index - 1)
  html = re.sub(
      r"(let|var|const)\s+(currentSlide|currentSlideIndex|activeIndex|activeSlide|current|slideIdx|idx)\s*=\s*\d+;",
      r"\1 \2 = " + str(zero_idx) + ";",
      html,
  )
  html = re.sub(
      r"\b(updateSlide|showSlide|goToSlide|setSlide|renderSlide)\(\s*0\s*\)",
      rf"\1({zero_idx})",
      html,
  )

  return html


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
        orig_num_match = re.search(r"\d+", str(sid))
        slide_idx_0 = max(0, int(orig_num_match.group(0)) - 1) if orig_num_match else (idx - 1)
      else:
        sid, slide_file = f"slide-{idx}", item
        slide_idx_0 = idx - 1

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

      # Extract exact browser-measured DOM geometry directly from rendered layout (and sync Chart.js / FA icons)
      geom = None
      try:
        js = JS_SLIDE_GEOMETRY_EXTRACTOR.replace("%SLIDE_ID%", str(sid)).replace("%SLIDE_IDX_0%", str(slide_idx_0))
        layout_res = await send_recv("Runtime.evaluate", {"expression": js, "returnByValue": True})
        geom = layout_res.get("result", {}).get("value")
      except Exception as geom_err:
        print(f"[CDP Layout Warning for {sid}]: {geom_err}")
      captured_geometries.append(geom)

      res = await send_recv("Page.captureScreenshot", {"format": "png"})
      if "data" in res:
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


