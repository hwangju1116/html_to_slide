import logging
import os
import re
import shutil
import subprocess
import urllib.request
from typing import Any, Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)

SKILL_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BUNDLED_FONTS_DIR = os.path.join(SKILL_ROOT, "assets", "fonts")
BUNDLED_ASSETS_DIR = os.path.join(SKILL_ROOT, "assets")
CACHE_ROOT = os.environ.get(
    "HTML_TO_SLIDE_CACHE_DIR", os.path.expanduser("~/.cache/html_to_slide")
)
FONT_CACHE_DIR = os.path.join(CACHE_ROOT, "fonts")
ASSET_CACHE_DIR = os.path.join(CACHE_ROOT, "assets")
OFFLINE = os.environ.get("HTML_TO_SLIDE_OFFLINE", "").lower() in ("1", "true", "yes")

LAST_FONT_REPORT: Dict[str, Any] = {}
CJK_FALLBACK_FAMILIES = ("HTS CJK KR", "HTS CJK JP", "HTS CJK SC")
CJK_FALLBACK_SUFFIX = ", " + ", ".join(f"'{f}'" for f in CJK_FALLBACK_FAMILIES)

_HANGUL_RANGE = "U+1100-11FF, U+3130-318F, U+A960-A97F, U+AC00-D7AF, U+D7B0-D7FF"
_KANA_RANGE = "U+3040-309F, U+30A0-30FF, U+31F0-31FF, U+FF66-FF9F"
_HAN_RANGE = (
    "U+2E80-2FDF, U+3000-303F, U+3400-4DBF, U+4E00-9FFF, U+F900-FAFF, "
    "U+FF00-FF65, U+FFE0-FFEE, U+20000-2A6DF"
)

CJK_FONT_SOURCES: Dict[str, Dict[str, Any]] = {
    "kr": {
        "family": "HTS CJK KR",
        "fc_lang": "ko",
        "bundled": ["NotoSansKR-Regular.otf", "NotoSansKR-Bold.otf"],
        "bundled_weights": [400, 700],
        "cache_file": "NotoSansKR[wght].ttf",
        "url": "https://raw.githubusercontent.com/google/fonts/main/ofl/notosanskr/NotoSansKR%5Bwght%5D.ttf",
        "system_families": [
            "Noto Sans CJK KR",
            "Noto Sans KR",
            "Pretendard",
            "Malgun Gothic",
            "Apple SD Gothic Neo",
            "NanumGothic",
        ],
        "ranges": _HANGUL_RANGE,
    },
    "jp": {
        "family": "HTS CJK JP",
        "fc_lang": "ja",
        "bundled": [],
        "bundled_weights": [],
        "cache_file": "NotoSansJP[wght].ttf",
        "url": "https://raw.githubusercontent.com/google/fonts/main/ofl/notosansjp/NotoSansJP%5Bwght%5D.ttf",
        "system_families": [
            "Noto Sans CJK JP",
            "Noto Sans JP",
            "Yu Gothic",
            "Hiragino Sans",
            "Meiryo",
        ],
        "ranges": _KANA_RANGE,
    },
    "sc": {
        "family": "HTS CJK SC",
        "fc_lang": "zh-cn",
        "bundled": [],
        "bundled_weights": [],
        "cache_file": "NotoSansSC[wght].ttf",
        "url": "https://raw.githubusercontent.com/google/fonts/main/ofl/notosanssc/NotoSansSC%5Bwght%5D.ttf",
        "system_families": [
            "Noto Sans CJK SC",
            "Noto Sans SC",
            "Microsoft YaHei",
            "PingFang SC",
            "WenQuanYi Micro Hei",
        ],
        "ranges": _HAN_RANGE,
    },
}

JS_ASSET_SOURCES: Dict[str, Dict[str, str]] = {
    "tailwindcss.min.js": {"url": "https://cdn.tailwindcss.com/3.4.17"},
    "chart.min.js": {
        "url": "https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"
    },
}


def _download(url: str, dest: str, timeout: float = 60.0) -> bool:
  if OFFLINE:
    return False
  os.makedirs(os.path.dirname(dest), exist_ok=True)
  tmp = dest + ".part"
  try:
    req = urllib.request.Request(url, headers={"User-Agent": "html-to-slide/1.0"})
    with urllib.request.urlopen(req, timeout=timeout) as resp, open(tmp, "wb") as fh:
      shutil.copyfileobj(resp, fh)
    if os.path.getsize(tmp) < 1024:
      os.remove(tmp)
      return False
    os.replace(tmp, dest)
    logger.info("Downloaded %s -> %s", url, dest)
    return True
  except Exception as e:
    logger.warning("Download failed for %s: %s", url, e)
    try:
      if os.path.exists(tmp):
        os.remove(tmp)
    except OSError:
      pass
    return False


def _fc_system_family_for_lang(fc_lang: str, preferred: List[str]) -> Optional[str]:
  if not shutil.which("fc-list"):
    return None
  try:
    out = subprocess.run(
        ["fc-list", f":lang={fc_lang}", "family"],
        capture_output=True,
        text=True,
        timeout=10,
    ).stdout
  except Exception:
    return None
  families = set()
  for line in out.splitlines():
    for fam in line.split(","):
      fam = fam.strip()
      if fam:
        families.add(fam)
  if not families:
    return None
  for pref in preferred:
    if pref in families:
      return pref
  return sorted(families)[0]


def resolve_cjk_fallback_font(key: str) -> Optional[Dict[str, Any]]:
  spec = CJK_FONT_SOURCES[key]
  bundled_paths = [os.path.join(BUNDLED_FONTS_DIR, f) for f in spec["bundled"]]
  if bundled_paths and os.path.exists(bundled_paths[0]):
    faces = []
    for p, w in zip(bundled_paths, spec["bundled_weights"]):
      if os.path.exists(p):
        faces.append({"src": f"url('file://{p}') format('opentype')", "weight": str(w)})
    return {"family": spec["family"], "faces": faces, "source": "bundled", "path": bundled_paths[0]}

  cached = os.path.join(FONT_CACHE_DIR, spec["cache_file"])
  if os.path.exists(cached):
    return {
        "family": spec["family"],
        "faces": [{"src": f"url('file://{cached}') format('truetype')", "weight": "100 900"}],
        "source": "cache",
        "path": cached,
    }

  sys_family = _fc_system_family_for_lang(spec["fc_lang"], spec["system_families"])
  if sys_family:
    return {
        "family": spec["family"],
        "faces": [
            {"src": f"local('{sys_family}')", "weight": "400"},
            {"src": f"local('{sys_family} Bold'), local('{sys_family}')", "weight": "700"},
        ],
        "source": "system",
        "path": sys_family,
    }

  if _download(spec["url"], cached):
    return {
        "family": spec["family"],
        "faces": [{"src": f"url('file://{cached}') format('truetype')", "weight": "100 900"}],
        "source": "download",
        "path": cached,
    }
  return None


def detect_han_owner(html: str) -> str:
  m_lang = re.search(r'<html[^>]*\blang=["\']([^"\']+)["\']', html, re.IGNORECASE)
  lang = m_lang.group(1).strip().lower() if m_lang else ""
  if lang.startswith("ja"):
    return "jp"
  if lang.startswith("zh"):
    return "sc"
  if lang.startswith("ko"):
    return "kr"
  body = re.sub(r"<(script|style)[^>]*>.*?</\1>", "", html, flags=re.IGNORECASE | re.DOTALL)
  if re.search(r"[\u3040-\u30FF]", body):
    return "jp"
  if re.search(r"[\uAC00-\uD7A3]", body):
    return "kr"
  if re.search(r"[\u4E00-\u9FFF]", body):
    return "sc"
  return "kr"


def build_cjk_fallback_font_face_css(html: str) -> Tuple[str, Dict[str, Any]]:
  global LAST_FONT_REPORT
  han_owner = detect_han_owner(html)
  css_parts = []
  report: Dict[str, Any] = {"han_owner": han_owner, "fonts": {}, "warnings": []}
  for key in ("kr", "jp", "sc"):
    spec = CJK_FONT_SOURCES[key]
    resolved = resolve_cjk_fallback_font(key)
    if not resolved:
      report["fonts"][key] = {"source": None}
      msg = (
          f"No {spec['fc_lang']}-capable font found (bundled/cache/system/download); "
          f"{spec['family']} text may render as tofu. Install one (e.g. `sudo apt install "
          "fonts-noto-cjk`) or allow network access for a one-time download."
      )
      report["warnings"].append(msg)
      logger.warning(msg)
      continue
    report["fonts"][key] = {"source": resolved["source"], "path": resolved["path"]}
    ranges = spec["ranges"]
    if key == han_owner and key != "sc":
      ranges = f"{ranges}, {_HAN_RANGE}"
    for face in resolved["faces"]:
      css_parts.append(
          "@font-face {"
          f" font-family: '{resolved['family']}';"
          f" font-weight: {face['weight']};"
          f" src: {face['src']};"
          f" unicode-range: {ranges};"
          " }"
      )
  LAST_FONT_REPORT = report
  return "\n".join(css_parts), report


_FONT_FACE_BLOCK = re.compile(r"@font-face\s*\{[^}]*\}", re.IGNORECASE | re.DOTALL)
_FONT_FAMILY_DECL = re.compile(
    r"(font-family\s*:\s*)((?:[^;}\"'<>]|\"[^\"]*\"|'[^']*')+?)(\s*!important)?(?=\s*[;}\"'<])",
    re.IGNORECASE,
)


def append_cjk_fallback_to_font_families(html: str) -> str:
  placeholders: List[str] = []

  def _protect(m):
    placeholders.append(m.group(0))
    return f"\x00FONTFACE{len(placeholders) - 1}\x00"

  protected = _FONT_FACE_BLOCK.sub(_protect, html)

  def _rewrite(m):
    value = m.group(2)
    if "HTS CJK" in value:
      return m.group(0)
    return f"{m.group(1)}{value.rstrip()}{CJK_FALLBACK_SUFFIX}{m.group(3) or ''}"

  rewritten = _FONT_FAMILY_DECL.sub(_rewrite, protected)

  def _restore(m):
    return placeholders[int(m.group(1))]

  return re.sub(r"\x00FONTFACE(\d+)\x00", _restore, rewritten)


JS_APPLY_CJK_FALLBACK = (
    "(() => {"
    f" const SUF = {CJK_FALLBACK_SUFFIX!r};"
    " let n = 0;"
    " document.querySelectorAll('body, body *').forEach(el => {"
    "   try {"
    "     const ff = getComputedStyle(el).fontFamily || '';"
    "     if (ff && !ff.includes('HTS CJK')) { el.style.fontFamily = ff + SUF; n++; }"
    "   } catch (e) {}"
    " });"
    " return n;"
    "})()"
)


def resolve_js_asset(filename: str) -> Optional[str]:
  bundled = os.path.join(BUNDLED_ASSETS_DIR, filename)
  if os.path.exists(bundled):
    return bundled
  cached = os.path.join(ASSET_CACHE_DIR, filename)
  if os.path.exists(cached):
    return cached
  spec = JS_ASSET_SOURCES.get(filename)
  if spec and _download(spec["url"], cached):
    return cached
  return None
