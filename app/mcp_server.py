import argparse
import base64
import json
import os
from pathlib import Path
import shutil
import sys
import uuid
from typing import Optional

PROJECT_ROOT = str(Path(__file__).resolve().parent.parent)
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from dotenv import load_dotenv
from mcp.server.fastmcp import FastMCP
from mcp.server.transport_security import TransportSecuritySettings
from starlette.requests import Request
from starlette.responses import FileResponse, JSONResponse

load_dotenv(os.path.join(PROJECT_ROOT, ".env"))

from app.genai_client import (
    configure_vertex_environment,
    resolve_gcp_project,
    verify_gcp_auth,
)

configure_vertex_environment()
os.environ["GOOGLE_CLOUD_PROJECT"] = resolve_gcp_project()
os.environ["GOOGLE_CLOUD_LOCATION"] = os.environ.get("VERTEX_AI_LOCATION", "global")
os.environ.setdefault("GOOGLE_GENAI_USE_VERTEXAI", "True")

from app.html_to_pptx_converter import (
    capture_html_slides,
    convert_html_to_pptx,
)
from app.slide_digitizer import digitize_slide_image

DOWNLOAD_DIR = os.environ.get(
    "HTML_TO_PPTX_DOWNLOAD_DIR", "/tmp/html_to_pptx_downloads"
)
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

DEFAULT_PORT = int(os.environ.get("PORT", "8080"))
DEFAULT_HOST = os.environ.get("HOST", "0.0.0.0")

mcp = FastMCP(
    "html-to-pptx",
    host=DEFAULT_HOST,
    port=DEFAULT_PORT,
    transport_security=TransportSecuritySettings(
        enable_dns_rebinding_protection=False
    ),
    stateless_http=True,
)


@mcp.custom_route("/health", methods=["GET"])
async def health_check(request: Request) -> JSONResponse:
    """Cloud Run health check endpoint."""
    return JSONResponse(
        {
            "status": "ok",
            "service": "html-to-pptx-mcp",
            "gcp_project": resolve_gcp_project(),
            "location": os.environ.get("GOOGLE_CLOUD_LOCATION", "global"),
        }
    )


@mcp.custom_route("/downloads/{filename}", methods=["GET"])
async def download_artifact(request: Request) -> FileResponse | JSONResponse:
    """Serves generated .pptx or .png files when running as a remote Cloud Run MCP server."""
    filename = os.path.basename(request.path_params.get("filename", ""))
    file_path = os.path.join(DOWNLOAD_DIR, filename)
    if not filename or not os.path.isfile(file_path):
        return JSONResponse(
            {"error": f"Artifact '{filename}' not found."}, status_code=404
        )
    media_type = (
        "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        if filename.endswith(".pptx")
        else "image/png"
    )
    return FileResponse(
        file_path,
        media_type=media_type,
        filename=filename,
    )


def _is_raw_html(text: str) -> bool:
    """Checks whether the input string is raw HTML markup rather than a file path."""
    stripped = text.lstrip()
    return (
        stripped.startswith("<")
        or "<!doctype" in stripped[:200].lower()
        or "<html" in stripped[:500].lower()
        or 'class="slide' in stripped
    )


def _resolve_path(input_path: str) -> str:
    """Resolves relative file paths against current working directory or project root."""
    if _is_raw_html(input_path):
        return input_path
    if os.path.isabs(input_path) or os.path.exists(input_path):
        return os.path.abspath(input_path) if os.path.exists(input_path) else input_path
    candidate = os.path.join(PROJECT_ROOT, input_path)
    if os.path.exists(candidate):
        return os.path.abspath(candidate)
    return input_path


def _stage_for_download(local_file_path: str) -> tuple[str, Optional[str]]:
    """Copies a generated artifact into DOWNLOAD_DIR and returns (staged_filename, public_download_url)."""
    if not local_file_path or not os.path.isfile(local_file_path):
        return "", None
    ext = Path(local_file_path).suffix or ".pptx"
    stem = Path(local_file_path).stem or "presentation"
    short_id = uuid.uuid4().hex[:8]
    staged_name = f"{stem}_{short_id}{ext}"
    staged_path = os.path.join(DOWNLOAD_DIR, staged_name)
    shutil.copy2(local_file_path, staged_path)

    base_url = os.environ.get("PUBLIC_BASE_URL", "").rstrip("/")
    if base_url:
        return staged_name, f"{base_url}/downloads/{staged_name}"
    return staged_name, f"/downloads/{staged_name}"


def _reload_converter_modules():
    """Returns converter module, optionally hot-reloading submodules if MCP_DEV_RELOAD=1."""
    import app.html_to_pptx_converter as _conv_mod

    if os.environ.get("MCP_DEV_RELOAD") == "1":
        import importlib
        import app.schemas as _m1
        import app.color_utils as _m2
        import app.js_geometry_extractor as _m3
        import app.browser_renderer as _m4
        import app.pptx_native_builders as _m5
        import app.vision_fallback_builder as _m6

        for _mod in (_m1, _m2, _m3, _m4, _m5, _m6, _conv_mod):
            importlib.reload(_mod)
    return _conv_mod


@mcp.tool()
def convert_html_to_pptx_mcp(
    html_input_path: str,
    output_pptx_path: Optional[str] = None,
    target_slide_id: Optional[str] = None,
    native_mode: bool = True,
) -> str:
    """Converts a single-slide or multi-slide HTML file or raw HTML string into a 100% native editable 16:9 PowerPoint (.pptx) presentation.

    Uses live headless Chromium DOM geometry extraction (`getBoundingClientRect`
    + `getComputedStyle`), SVG/Canvas rasterization, and deterministic post-build
    collision auditing so that cards, badges, tables, and typography are built as
    editable native PowerPoint shapes and text boxes without overlap.

    NOTE FOR REMOTE (CLOUD RUN) SERVERS:
    If connected to a remote Cloud Run MCP server, the server cannot read local
    client paths directly. Pass the raw HTML markup string in `html_input_path`,
    then download the returned `download_url` to the user's local `output_pptx_path`.

    Args:
        html_input_path: Absolute/relative path to the HTML file (local mode), OR raw HTML markup string (`<!DOCTYPE html>...`).
        output_pptx_path: Optional destination path for the generated .pptx file. Defaults to <input_stem>.pptx alongside the source HTML.
        target_slide_id: Optional 1-based slide index or HTML id (e.g. '1', 'slide-3') to convert only a single slide. Omit to convert all slides.
        native_mode: If True (default), constructs 100% editable native PPTX shapes and text boxes. If False, embeds high-DPI slide screenshots.

    Returns:
        JSON string containing status, slide_count, output_pptx_path, download_url, and preview_image_paths.
    """
    auth_ok, _, auth_err = verify_gcp_auth()
    if not auth_ok:
        return json.dumps(
            {"status": "error", "error": auth_err},
            ensure_ascii=False,
            indent=2,
        )

    resolved_input = _resolve_path(html_input_path)

    if not _is_raw_html(resolved_input) and not os.path.exists(resolved_input):
        return json.dumps(
            {
                "status": "error",
                "error": (
                    f"File path '{html_input_path}' does not exist on the MCP server filesystem. "
                    "Because this MCP server may be running remotely (e.g. on Cloud Run), "
                    "please read the local HTML file content first and pass the raw HTML markup string "
                    "directly in `html_input_path`."
                ),
            },
            ensure_ascii=False,
            indent=2,
        )

    if output_pptx_path and not _is_raw_html(resolved_input):
        resolved_output = _resolve_path(output_pptx_path)
    elif output_pptx_path:
        # Check if parent directory of output_pptx_path is writable locally; otherwise stage in DOWNLOAD_DIR
        parent_dir = os.path.dirname(os.path.abspath(output_pptx_path))
        try:
            os.makedirs(parent_dir, exist_ok=True)
            resolved_output = os.path.abspath(output_pptx_path)
        except Exception:
            stem = Path(output_pptx_path).stem or "slide"
            resolved_output = os.path.join(DOWNLOAD_DIR, f"{stem}.pptx")
    else:
        resolved_output = (
            os.path.join(DOWNLOAD_DIR, f"presentation_{uuid.uuid4().hex[:8]}.pptx")
            if _is_raw_html(resolved_input)
            else None
        )

    _conv_mod = _reload_converter_modules()

    result = _conv_mod.convert_html_to_pptx(
        html_input=resolved_input,
        output_pptx_path=resolved_output,
        native_mode=native_mode,
        target_slide_id=target_slide_id,
    )

    final_pptx_path = result.get("output_pptx_path", "")
    staged_name, download_url = _stage_for_download(final_pptx_path)

    return json.dumps(
        {
            "status": "success",
            "slide_count": result.get("slide_count", 0),
            "output_pptx_path": final_pptx_path,
            "download_filename": staged_name,
            "download_url": download_url,
            "native_mode": result.get("native_mode", native_mode),
            "preview_image_paths": result.get("image_paths", []),
        },
        ensure_ascii=False,
        indent=2,
    )


@mcp.tool()
def capture_html_slides_mcp(
    html_input_path: str,
    output_dir: Optional[str] = None,
    target_slide_id: Optional[str] = None,
) -> str:
    """Renders an HTML presentation inside a headless 1920x1080 (16:9 widescreen) Chromium viewport and captures high-DPI PNG previews of each slide.

    Args:
        html_input_path: Absolute/relative path to the HTML presentation file, or raw HTML markup string.
        output_dir: Optional directory path where captured slide PNGs should be saved. Defaults to <project_root>/captures.
        target_slide_id: Optional 1-based slide number or HTML id to capture only a specific slide.

    Returns:
        JSON string containing status, slide_count, image_paths, and download_urls of the captured 16:9 PNG slides.
    """
    auth_ok, _, auth_err = verify_gcp_auth()
    if not auth_ok:
        return json.dumps(
            {"status": "error", "error": auth_err},
            ensure_ascii=False,
            indent=2,
        )

    resolved_input = _resolve_path(html_input_path)
    if not _is_raw_html(resolved_input) and not os.path.exists(resolved_input):
        return json.dumps(
            {
                "status": "error",
                "error": (
                    f"File path '{html_input_path}' does not exist on the MCP server filesystem. "
                    "Please read the local HTML file content first and pass the raw HTML markup string "
                    "directly in `html_input_path`."
                ),
            },
            ensure_ascii=False,
            indent=2,
        )

    target_dir = (
        _resolve_path(output_dir)
        if output_dir
        else os.path.join(DOWNLOAD_DIR, "captures")
    )
    try:
        os.makedirs(target_dir, exist_ok=True)
    except Exception:
        target_dir = os.path.join(DOWNLOAD_DIR, "captures")
        os.makedirs(target_dir, exist_ok=True)

    _conv_mod = _reload_converter_modules()

    result = _conv_mod.capture_html_slides(
        html_input=resolved_input,
        output_dir=target_dir,
        target_slide_id=target_slide_id,
    )

    download_urls = []
    for img_p in result.get("image_paths", []):
        _, d_url = _stage_for_download(img_p)
        if d_url:
            download_urls.append(d_url)

    return json.dumps(
        {
            "status": "success",
            "slide_count": result.get("slide_count", 0),
            "image_paths": result.get("image_paths", []),
            "download_urls": download_urls,
        },
        ensure_ascii=False,
        indent=2,
    )


@mcp.tool()
def recreate_editable_slide_mcp(
    image_path: str,
    output_pptx_path: Optional[str] = None,
) -> str:
    """Digitizes a raster slide screenshot (PNG/JPG or base64 data URI) into a 100% editable native 16:9 PowerPoint (.pptx) file using Gemini 3.8 Flash multimodal vision decomposition.

    Args:
        image_path: Absolute/relative path to the slide screenshot image (PNG/JPG), OR a base64 data URI ('data:image/png;base64,...').
        output_pptx_path: Optional destination path for the generated .pptx file.

    Returns:
        JSON string containing status, output_pptx_path, download_url, preview_html_path, and fidelity_score.
    """
    auth_ok, _, auth_err = verify_gcp_auth()
    if not auth_ok:
        return json.dumps(
            {"status": "error", "error": auth_err},
            ensure_ascii=False,
            indent=2,
        )

    if image_path.startswith("data:image/"):
        header, b64_data = image_path.split(",", 1)
        ext = ".jpg" if "jpeg" in header or "jpg" in header else ".png"
        resolved_image = os.path.join(
            DOWNLOAD_DIR, f"input_slide_{uuid.uuid4().hex[:8]}{ext}"
        )
        with open(resolved_image, "wb") as f:
            f.write(base64.b64decode(b64_data))
    else:
        resolved_image = _resolve_path(image_path)
        if not os.path.exists(resolved_image):
            return json.dumps(
                {
                    "status": "error",
                    "error": (
                        f"Image path '{image_path}' not found on the MCP server. "
                        "If connected to a remote Cloud Run server, pass a 'data:image/png;base64,...' URI."
                    ),
                },
                ensure_ascii=False,
                indent=2,
            )

    out_dir = os.path.dirname(resolved_image) or DOWNLOAD_DIR
    stem = Path(resolved_image).stem

    result = digitize_slide_image(
        image_input=resolved_image,
        output_basename=f"{stem}_editable",
        workspace_dir=out_dir,
    )

    final_pptx = result.get("pptx_path", "")
    if output_pptx_path and final_pptx and os.path.exists(final_pptx):
        try:
            resolved_out = _resolve_path(output_pptx_path)
            os.makedirs(os.path.dirname(os.path.abspath(resolved_out)), exist_ok=True)
            shutil.copy2(final_pptx, resolved_out)
            final_pptx = resolved_out
        except Exception:
            pass

    staged_name, download_url = _stage_for_download(final_pptx)

    return json.dumps(
        {
            "status": "success",
            "output_pptx_path": final_pptx,
            "download_filename": staged_name,
            "download_url": download_url,
            "preview_html_path": result.get("html_path", ""),
            "fidelity_score": result.get("fidelity_score", 0),
        },
        ensure_ascii=False,
        indent=2,
    )


def main():
    """CLI entrypoint for running the MCP server over stdio (local) or SSE / HTTP (Cloud Run)."""
    parser = argparse.ArgumentParser(description="HTML-to-PPTX MCP Server")
    default_transport = os.environ.get(
        "MCP_TRANSPORT",
        "streamable-http" if os.environ.get("K_SERVICE") else "stdio",
    )
    parser.add_argument(
        "--transport",
        choices=["stdio", "sse", "streamable-http"],
        default=default_transport,
        help="MCP transport protocol (default: stdio locally, streamable-http on Cloud Run)",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=DEFAULT_PORT,
        help="Port to bind when running SSE or streamable-http transport",
    )
    args, _ = parser.parse_known_args()

    mcp.settings.port = args.port
    mcp.settings.host = DEFAULT_HOST
    mcp.run(transport=args.transport)


if __name__ == "__main__":
    main()

