---
name: html-to-pptx
description: Converts single-slide or multi-slide HTML presentations into 100% native editable 16:9 widescreen PowerPoint (.pptx) decks via the html-to-pptx MCP server. Use when the user asks to convert an .html file or HTML slide deck into PPTX/PowerPoint, capture 16:9 widescreen slide previews from HTML, or digitize slide images into editable PowerPoint shapes and text boxes.
allowed-tools: convert_html_to_pptx_mcp capture_html_slides_mcp recreate_editable_slide_mcp
metadata:
  adk_additional_tools:
    - convert_html_to_pptx_mcp
    - capture_html_slides_mcp
    - recreate_editable_slide_mcp
---

# HTML-to-PPTX Native Presentation Converter Skill (MCP-Connected)

This skill connects directly to the **`html-to-pptx` MCP server** (`app/mcp_server.py`) to convert HTML presentations into 100% editable native PowerPoint (`.pptx`) slides without needing a separate sub-agent.

## Available MCP Tools

1. **`convert_html_to_pptx_mcp`**
   - **Purpose**: End-to-end conversion of a single-slide or multi-slide HTML file (or raw HTML string) into a 16:9 widescreen PowerPoint (`.pptx`) deck with 100% editable native shapes, cards, badges, tables, and text boxes.
   - **Parameters**:
     - `html_input_path` *(str, required)*: Absolute or workspace-relative path to the `.html` file, or raw HTML markup.
     - `output_pptx_path` *(str, optional)*: Output path for the generated `.pptx` file. Defaults to `<input_stem>.pptx` in the same directory.
     - `target_slide_id` *(str, optional)*: Optional 1-based slide index or HTML id (e.g., `"1"`, `"slide-3"`) to convert only a single slide. Omit to convert all slides.
     - `native_mode` *(bool, default `True`)*: Keep `True` to generate 100% editable native PowerPoint shapes and text boxes.

2. **`capture_html_slides_mcp`**
   - **Purpose**: Renders an HTML slide deck inside a locked `1920x1080` (16:9 widescreen) headless Chromium viewport and saves high-DPI PNG previews of each slide.
   - **Parameters**:
     - `html_input_path` *(str, required)*: Path to the `.html` file or raw HTML markup.
     - `output_dir` *(str, optional)*: Directory where slide PNGs will be written (defaults to `<workspace>/captures`).
     - `target_slide_id` *(str, optional)*: Optional 1-based slide index or HTML id.

3. **`recreate_editable_slide_mcp`**
   - **Purpose**: Uses Gemini 3.8 Flash multimodal vision decomposition to convert a raster slide screenshot (`.png` / `.jpg`) into an editable 16:9 `.pptx` slide.
   - **Parameters**:
     - `image_path` *(str, required)*: Path to the slide image file.
     - `output_pptx_path` *(str, optional)*: Destination path for the generated `.pptx` file.

---

## Execution Workflows

### Workflow A: Direct HTML-to-PPTX Conversion (Default — Local & Cloud Run Compatible)
When the user provides an `.html` file (or raw HTML slide code) and asks to convert it to PowerPoint / `.pptx`:
1. Invoke `convert_html_to_pptx_mcp` with:
   - `html_input_path`: Path to the user's HTML file (or raw HTML markup string if the MCP server is hosted remotely on Cloud Run).
   - `output_pptx_path`: Optional local path if requested by the user.
   - `native_mode`: `True` (ensures 100% editable shapes, text boxes, and tables).
2. **Remote Cloud Run Fallback**: If `convert_html_to_pptx_mcp` returns `"status": "error"` stating that the file path does not exist on the server filesystem (because the MCP server is running on Cloud Run):
   - Read the local `.html` file contents using `view_file` (or python/cat).
   - Call `convert_html_to_pptx_mcp(html_input_path=<raw_html_string>, native_mode=True)`.
3. **Save PPTX Locally (When Connected to Cloud Run)**:
   - Parse the JSON response (`output_pptx_path`, `download_url`, `slide_count`, `preview_image_paths`).
   - If `download_url` starts with `http://` or `https://` and `output_pptx_path` is a remote `/tmp/...` path, download the generated `.pptx` file into the user's local workspace directory using `curl -sSL "<download_url>" -o "<local_target.pptx>"`.
4. Report the final local `.pptx` file path (and `download_url` if remote) and slide count clearly to the user.

### Workflow B: Visual Preview First, Then Convert
When the user asks to preview or inspect how an HTML slide renders in 16:9 before creating the PPTX:
1. Invoke `capture_html_slides_mcp(html_input_path=...)` (pass raw HTML markup if running against a remote Cloud Run server).
2. Show the captured 16:9 PNG file paths (`image_paths` or `download_urls`) to the user for confirmation.
3. Once confirmed, call `convert_html_to_pptx_mcp(html_input_path=..., native_mode=True)` (for exact DOM-to-PPTX vector conversion) or `recreate_editable_slide_mcp(image_path=...)` (for vision-based slide digitization).

---

## Architecture & Quality Guarantees
- **Locked 16:9 Viewport (`1920x1080`)**: All HTML slides are measured inside a fixed `1920x1080` Chromium viewport with viewer chrome hidden (`.chrome { display: none }`) and `.frame` / `.stage` / `.slide` expanded to `1920x1080`, mapping 1:1 to `13.333 in × 7.5 in` (`960pt × 540pt`) widescreen PowerPoint coordinates.
- **Live DOM Geometry Extraction**: Every card, badge, Flex/Grid row, table cell, and SVG/Canvas diagram is extracted via live `getBoundingClientRect()` and `getComputedStyle()` rather than static regex parsing.
- **Zero-Overlap Post-Build Audit**: `audit_and_resolve_slide_collisions` automatically removes duplicate full-bleed background rectangles, deduplicates overlapping identical text boxes, separates horizontal badge collisions, and resolves vertical text box overlaps while preserving exact source font sizes.

