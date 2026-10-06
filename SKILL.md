---
name: html-to-pptx
description: >-
  Converts single-slide or multi-slide HTML presentations and slide images into
  100% native editable 16:9 widescreen PowerPoint (.pptx) decks using local
  headless Chromium DOM geometry extraction and python-pptx. Use when the user
  asks to convert an .html file or HTML slide deck into PPTX/PowerPoint, or
  digitize slide images (.png/.jpg) into editable PowerPoint shapes, tables,
  charts, and text boxes.
---

# HTML-to-PPTX Native Presentation Converter Skill

Converts HTML slide decks (and slide images via the Host Agent's vision) into **100% editable native 16:9 PowerPoint (`.pptx`)** presentations locally. All cards, badges, text frames, tables, and Chart.js canvases are built directly from live DOM geometry as native PowerPoint shapes, tables, and charts—requiring **no screen captures, no MCP server, no Cloud Run deployment, and no GCP project configuration**.

## Directory Structure

```text
html-to-pptx/
├── SKILL.md                                  # Core workflow instructions (this file)
├── scripts/
│   ├── convert_html_to_pptx.py               # CLI: HTML -> Native editable 16:9 .pptx
│   ├── browser_renderer.py                   # Headless Chrome 1920x1080 CDP live DOM geometry engine
│   ├── js_geometry_extractor.py              # Live DOM bounding-box & computed-style extractor
│   ├── pptx_native_builders.py               # Native python-pptx shape/table/chart builders & collision resolver
│   └── color_utils.py                        # CSS variable & RGBA alpha-blending parser
├── references/
│   ├── html_slide_authoring_guide.md         # 16:9 HTML/CSS semantic rules for image-to-slide digitization
│   └── cli_and_architecture.md               # Detailed CLI options, custom builder hook, and troubleshooting
└── assets/
    ├── fonts/                                # Bundled NotoSansKR fallback fonts
    ├── chart.min.js                          # Offline Chart.js bundle
    └── tailwindcss.min.js                    # Offline Tailwind CSS bundle
```

---

## Decision Tree & Workflows

Determine the user's input type and execute the matching workflow using `<SKILL_DIR>` (the directory containing this `SKILL.md`):

1. **Input is an `.html` file (or raw HTML slide code) → Convert to `.pptx`**: Follow **Workflow A**.
2. **Input is a slide image (`.png`, `.jpg`, `.webp`) → Convert to editable `.pptx` / `.html`**: Follow **Workflow B**.

---

### Workflow A: Direct HTML → Native Editable PPTX Conversion

1. If the user provided raw HTML in chat rather than a file path, save it to a `.html` file in the workspace first.
2. Run [`scripts/convert_html_to_pptx.py`](scripts/convert_html_to_pptx.py):
   ```bash
   python "<SKILL_DIR>/scripts/convert_html_to_pptx.py" "<input.html>" -o "<output.pptx>" --json
   ```
   *(Note: If `<SKILL_DIR>/.venv/bin/python` exists, prefer using `"<SKILL_DIR>/.venv/bin/python"`.)*
   - To convert only a specific slide, append `-s <slide_number_or_id>` (e.g., `-s 2` or `-s slide-3`).
   - For advanced CLI flags (`--builder-script`, `--scale`), consult [`references/cli_and_architecture.md`](references/cli_and_architecture.md).
3. Verify that the JSON output reports `"status": "success"` and report the generated `output_pptx_path` and `slide_count` to the user.

---

### Workflow B: Slide Image (`.png` / `.jpg`) → Editable HTML & Native PPTX (Host AI Vision)

Instead of calling external APIs from Python, **you (the Host Agent)** perform the multimodal vision transcription directly:

1. **Inspect the Slide Image**: Call `view_file` on the user's slide image (`.png` / `.jpg`) to analyze its layout, colors, typography, cards, tables, and charts.
2. **Read the Authoring Guide**: Read [`references/html_slide_authoring_guide.md`](references/html_slide_authoring_guide.md) to use the exact semantic classes (`.slide`, `.slide-tag`, `.card`, `.sub-card`, `.f-item`, `.highlight-box`, `<table>`, Chart.js `<canvas>`) recognized by the native PPTX geometry builder.
3. **Generate 16:9 Semantic HTML**: Write a self-contained `1920x1080` `.html` file that reproduces 100% of the Korean/English text, HEX colors, and layout from the image.
4. **Convert to Native PPTX**: Run [`scripts/convert_html_to_pptx.py`](scripts/convert_html_to_pptx.py) on the generated `.html` file to produce the editable `.pptx` deck.
