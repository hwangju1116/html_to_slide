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

Converts HTML slide decks (and slide images via the Host Agent's vision) into **100% editable native 16:9 PowerPoint (`.pptx`)** presentations locally. All cards, badges, text frames, tables, and Chart.js canvases are built directly from live DOM geometry as native PowerPoint shapes, tables, and charts—requiring **no MCP server, no Cloud Run deployment, and no GCP project configuration**.

## Directory Structure

```text
html-to-pptx/
├── SKILL.md                                  # Core workflow instructions (this file)
├── scripts/
│   ├── convert_html_to_pptx.py               # CLI: HTML -> Native editable 16:9 .pptx (with --prepare & --builder-script)
│   ├── browser_renderer.py                   # Headless Chrome 1920x1080 CDP live DOM geometry & screenshot engine
│   ├── js_geometry_extractor.py              # Live DOM bounding-box & computed-style extractor
│   ├── pptx_native_builders.py               # Native python-pptx shape/table/chart builders & collision resolver
│   └── color_utils.py                        # CSS variable & RGBA alpha-blending parser
├── references/
│   ├── html_slide_authoring_guide.md         # 16:9 HTML/CSS semantic rules for image-to-slide digitization
│   └── cli_and_architecture.md               # Detailed CLI options, Host AI custom builder hook, and troubleshooting
└── assets/
    ├── fonts/                                # Bundled NotoSansKR fallback fonts
    ├── chart.min.js                          # Offline Chart.js bundle
    └── tailwindcss.min.js                    # Offline Tailwind CSS bundle
```

---

## Decision Tree & Workflows

Determine the user's input type and execute the matching workflow using `<SKILL_DIR>` (the directory containing this `SKILL.md`):

1. **Input is an `.html` file (or raw HTML slide code) → Convert to `.pptx`**:
   - **Host AI Vision + DOM Manifest Conversion (Default — Recommended for 1:1 visual fidelity)**: Follow **Workflow A-1** (`--prepare` → inspect 16:9 screenshots & `manifest.md` → write `builder.py` → `--builder-script`).
   - **Fast Deterministic CLI Conversion (Only when explicitly requested for speed without AI inspection)**: Follow **Workflow A-2**.
2. **Input is a slide image (`.png`, `.jpg`, `.webp`) → Convert to editable `.pptx` / `.html`**: Follow **Workflow B**.

---

### Workflow A-1 (Default): Host AI Vision + Full DOM Manifest Custom Builder (`--prepare` + `--builder-script`)

To achieve 1:1 visual fidelity (graphs, text colors, process flows, sub-cards, and tables), **you (the Host Agent)** must inspect the rendered 16:9 slide screenshots and the extracted live DOM geometry manifest before compiling the final `.pptx`:

1. If the user provided raw HTML in chat rather than a file path, save it to a `.html` file in the workspace first.
2. **Prepare 16:9 Screenshots & Full DOM Style Manifest**:
   ```bash
   python "<SKILL_DIR>/scripts/convert_html_to_pptx.py" "<input.html>" --prepare "<work_dir>" --json
   ```
   *(Note: If `<SKILL_DIR>/.venv/bin/python` exists, prefer using `"<SKILL_DIR>/.venv/bin/python"`.)*
3. **Inspect Screenshots & Manifest (`manifest.md`)**:
   - Call `view_file` on `<work_dir>/manifest.md` (which contains `[AUTHENTIC SLIDE STRUCTURE & STYLE MANIFEST]`, `[STRICT INSTRUCTIONS FOR THE SLIDE BUILDER AI]`, exact `left`/`top`/`width`/`height` inch coordinates, RGB/HEX colors, `pipelineSteps`, `flows`, `fItems`, `subCards`, `charts`, `tables`, and `[EXACT MEASURED BROWSER LAYOUT GEOMETRY]`).
   - Call `view_file` on each rendered screenshot (`<work_dir>/slide-1.png`, `<work_dir>/slide-2.png`, ...) to visually verify colors, process flows, charts, and typography.
4. **Write Custom Builder Script (`<work_dir>/builder.py`)**:
   - Define `def build_slide_1(prs, slide): ...`, `def build_slide_2(prs, slide): ...` (or `def build_slide(prs, slide, slide_index): ...`).
   - Inside each `build_slide_<N>(prs, slide)`, you can either:
     - Call `build_slide_from_geometry(slide, slide_geometry)` and then refine/add any custom visual details observed in `slide-<N>.png`, OR
     - Construct the slide shapes directly using the exact `Inches(...)` coordinates and `RGBColor(...)` values from `manifest.md`, using `build_styled_native_chart(slide, ch, ...)` for `charts_data` and `build_styled_native_table(slide, tbl, ...)` for `table_data`.
   - Available pre-injected scope variables: `prs`, `slide`, `slide_index`, `slide_geometry`, `slide_dom_data`, `design_tokens`, `table_data`, `charts_data`, `build_styled_native_table`, `build_styled_native_chart`, `apply_semantic_styles_to_table`, `build_slide_from_geometry`, `Inches`, `Pt`, `RGBColor`, `MSO_SHAPE`, `PP_ALIGN`, `MSO_ANCHOR`.
5. **Compile Final PPTX**:
   ```bash
   python "<SKILL_DIR>/scripts/convert_html_to_pptx.py" "<input.html>" -o "<output.pptx>" --builder-script "<work_dir>/builder.py" --json
   ```
6. Verify that the JSON output reports `"status": "success"` and report the generated `output_pptx_path` and `slide_count` to the user.

---

### Workflow A-2: Fast Deterministic Direct Conversion (No Custom Builder Script)

When fast deterministic conversion is needed without custom per-slide code synthesis:

1. Run [`scripts/convert_html_to_pptx.py`](scripts/convert_html_to_pptx.py) directly:
   ```bash
   python "<SKILL_DIR>/scripts/convert_html_to_pptx.py" "<input.html>" -o "<output.pptx>" --json
   ```
   - To convert only a specific slide, append `-s <slide_number_or_id>` (e.g., `-s 2` or `-s slide-3`).

---

### Workflow B: Slide Image (`.png` / `.jpg`) → Editable HTML & Native PPTX (Host AI Vision)

Instead of calling external APIs from Python, **you (the Host Agent)** perform the multimodal vision transcription directly:

1. **Inspect the Slide Image**: Call `view_file` on the user's slide image (`.png` / `.jpg`) to analyze its layout, colors, typography, cards, tables, and charts.
2. **Read the Authoring Guide**: Read [`references/html_slide_authoring_guide.md`](references/html_slide_authoring_guide.md) to use the exact semantic classes (`.slide`, `.slide-tag`, `.card`, `.sub-card`, `.f-item`, `.highlight-box`, `<table>`, Chart.js `<canvas>`) recognized by the native PPTX geometry builder.
3. **Generate 16:9 Semantic HTML**: Write a self-contained `1920x1080` `.html` file that reproduces 100% of the Korean/English text, HEX colors, and layout from the image.
4. **Convert to Native PPTX**: Follow **Workflow A-1** on the generated `.html` file to produce the editable `.pptx` deck.
