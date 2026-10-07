# CLI Reference, Custom Builder Hook & Troubleshooting

This document provides detailed CLI flags, JSON output contracts, custom `python-pptx` code hooks, and architecture guarantees for the `html-to-pptx` skill scripts.

---

## 1. `scripts/convert_html_to_pptx.py`

Converts single-slide or multi-slide HTML files into 100% native editable 16:9 PowerPoint (`.pptx`) presentations using headless Chromium DOM geometry extraction and `python-pptx`.

### Usage

```bash
python scripts/convert_html_to_pptx.py <input.html> [options]
```

### Options
- `-o, --output <path.pptx>`: Output path for the generated `.pptx` file. Defaults to `<input_stem>.pptx` in the same directory as the input HTML.
- `-s, --slide <id_or_index>`: Convert only a specific slide (e.g., `1`, `2`, `slide-3`). Omit to convert all slides in the deck.
- `--scale <int>`: Device scale factor for headless Chromium rendering (default: `2` for Retina DPI).
- `--prepare <work_dir>`: Render `16:9` PNG screenshots (`slide-1.png`, ...) and generate `manifest.md` (`[AUTHENTIC SLIDE STRUCTURE & STYLE MANIFEST]`) + `geometry.json` in `<work_dir>` for Host AI Vision code synthesis.
- `--builder-script <script.py>`: Optional path to a custom Python script defining `def build_slide_<N>(prs, slide): ...` or `def build_slide(prs, slide): ...` written by the Host AI. When provided, the converter executes the custom builder, applies table/chart/SVG safety nets, and runs typography harmonization and collision resolution.
- `--json`: Print structured JSON result to `stdout`.

### JSON Output Schema (`--json`)

```json
{
  "status": "success",
  "success": true,
  "slide_count": 3,
  "slide_ids": ["slide-1", "slide-2", "slide-3"],
  "output_pptx_path": "/abs/path/to/deck.pptx",
  "file_size_bytes": 54320
}
```

---

## 2. Host AI Vision + Custom `python-pptx` Builder Script Hook (`--prepare` + `--builder-script`)

To synthesize custom `python-pptx` code guided by `16:9` rendered screenshots and the DOM style manifest:

1. **Generate Screenshots & Manifest**:
   ```bash
   python scripts/convert_html_to_pptx.py input.html --prepare ./prep_dir --json
   ```
2. **Write `custom_slide.py`**:
   You can define per-slide functions `build_slide_1(prs, slide)`, `build_slide_2(prs, slide)`, etc. (any slide without a `build_slide_<N>` function automatically falls back to `build_slide_from_geometry(slide, slide_geometry)`):
   ```python
   def build_slide_1(prs, slide):
       card = slide.shapes.add_shape(
           MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.8), Inches(5.5), Inches(4.2)
       )
       card.fill.solid()
       card.fill.fore_color.rgb = RGBColor(248, 250, 252)
       tf = card.text_frame
       tf.word_wrap = True
       p = tf.paragraphs[0]
       p.text = "Custom Native Card"
       p.font.name = "Pretendard"
       p.font.size = Pt(16)
       p.font.bold = True
   ```
3. **Build the Final PPTX**:
   ```bash
   python scripts/convert_html_to_pptx.py input.html -o output.pptx --builder-script custom_slide.py --json
   ```

---

## 3. Architecture & Post-Build Quality Guarantees

- **Live DOM Geometry & Computed Style Extraction**: Slides are converted from live browser DOM bounding boxes (`getBoundingClientRect()`) and computed CSS styles (`getComputedStyle()`).
- **PowerPoint Table Style Override Prevention**: Automatically disables PowerPoint's default `firstRow="1"` / `bandRow="1"` theme overrides (`_disable_default_pptx_table_style`), applies run-level `<a:rPr>` font colors and sizes, draws parent `.info-card` wrappers and `<h3>` titles, and renders crisp horizontal cell borders.
- **Zero External AI / Zero GCP Dependency**: All scripts under `scripts/` execute 100% locally using headless Chrome and `python-pptx`. No `GOOGLE_CLOUD_PROJECT`, `gcloud auth`, or external LLM API keys are needed.
- **Offline Bundled Assets (`assets/`)**: External CDN references to Tailwind CSS and Chart.js are automatically rewritten to local `assets/tailwindcss.min.js` and `assets/chart.min.js` so headless Chrome renders instantaneously without network timeouts.
- **Font Resolution (`assets/fonts/`)**: Uses local `Pretendard` (`~/.local/share/fonts/pretendard`) and bundled `assets/fonts/NotoSansKR-*.otf` fallback fonts for 1:1 Korean/English typography metrics.
- **Collision Resolution (`audit_and_resolve_slide_collisions`)**: Automatically deduplicates full-bleed background shapes, removes duplicate overlapping text frames, and separates horizontal badge-vs-text overlaps.
