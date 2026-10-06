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
- `--builder-script <script.py>`: Optional path to a custom Python script defining `def build_slide(slide, prs): ...` written by the Host AI. When provided, the converter executes `build_slide`, applies table/chart/SVG safety nets, and runs typography harmonization and collision resolution.
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

## 2. Custom `python-pptx` Builder Script Hook (`--builder-script`)

If the Host AI wants to construct custom vector shapes directly with `python-pptx` while still leveraging the skill's automatic canvas background, table/chart builders, accent-bar refinement, and collision resolver:

```python
def build_slide(prs, slide):
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

Run with:
```bash
python scripts/convert_html_to_pptx.py input.html -o output.pptx --builder-script custom_slide.py --json
```

---

## 3. Architecture & Post-Build Quality Guarantees

- **Zero Disk Screen Capture**: Slides are converted directly from live browser DOM bounding boxes (`getBoundingClientRect()`) and computed CSS styles (`getComputedStyle()`). In-memory byte cropping is invoked only when a slide contains `<svg>` or `<img>` elements that need to be embedded as cropped visual assets.
- **Zero External AI / Zero GCP Dependency**: All scripts under `scripts/` execute 100% locally using headless Chrome and `python-pptx`. No `GOOGLE_CLOUD_PROJECT`, `gcloud auth`, or external LLM API keys are needed.
- **Offline Bundled Assets (`assets/`)**: External CDN references to Tailwind CSS and Chart.js are automatically rewritten to local `assets/tailwindcss.min.js` and `assets/chart.min.js` so headless Chrome renders instantaneously without network timeouts.
- **Font Resolution (`assets/fonts/`)**: Uses local `Pretendard` (`~/.local/share/fonts/pretendard`) and bundled `assets/fonts/NotoSansKR-*.otf` fallback fonts for 1:1 Korean/English typography metrics.
- **Collision Resolution (`audit_and_resolve_slide_collisions`)**: Automatically deduplicates full-bleed background shapes, removes duplicate overlapping text frames, and separates horizontal badge-vs-text overlaps.
