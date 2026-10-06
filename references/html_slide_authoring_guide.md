# 16:9 HTML Slide Authoring Guide for Native PPTX Conversion

When the Host Agent needs to digitize a raster slide image (`.png`/`.jpg`) into an editable `.pptx` or refine an HTML slide deck before running [`scripts/convert_html_to_pptx.py`](../scripts/convert_html_to_pptx.py), structure the HTML using the semantic classes below so the DOM geometry extractor maps every element to a 100% native PowerPoint shape, table, or chart.

---

## 1. 16:9 Viewport & Slide Container Structure

All slides are rendered inside a locked `1920px × 1080px` viewport (`13.333 in × 7.5 in` in PowerPoint).

```html
<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8">
  <style>
    :root {
      --brand-color: #2563EB;
      --bg-color: #FFFFFF;
      --text-main: #111115;
      --text-muted: #64748B;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      width: 1920px;
      height: 1080px;
      font-family: 'Pretendard', 'Noto Sans KR', sans-serif;
      background: var(--bg-color);
      color: var(--text-main);
    }
    .slide {
      width: 1920px;
      height: 1080px;
      padding: 64px 80px;
      display: flex;
      flex-direction: column;
      position: relative;
      background: var(--bg-color);
    }
  </style>
</head>
<body>
  <section id="slide-1" class="slide">
    <!-- Slide content -->
  </section>
</body>
</html>
```

---

## 2. Semantic Element Mapping (`HTML DOM` → `Native PPTX`)

| Visual Component | Recommended HTML Selector / Class | Resulting Native PowerPoint Object |
| :--- | :--- | :--- |
| **Slide Container** | `<section id="slide-1" class="slide">` or `<div class="slide">` | 16:9 Slide (`13.333" × 7.5"`) with full-bleed background fill |
| **Header Eyebrow / Pill** | `.slide-tag`, `.eyebrow`, `.badge-pill` | Rounded rectangle badge or left-aligned eyebrow text frame |
| **Slide Number** | `.slide-number`, `.slide-counter` | Top-right or footer slide number text box |
| **Master Title** | `<h1>` or `<h2>` (top-level inside `.slide`) | Master title text frame with exact measured font size & color |
| **Master Subtitle** | `p.subtitle`, `p.lead`, `p.premise` | Subtitle text frame below the master title |
| **Card / Panel** | `.card`, `.info-card`, `.panel`, `.box`, `.pillar` | Rounded rectangle container (`MSO_SHAPE.ROUNDED_RECTANGLE`) with fill, border, and optional top/left accent bar |
| **Card Accent Bar** | CSS `border-top: 4px solid ...` or `border-left: 4px solid ...` on `.card` | Inset rounded capsule accent bar at top or left of the card |
| **Nested Sub-Card** | `.sub-card` inside a `.card` | Secondary rounded rectangle with heading and bullet list |
| **Feature Item (Icon + Text)** | `.f-item` containing `.icon` + text container | Separate rounded icon badge + aligned text box |
| **Data Table** | `<table>` with `<thead><tr><th>` and `<tbody><tr><td>` | Native PowerPoint table (`slide.shapes.add_table`) with cell fills, borders, and bold runs |
| **Chart.js Canvas** | `<canvas id="chart-1"></canvas>` initialized via `new Chart(ctx, config)` | Native editable PowerPoint chart (`LINE_MARKERS`, `BAR_CLUSTERED`, `COLUMN_CLUSTERED`, `DOUGHNUT`, `PIE`) |
| **Callout / Highlight Box** | `.highlight-box`, `.callout`, `.alert` | Bottom or inline callout banner with border and bold text |
| **Footnote** | `.footnote`, `footer p` | Footer text box with divider line if `border-top` is present |
| **Vector Diagram / Icon** | `<svg>` or `<img>` | High-DPI cropped PNG embedded at exact bounding-box coordinates |

---

## 3. Digitizing a Slide Image (`.png` / `.jpg`) via Host AI

When the user asks to convert a slide screenshot or image into an editable PowerPoint file:

1. **Inspect the Image**: Use `view_file` on the image path to examine its layout, colors, typography, cards, tables, and charts.
2. **Author 16:9 Semantic HTML**: Write a self-contained `.html` file using the classes in Section 2:
   - Transcribe 100% of all Korean and English text, numbers, badges, and footnotes verbatim.
   - Extract exact HEX colors for background, cards, borders, and text.
   - Use CSS Grid or Flexbox inside `.slide` (`1920px × 1080px`) to match the visual proportions.
3. **Convert to Native PPTX**: Run `python scripts/convert_html_to_pptx.py <generated.html> -o <output.pptx> --json`.
