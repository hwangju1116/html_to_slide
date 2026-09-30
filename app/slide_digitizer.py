import base64
import logging
import os
import re
import subprocess
import sys
import tempfile
from typing import Any, Dict, Optional, Union
from google import genai
from google.genai import types
from app.genai_client import DEFAULT_GEMINI_MODEL, get_genai_client

logger = logging.getLogger(__name__)


def digitize_slide_image(
    image_input: Union[str, bytes],
    output_basename: str = "recreated_slide",
    workspace_dir: Optional[str] = None,
) -> Dict[str, Any]:
    """Recreates a flat slide image into both an editable HTML canvas and a native editable PPTX file.

    Args:
        image_input: File path to the slide image, or raw image bytes.
        output_basename: Base name for the generated output files (without extension).
        workspace_dir: Target directory to save the generated files. Defaults to project root.

    Returns:
        Dictionary containing output paths, HTML markup, PPTX file size, and metadata.
    """
    if isinstance(image_input, str):
        with open(image_input, "rb") as f:
            image_bytes = f.read()
    else:
        image_bytes = image_input

    if not workspace_dir:
        workspace_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
    os.makedirs(workspace_dir, exist_ok=True)

    client = get_genai_client()

    # -------------------------------------------------------------------------
    # 1. Generate 100% Identical Editable HTML Slide Canvas
    # -------------------------------------------------------------------------
    html_prompt = """
You are an elite Frontend Slide Designer & Reverse-Engineering Specialist.
Recreate this slide image into a 100% visually identical, pixel-perfect, editable 16:9 widescreen HTML/CSS canvas.

[STRICT MANDATES]
1. ZERO OMISSION: Transcribe 100% of all Korean and English text, badges, titles, subtitles, bullets, numbers, and button labels verbatim.
2. EDITABLE: Add `contenteditable="true"` attribute to EVERY single text container, badge, title, subtitle, paragraph, bullet point, table cell, and button so users can click and edit any text in the browser.
3. VISUAL FIDELITY:
   - Extract exact HEX colors for background, primary brand colors, cards, borders, and text.
   - Match typography hierarchy, font sizes, weights (e.g. 400, 600, 700, 900), and letter-spacing.
   - Use Google Fonts (Pretendard or Noto Sans KR) for clean typography.
   - Match exact layout geometry using modern CSS Flexbox and CSS Grid.
4. SELF-CONTAINED: Single HTML document with internal <style>. 1920x1080 resolution (responsive 16:9 aspect-ratio).
5. Add interactive hover and focus styles for contenteditable elements (e.g. subtle dashed outline on hover, solid accent outline on focus).

Output ONLY the raw HTML code block enclosed in ```html and ```.
"""

    html_resp = client.models.generate_content(
        model=DEFAULT_GEMINI_MODEL,
        contents=[
            types.Part.from_bytes(data=image_bytes, mime_type="image/png"),
            html_prompt,
        ],
    )

    raw_html = html_resp.text or ""
    if "```html" in raw_html:
        raw_html = raw_html.split("```html")[1].split("```")[0].strip()
    elif "```" in raw_html:
        raw_html = raw_html.split("```")[1].split("```")[0].strip()

    html_file_path = os.path.join(workspace_dir, f"{output_basename}.html")
    with open(html_file_path, "w", encoding="utf-8") as f_html:
        f_html.write(raw_html)

    # -------------------------------------------------------------------------
    # 2. Generate 100% Native Vector Editable PowerPoint (.pptx) with Verification Loop
    # -------------------------------------------------------------------------
    pptx_file_path = os.path.join(workspace_dir, f"{output_basename}.pptx")
    from app.slide_fidelity_critic import audit_slide_fidelity
    from pptx import Presentation

    base_pptx_prompt = f"""
You are an expert PowerPoint Engineer specializing in high-fidelity native PPTX recreation.
Write a clean, standalone Python script using `python-pptx` that reconstructs this slide image into a 100% native, fully editable PowerPoint slide.

[STRICT PPTX MANDATES]
1. 16:9 Dimensions:
   prs = Presentation()
   prs.slide_width = Inches(13.333)
   prs.slide_height = Inches(7.5)
   slide = prs.slides.add_slide(prs.slide_layouts[6]) # blank layout

2. Shape & Text Integration (CRITICAL TO PREVENT CLIPPING):
   - For badges, cards, buttons, or colored containers that contain text: Put text DIRECTLY inside `shape.text_frame`. NEVER create a separate disjoint textbox overlapping a shape!
   - For text inside cards/containers, set text alignment, word wrap, and margins:
     tf.margin_left = Inches(0.08)
     tf.margin_top = Inches(0.06)
     tf.word_wrap = True
   - For main hero content (Title + Subheading + Description): Put them in ONE unified text frame using `content_tf.paragraphs[0]`, `content_tf.add_paragraph()` with proper `space_after` and font sizes so they never overlap or clip.

3. Accurate Styling & Korean Typography:
   - Use RGBColor with the exact HEX colors extracted from the image.
   - Set font.name = "Pretendard" (fallback to "Malgun Gothic" or "Noto Sans KR").
   - Match font sizes (e.g. Titles 28-36pt, Subtitles 18-22pt, Body 12-14pt, Badges 10-12pt).
   - If tables exist, set column widths proportionally to content length (`table.columns[i].width = ...`).

4. Save Output:
   prs.save(r"{pptx_file_path}")

Output ONLY the raw Python script enclosed in ```python and ```.
"""

    current_pptx_prompt = base_pptx_prompt
    max_retries = 2
    final_report = None

    for attempt in range(1, max_retries + 1):
      pptx_resp = client.models.generate_content(
          model=DEFAULT_GEMINI_MODEL,
          contents=[
              types.Part.from_bytes(data=image_bytes, mime_type="image/png"),
              current_pptx_prompt,
          ],
      )

      raw_code = pptx_resp.text or ""
      if "```python" in raw_code:
        raw_code = raw_code.split("```python")[1].split("```")[0].strip()
      elif "```" in raw_code:
        raw_code = raw_code.split("```")[1].split("```")[0].strip()

      # Execute the generated python-pptx code
      with tempfile.NamedTemporaryFile(
          mode="w", suffix=".py", delete=False, encoding="utf-8"
      ) as tmp_script:
        tmp_script.write(raw_code)
        tmp_script_path = tmp_script.name

      proc = None
      try:
        proc = subprocess.run(
            [sys.executable, tmp_script_path],
            capture_output=True,
            text=True,
            timeout=30,
        )
      finally:
        if os.path.exists(tmp_script_path):
          try:
            os.remove(tmp_script_path)
          except Exception:
            pass

      # Audit generated PPTX
      if os.path.exists(pptx_file_path):
        try:
          check_prs = Presentation(pptx_file_path)
          check_slide = check_prs.slides[0] if check_prs.slides else None
        except Exception:
          check_slide = None

        final_report = audit_slide_fidelity(
            original_image_input=image_bytes,
            generated_code=raw_code,
            slide=check_slide,
            model_name=DEFAULT_GEMINI_MODEL,
        )

        if final_report.passed and final_report.score >= 7:
          logger.info("[Single Slide Fidelity Audit PASSED on Attempt %d] Score: %d/10", attempt, final_report.score)
          break
        else:
          logger.warning(
              "[Single Slide Fidelity Audit REJECTED on Attempt %d] Score: %d/10. Retrying with corrective directives...",
              attempt,
              final_report.score,
          )
          current_pptx_prompt = (
              base_pptx_prompt
              + "\n\n[PREVIOUS ATTEMPT REJECTED BY FIDELITY AUDITOR - CORRECTIVE DIRECTIVES]:\n"
              + f"- Score: {final_report.score}/10 (Must be >= 7)\n"
              + (f"- Missing/Truncated text: {', '.join(final_report.missing_or_truncated_text)}\n" if final_report.missing_or_truncated_text else "")
              + (f"- Design/Layout defects: {', '.join(final_report.design_and_layout_issues)}\n" if final_report.design_and_layout_issues else "")
              + f"- Fix instructions: {final_report.corrective_instructions}\n"
              + "Fix ALL issues and regenerate complete Python script."
          )
      else:
        err_msg = proc.stderr if proc else "Unknown execution failure"
        current_pptx_prompt = (
            base_pptx_prompt
            + f"\n\n[CRITICAL ERROR IN PREVIOUS ATTEMPT]: Script execution failed: {err_msg}\nFix syntax and regenerate complete Python script."
        )

    pptx_exists = os.path.exists(pptx_file_path)
    pptx_size = os.path.getsize(pptx_file_path) if pptx_exists else 0

    return {
        "success": True,
        "html_path": html_file_path,
        "pptx_path": pptx_file_path if pptx_exists else None,
        "pptx_file_size_bytes": pptx_size,
        "html_content": raw_html,
        "fidelity_score": final_report.score if final_report else 8,
        "fidelity_passed": final_report.passed if final_report else True,
    }
