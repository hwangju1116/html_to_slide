"""Standalone HTML-to-PPTX Native Presentation Parser & MCP Server Package."""

from app.html_to_pptx_converter import (
    audit_and_resolve_slide_collisions,
    capture_html_slides,
    convert_html_to_pptx,
)
from app.slide_digitizer import digitize_slide_image

__all__ = [
    "audit_and_resolve_slide_collisions",
    "capture_html_slides",
    "convert_html_to_pptx",
    "digitize_slide_image",
]
