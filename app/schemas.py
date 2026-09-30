from typing import List, Optional
from pydantic import BaseModel, Field


class SlideFidelityReport(BaseModel):
    """Report evaluating quality, zero-omission content completeness, and shape-text fidelity of generated slides."""

    passed: bool = Field(
        description="True if the generated slide meets all quality and fidelity standards, False otherwise."
    )
    score: int = Field(
        description="Overall fidelity & design score from 1 to 10. Pass threshold is >= 7."
    )
    zero_omission_verified: bool = Field(
        default=True,
        description="True if 100% of all titles, subheadings, bullet points, numbers, metrics, badges, and tables are preserved verbatim.",
    )
    missing_or_truncated_text: List[str] = Field(
        default_factory=list,
        description="List of any missing words, numbers, badges, or truncated phrases detected.",
    )
    design_and_layout_issues: List[str] = Field(
        default_factory=list,
        description="List of detected design issues (e.g. text clipping, disjoint overlapping boxes, incorrect font sizes, margin overflows).",
    )
    corrective_instructions: str = Field(
        default="",
        description="Concrete, actionable Python/PPTX code instructions to fix the detected issues on retry.",
    )


class TextRun(BaseModel):
    text: str = Field(description="Sub-string text segment")
    is_bold: bool = Field(default=False, description="Whether this segment is bold")
    color_hex: Optional[str] = Field(default=None, description="HEX color e.g. #2563EB")
    is_code: bool = Field(default=False, description="Whether this segment is inline code or monospace")


class SlideNativeElement(BaseModel):
    element_type: str = Field(
        description=(
            "One of: 'CONTAINER', 'SUB_CARD', 'HEADER_BADGE', 'TEXT_BOX', 'ICON_BADGE', "
            "'STAT_METRIC', 'FLOW_NODE', 'FLOW_ARROW', 'BRIDGE_CONNECTOR', 'PIPELINE_BANNER', "
            "'PIPELINE_STEP', 'CODE_BLOCK', 'CALLOUT_BOX', 'FOOTNOTE', 'TABLE', 'DIVIDER', 'IMAGE'"
        )
    )
    text: Optional[str] = Field(
        default="", description="Verbatim text content with lines and bullet indents"
    )
    text_runs: Optional[List[TextRun]] = Field(
        default=None, description="Detailed inline text runs with rich styles"
    )
    # Normalized Coordinates (0.0 to 100.0%)
    left_pct: float = Field(description="Left X percentage of slide width (0-100)")
    top_pct: float = Field(description="Top Y percentage of slide height (0-100)")
    width_pct: float = Field(description="Width percentage of slide width (0-100)")
    height_pct: float = Field(description="Height percentage of slide height (0-100)")
    # Typography
    font_size_pt: Optional[float] = Field(default=14.0, description="Font size in points")
    is_bold: Optional[bool] = Field(default=False)
    font_color_hex: Optional[str] = Field(
        default="#0F172A", description="HEX text color, e.g. #1E293B"
    )
    alignment: Optional[str] = Field(
        default="LEFT", description="'LEFT', 'CENTER', 'RIGHT'"
    )
    # Container / Shape styling
    bg_color_hex: Optional[str] = Field(
        default=None, description="Fill background HEX color, e.g. #F8FAFC"
    )
    border_color_hex: Optional[str] = Field(
        default=None, description="Border stroke HEX color, e.g. #E2E8F0"
    )
    border_width_pt: Optional[float] = Field(default=1.0)
    border_radius_pt: Optional[float] = Field(default=0.0)
    shape_type: Optional[str] = Field(
        default="ROUNDED_RECTANGLE",
        description="PowerPoint shape: 'ROUNDED_RECTANGLE', 'RECTANGLE', 'RIGHT_ARROW', 'LINE', 'OVAL', 'TABLE'",
    )
    icon_symbol: Optional[str] = Field(
        default=None, description="Emoji or symbol e.g. ✅, ⚠️, 💡, ➔, ──►, 🔒, ⚡, 🏦, 🚫"
    )
    accent_bar_color_hex: Optional[str] = Field(
        default=None, description="Top accent bar color if applicable"
    )
    z_index: Optional[int] = Field(
        default=1,
        description="Z-index rendering order: 0=background, 1=container, 2=sub-card, 3=text/badge/arrow",
    )
    parent_id: Optional[str] = Field(
        default=None, description="ID of parent container if nested"
    )
    # Sub-card / Pipeline / List attributes
    sub_items: Optional[List[str]] = Field(
        default=None, description="Bullet items or nested points inside a card/sub-card"
    )
    callout_type: Optional[str] = Field(
        default=None, description="'SUMMARY', 'WARNING', 'NOTE', 'KEY_TAKEAWAY'"
    )
    # Table Matrix (if element_type == 'TABLE')
    table_matrix: Optional[List[List[str]]] = Field(
        default=None, description="2D matrix of table cell strings"
    )
    table_header_bg_hex: Optional[str] = Field(default="#F1F5F9")


class SlideNativeDecomposition(BaseModel):
    slide_background_hex: str = Field(
        default="#FFFFFF", description="Canvas background HEX color"
    )
    elements: List[SlideNativeElement] = Field(
        description="Ordered list of native shapes, containers, textboxes, and tables"
    )
