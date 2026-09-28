"""Pydantic schemas for the HTML-to-PPTX MCP Parser & Fidelity Critic."""

from typing import List
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
