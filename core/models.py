"""Shared data structures passed between pipeline stages."""

from dataclasses import dataclass, field


@dataclass
class LineItem:
    statement: str          # "P&L" / "BS" / "CF"
    line_item: str          # raw name as it appears in the Annual Report
    statement_tag: str      # controlled-vocabulary tag (blank for P&L expense rows)
    fy_values: dict         # {"FY22": 12.3, "FY23": None, ...} -- None means not provided

    # populated later by classifier.py (P&L expense rows only)
    category: str = ""
    confidence: str = ""    # "Matched" / "Uncertain"
    matched_keywords: list = field(default_factory=list)

    # populated later by computation.py
    common_size: dict = field(default_factory=dict)   # {"FY22": 12.3, ...} percent
    yoy_change_pct: dict = field(default_factory=dict)  # {"FY23": 5.1, ...} % change vs prior FY

    # provenance + placement, filled by core/pdf_mapper.py (empty for plain Excel input)
    source: str = ""          # "excel" | "pdf-note" | "pdf-face"
    pdf_heading: str = ""     # heading exactly as the PDF presents it (note title / section)
    pdf_page: int = 0
    is_new: bool = False      # not in the taxonomy -> goes in a spare row, highlighted yellow
    matched_item: str = ""    # taxonomy item / alias it matched
    match_score: float = 0.0
    schedule_override: str = ""  # user's choice in the review table; beats automatic routing
    flags: list = field(default_factory=list)
    alt_values: dict = field(default_factory=dict)  # figure the note shows when it differs from the face figure (rupees)


@dataclass
class Finding:
    """A single flagged observation, destined for the Analysis sheet and/or the IRL."""
    section: str             # "Executive Summary" / "YoY Variance" / "Common Size Analysis" /
                              # "Ratio Commentary" / "Data Integrity"
    area: str                # Financial Statement Area, e.g. "COGS", "Trade Receivables"
    observation: str
    reason_for_concern: str = ""
    info_required: str = ""
    needs_irl: bool = True   # False for purely descriptive Executive Summary lines
