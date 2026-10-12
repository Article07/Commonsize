"""
Row map for the schedules on the "Sch." sheet of
`Reference file/Common size format_Mfg-Trading firm.xlsx`.

The numbers are NOT hand-maintained: they live in config/schedule_layout.json,
which tools/build_master_template.py writes every time it rebuilds the
Reference file (taxonomy rows + hidden spare rows added to the original), so
the map can never drift from the workbook. `python tools/verify_layout.py`
re-checks the JSON against the actual file.

Schedule patterns:
  - "total_last":  item rows, then a Total row below with SUM(first:last).
  - "total_first": a combined title+Total row, then item rows below it.
  - "single_cell": a fixed input cell, not a growable list (Share Capital,
    Reserves opening balance, Securities Premium, ...).

`spare_rows` is the number of HIDDEN blank rows kept inside a growable
schedule's SUM range, immediately above its last item row, i.e. rows
[last_item_row - spare_rows, last_item_row - 1]. A line item that is not in
the taxonomy is written into one of them (row unhidden, highlighted yellow)
so it is included in the schedule total and every statement downstream.
"""

import json
import re
from pathlib import Path

_LAYOUT_PATH = Path(__file__).with_name("schedule_layout.json")

if not _LAYOUT_PATH.exists():
    raise FileNotFoundError(
        f"{_LAYOUT_PATH} is missing. Generate it with: python tools/build_master_template.py"
    )

_layout = json.loads(_LAYOUT_PATH.read_text(encoding="utf-8"))

SHEET_NAME = _layout["sheet"]

# Company data's FY label -> the template's fixed FY22-FY26 columns.
FY_COLUMN_MAP = {
    "FY22": "C",
    "FY23": "D",
    "FY24": "E",
    "FY25": "F",
    "FY26": "G",
}



def column_map(fy_columns):
    """
    {FY: template column} for the years being written. The template has five year columns (C..G), labelled
    FY22..FY26; when the years do not fit there (an FY2019-20 filing), the five columns start at the earliest
    year instead (from the latest five when the years span more) and the workbook's year headings say so.
    """
    nums = sorted({int(f[2:]) for f in fy_columns if re.fullmatch(r"FY\d\d", f or "")})
    if not nums or {f"FY{n:02d}" for n in nums} <= set(FY_COLUMN_MAP):
        return dict(FY_COLUMN_MAP)
    start = max(nums[0], nums[-1] - (len(FY_COLUMN_MAP) - 1))
    return {f"FY{(start + i) % 100:02d}": col for i, col in enumerate(FY_COLUMN_MAP.values())}


ALL_SCHEDULES = _layout["schedules"]
ALL_SCHEDULE_ORDER = _layout["order"]


def spare_row_numbers(schedule_key):
    """Rows of the hidden spare slots for a growable schedule (top to bottom)."""
    s = ALL_SCHEDULES[schedule_key]
    return list(range(s["last_item_row"] - s["spare_rows"], s["last_item_row"]))
