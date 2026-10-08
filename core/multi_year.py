"""
Combines several financial statements of the same company (for example the
FY23-24, FY24-25 and FY25-26 PDFs) into one set of line items.

Which file a year is taken from
  * the file in which that year is the CURRENT year (its latest column);
  * a year that is only ever a comparative (the oldest year) comes from the file
    closest to it, i.e. the one with the earliest current year that shows it.
So a figure restated in a later year's comparatives does not replace the
figure from that year's own statements.

Items are merged by (statement, label, schedule): the same line from two files
becomes one row with each year's figure from its own file. Opening reserves are
kept for every year; the template uses only the earliest year's (later years
roll forward from profit), and each file's own reconciliation uses its own.
"""

import re
from copy import deepcopy
from dataclasses import dataclass, field

from config.schedule_map import FY_COLUMN_MAP
from core.routing import schedule_for

_ORDER = {fy: i for i, fy in enumerate(FY_COLUMN_MAP)}


def fy_key(fy):
    """Sort key for 'FY24' style labels, also outside the template's FY22-FY26 range."""
    m = re.match(r"FY(\d{2})$", fy or "")
    return int(m.group(1)) if m else 999


@dataclass
class Source:
    name: str
    items: list
    fy_cols: list                 # every year the file shows
    mult: float = 1.0             # document unit -> rupees
    unit: str = "Rs. (whole rupees)"
    extraction: object = None     # None for the structured Excel layout
    log: list = field(default_factory=list)

    @property
    def current(self):
        return max(self.fy_cols, key=fy_key) if self.fy_cols else None


@dataclass
class Combined:
    items: list
    fy_cols: list                 # ascending
    owner: dict                   # fy -> index of the source it is taken from
    warnings: list = field(default_factory=list)

    def years_from(self, index):
        return [fy for fy, i in self.owner.items() if i == index]


def _norm(text):
    return re.sub(r"[^a-z0-9]", "", (text or "").lower())


def combine(sources):
    """-> Combined. sources: list[Source], in any order."""
    warnings = []
    all_fys = sorted({fy for s in sources for fy in s.fy_cols}, key=fy_key)
    owner = {}
    for fy in all_fys:
        current_in = [i for i, s in enumerate(sources) if s.current == fy]
        if current_in:
            owner[fy] = current_in[0]
            if len(current_in) > 1:
                names = ", ".join(sources[i].name for i in current_in)
                warnings.append(f"{fy} is the current year of more than one file ({names}); "
                                f"'{sources[current_in[0]].name}' is used.")
            continue
        showing = [i for i, s in enumerate(sources) if fy in s.fy_cols]
        owner[fy] = min(showing, key=lambda i: fy_key(sources[i].current))

    merged, order = {}, []
    for index, source in enumerate(sources):
        years = [fy for fy in all_fys if owner[fy] == index]
        for item in source.items:
            values = {fy: item.fy_values.get(fy) for fy in years if item.fy_values.get(fy) is not None}
            if not values:
                continue
            key = (item.statement, _norm(item.line_item), schedule_for(item) or "", item.statement_tag)
            target = merged.get(key)
            if target is None:
                target = deepcopy(item)
                target.fy_values = {fy: None for fy in all_fys}
                target.alt_values = {}
                merged[key] = target
                order.append(key)
            else:
                target.is_new = target.is_new or item.is_new
                target.flags = list(dict.fromkeys(target.flags + item.flags))
            for fy, value in values.items():
                current = target.fy_values.get(fy)
                target.fy_values[fy] = value if current is None else current + value
            for fy in years:
                if fy in item.alt_values:
                    target.alt_values[fy] = item.alt_values[fy]
    return Combined([merged[k] for k in order], all_fys, owner, warnings)
