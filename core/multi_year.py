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


_COMPANY = re.compile(r"(?:M/S\.?\s*)?([A-Z][A-Za-z0-9&.,'\- ]{2,80}?\s(?:private\s+limited|pvt\.?\s*ltd\.?|limited|ltd\.?|llp))\b",
                      re.I)
_LEGAL_FORM = re.compile(r"\b(?:private|pvt|limited|ltd|llp|m/s|the|company|co)\b\.?", re.I)


def named_company(source):
    """The company name a file's statements print at their top ("Lunkad Foods Private Limited"), else None."""
    if source.extraction is None:
        return None
    for page in source.extraction.pages:
        if page.page_type not in ("BS", "P&L"):
            continue
        for line in page.lines[:4]:
            m = _COMPANY.search(line.replace("\n", " "))
            if m:
                return re.sub(r"\s+", " ", m.group(1)).strip()
    return None


def same_company(a, b):
    """True when two company names are the same company ("Shirodkar Preci Comp Pvt Ltd" ~ "... Private Limited")."""
    from rapidfuzz import fuzz
    initials = lambda s: "".join(w[0] for w in re.findall(r"[A-Za-z]+", s or "")).lower()
    if any(re.sub(r"[^a-z]", "", x.lower()) in (initials(y), initials(_LEGAL_FORM.sub(" ", y)))
           for x, y in ((a, b), (b, a)) if x and len(x.split()) == 1):
        return True                                         # "LFPL" for "Lunkad Foods Private Limited"
    core = lambda s: re.sub(r"\s+", " ", _LEGAL_FORM.sub(" ", s or "")).strip().lower()
    a, b = core(a), core(b)
    return not a or not b or fuzz.token_set_ratio(a, b) >= 80


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

    # a year's Balance Sheet and P&L are each taken from the file where the year is current, if that file has
    # that statement (some reports omit the P&L); otherwise from the nearest file that shows it as comparative
    def has_figures(source, fy, statement):
        return any(i.statement == statement and i.fy_values.get(fy) for i in source.items)

    owner_by = {}
    for fy in all_fys:
        for statement in ("BS", "P&L"):
            current = [i for i, s in enumerate(sources) if s.current == fy and has_figures(s, fy, statement)]
            showing = [i for i, s in enumerate(sources) if has_figures(s, fy, statement)]
            if current:
                owner_by[(fy, statement)] = current[0]
            elif showing:
                owner_by[(fy, statement)] = min(showing, key=lambda i: fy_key(sources[i].current))
                if fy in owner and sources[owner[fy]].current == fy:
                    warnings.append(f"{sources[owner[fy]].name} has no {'Profit & Loss' if statement == 'P&L' else 'Balance Sheet'}"
                                    f" figures for {fy}; they are taken from {sources[owner_by[(fy, statement)]].name}.")

    merged, order = {}, []
    for index, source in enumerate(sources):
        for item in source.items:
            years = [fy for fy in all_fys if owner_by.get((fy, item.statement), owner[fy]) == index]
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
