"""
The review page's statements: the Income Statement and the Balance Sheet laid out as the uploaded file prints
their face, with, for every year, the figure the file prints ("As per financials"), the figure the common size adds
up to ("Common size") and their difference.

    * a face line's common-size figure is the sum of the line items read for it (each item remembers its face line,
      LineItem.face_line, and whether the face prints it the other way round, LineItem.face_sign);
    * the reserves line is the closing reserves the common size rolls forward to;
    * a total is the sum of the lines it covers -- found from the figures the file prints (walking back from the
      total until the printed lines add up to it) -- else, for profit before tax / profit / balance-sheet totals,
      the app's own computation of that total.
No Streamlit in here.
"""

import re
from dataclasses import dataclass, field

from core.pdf_mapper import _SUMMARY_LINE
from core.pipeline import _PAT_LABEL, computed_totals, tolerance

RESERVE_SCHEDULES = {"Reserves and Surplus", "Reserves Adjustments", "Securities Premium", "Other Reserves"}
_KNOWN_TOTALS = [            # (key in pipeline.computed_totals, statement, pattern on the printed label): the exact lines only
    ("total_income", "P&L", r"^total\s*(?:income|revenue)\W*$"),
    ("total_expenses", "P&L", r"^total\s*expenses?\W*$"),
    ("pbt", "P&L", r"^(?:net\s*|total\s*)?profit\s*(?:/?\s*\(?loss\)?\s*)?before\s*tax\W*$"),
    ("tax", "P&L", r"^total\s*tax|^tax\s*expense(?!.*discon|.*discount)"),
    ("pat", "P&L", None),
    ("total_liab", "BS", r"^total\s*equity\s*(?:and|&)\s*liabilities|^total\W*$"),
    ("total_assets", "BS", r"^total\s*assets"),
]
_INCOME = {"Revenue", "Other Income"}
_NOT_AMOUNT = re.compile(r"\bratios?\b|%|per\s*(?:equity\s*)?share|\beps\b", re.I)
OTHER_LINES = "Other lines read (not a line of the face)"
ROUNDING = 1e-9                     # relative: sums of paise drift by 0.01 on figures of crores


@dataclass
class StatementRow:
    label: str
    is_total: bool
    as_per: dict                       # {fy: figure in the display unit | None}
    common_size: dict                  # {fy: figure | None}
    items: list = field(default_factory=list)   # the LineItems behind a line (editable on the page)
    reserves: bool = False             # the reserves line: its figure is the rolled-forward closing reserves
    extra: bool = False                # lines read that the face does not print as a line: nothing to compare
    tol: float = 0.0                   # rounding of the file's own unit (display units): not a difference

    def difference(self, fy):
        """Common size less as per financials; None when neither has a figure, 0 within rounding (paise)."""
        a, c = self.as_per.get(fy), self.common_size.get(fy)
        if self.extra or (a is None and c is None):
            return None
        d = (c or 0.0) - (a or 0.0)
        return 0.0 if abs(d) <= max(self.tol, ROUNDING * abs(a or 0.0)) else d


def norm(label):
    return re.sub(r"[^a-z0-9]", "", str(label or "").lower())


def _face_rows(extraction, statement):
    # amounts only: "EBITDA ratio %", "Earnings per share" are not lines the common size adds up
    rows = [r for r in extraction.rows if r.statement == statement and not _NOT_AMOUNT.search(r.label)]
    if statement == "P&L":
        # the face ends at the profit for the period (earnings per share, OCI and the rest follow it)
        end = next((k for k, r in enumerate(rows) if _PAT_LABEL.search(r.label)
                    and not re.search(r"before|continuing", r.label, re.I)), None)
        rows = rows[:end + 1] if end is not None else rows
    return rows


def build(statement, sources, owner, years, items, display_mult):
    """
    -> [StatementRow] for 'BS' or 'P&L'. sources: list of Source (with .extraction, .mult); owner: {fy: source index}
    (the file each year is taken from); items: the combined LineItems; display_mult: rupees per display unit.
    """
    readable = [s for s in sources if s.extraction is not None]
    if not readable or not years:
        return []
    latest = max(years, key=lambda f: int(f[2:]))
    main = sources[owner.get(latest, 0)] if sources[owner.get(latest, 0)].extraction is not None else readable[0]

    # the face lines, in the order the main file prints them; lines only another year's file has are added after.
    # A label can be printed twice ("TOTAL" under liabilities and under assets): its n-th print is its own line.
    def keyed(src):
        count, out = {}, []
        for r in _face_rows(src.extraction, statement):
            n = count[norm(r.label)] = count.get(norm(r.label), -1) + 1
            out.append(((norm(r.label), n), r))
        return out

    order, seen = [], set()
    for src in [main] + [s for s in readable if s is not main]:
        for key, r in keyed(src):
            if key not in seen:
                seen.add(key)
                order.append((key, r))

    def printed(key, fy):
        src = sources[owner[fy]] if fy in owner and sources[owner[fy]].extraction is not None else main
        match = next((x for k, x in keyed(src) if k == key), None)
        v = match.values.get(fy) if match is not None else None
        return None if v is None else v * src.mult / display_mult

    as_per = {key: {fy: printed(key, fy) for fy in years} for key, _ in order}
    prints = {}
    for key in as_per:
        prints.setdefault(key[0], []).append(key)
    by_label = {}
    for it in items:
        if it.statement == statement:
            by_label.setdefault(norm(it.face_line), []).append(it)
    linked = {}
    for label, group in by_label.items():
        keys = prints.get(label)
        if not keys:
            linked.setdefault(None, []).extend(group)
            continue
        # printed more than once: the print whose figures are nearest what its lines add up to
        best = min(keys, key=lambda k: sum(abs((as_per[k].get(fy) or 0.0) * display_mult
                                               - sum(_face_value(i, fy) for i in group)) for fy in years))
        linked[best] = group
    totals = computed_totals([i for i in items if "excluded_by_user" not in i.flags],
                             sorted(years, key=lambda f: int(f[2:])))

    rows = []
    for key, r in order:
        its = linked.pop(key, [])
        if not its and all(v is None for v in as_per[key].values()):
            continue
        is_total = not its and (r.is_total or bool(_SUMMARY_LINE.search(r.label)))
        row = StatementRow(r.label, is_total, as_per[key], {}, its)
        row.reserves = any(_schedule(i) in RESERVE_SCHEDULES for i in its)
        rows.append(row)
    # lines read from the notes / from text the face does not show as a line of its own
    rest = [i for group in linked.values() for i in group]
    if rest:
        rows.append(StatementRow(OTHER_LINES, False, {fy: None for fy in years}, {}, rest, extra=True))

    tol = max(tolerance(s.extraction.unit_multiplier, s.extraction.whole_units) for s in readable) / display_mult
    for row in rows:
        row.tol = tol
        if row.reserves:
            row.common_size = {fy: totals[fy]["reserves"] / display_mult for fy in years}
        elif not row.is_total:
            row.common_size = {fy: sum(_face_value(i, fy) for i in row.items) / display_mult for fy in years}
    for k, row in enumerate(rows):
        if row.is_total:
            row.common_size = _total_value(rows, k, years, statement, totals, display_mult)
    return rows


def _schedule(item):
    from core.routing import schedule_for
    return item.schedule_override or schedule_for(item)


def _face_value(item, fy):
    if "excluded_by_user" in item.flags:
        return 0.0
    return (item.fy_values.get(fy) or 0.0) * item.face_sign


def _total_value(rows, k, years, statement, totals, display_mult):
    """A total: the lines above it that its printed figure is the sum of; else the app's own computed total."""
    parts = _components(rows, k, years, {})
    if parts is not None and all(rows[j].common_size.get(fy) is not None for j in parts for fy in years):
        return {fy: sum(rows[j].common_size[fy] for j in parts) for fy in years}
    label = rows[k].label
    for key, st, pattern in _KNOWN_TOTALS:
        if st != statement:
            continue
        # the profit for the period is the last profit line (a consolidated face prints the group's profit from
        # continuing operations before the minority's share and the associates)
        hit = (_is_pat(label) and not any(_is_pat(r.label) for r in rows[k + 1:])) if pattern is None \
            else re.search(pattern, label, re.I)
        if hit:
            return {fy: totals[fy][key] / display_mult for fy in years}
    if statement == "P&L":
        # a profit sub-total ("Profit before exceptional items and tax", "Net profit before partners' remuneration"):
        # what the file prints, moved by the differences of the lines above it (income adds, expenses deduct)
        out = {}
        for fy in years:
            printed = rows[k].as_per.get(fy)
            if printed is None:
                out[fy] = None
                continue
            moved = 0.0
            for row in rows[:k]:
                if row.is_total or row.reserves or row.extra:
                    continue
                d = row.difference(fy) or 0.0
                moved += d * _effect(row)
            out[fy] = printed + moved
        return out
    return {fy: None for fy in years}


def _is_pat(label):
    return bool(_PAT_LABEL.search(label)) and not re.search(r"before|discontinu", label, re.I)


def _effect(row):
    """+1 when a P&L line adds to profit as printed, -1 when it is deducted."""
    if row.items:
        it = row.items[0]
        return (1 if _schedule(it) in _INCOME else -1) * it.face_sign
    return 1 if re.search(r"income|revenue|sales|^add", row.label, re.I) else -1


def _span_start(rows, k, years, memo):
    """First row a total covers, through its sub-totals (a total of totals reaches back to their lines)."""
    inner = _components(rows, k, years, memo)
    if not inner:
        return k
    return min(_span_start(rows, i, years, memo) if rows[i].is_total else i for i in inner)


def _components(rows, k, years, memo):
    """Indexes of the rows above total k whose printed figures add up to it (a sub-total stands for its lines)."""
    if k in memo:
        return memo[k]
    memo[k] = None
    target = rows[k].as_per
    if all(target.get(fy) is None for fy in years):
        return None
    if all(not target.get(fy) for fy in years):
        # a total the file prints as nil: the nil lines right above it (none, for a heading with nothing under it)
        acc, j = [], k - 1
        while j >= 0 and not rows[j].is_total and not any(rows[j].as_per.get(fy) for fy in years):
            acc.append(j)
            j -= 1
        memo[k] = acc
        return acc
    acc, j = [], k - 1
    while j >= 0 and len(acc) < 80:
        acc.append(j)
        j = (_span_start(rows, j, years, memo) if rows[j].is_total else j) - 1
        ok = True
        for fy in years:
            t = target.get(fy)
            if t is None:
                continue
            got = sum(rows[i].as_per.get(fy) or 0.0 for i in acc)
            if abs(got - t) > 0.011 + 0.005 * len(acc) + 1e-6 * abs(t):
                ok = False
                break
        if ok and any(rows[i].as_per.get(fy) for i in acc for fy in years):
            memo[k] = acc
            return acc
    return None
