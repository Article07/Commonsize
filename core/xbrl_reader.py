"""
Reads the MCA AOC-4 XBRL rendering of financial statements ("<COMPANY> Standalone Financial Statements for
period 01/04/2023 to 31/03/2024", sections "[110000] Balance sheet", "[210000] Statement of profit and loss" ...).

The layout is fixed by the MCA taxonomy, so it is read by structure rather than like a printed statement:
  * each section is a set of tables; a table's columns are the dates under "Unless otherwise specified ..."
    (31/03/2024 an instant, "01/04/2023 to 31/03/2024" a period);
  * "[Abstract]" lines are headings, "Total ..." lines are totals, a figure may carry a footnote marker
    "(A) 36.61" whose footnote gives its breakdown ("(A) Income tax assets : 36.61");
  * the face statements give the figures; "[500100] Subclassification ... income and expenses" itemises the
    P&L lines, a footnote breakdown itemises a line further when it adds up; "[400200] Statement of changes in
    equity" gives the opening balance of Other equity and its movements other than the profit.
Line names go through the same taxonomy / resolver as every other input.
"""

import re
from dataclasses import dataclass, field

from core.pdf_mapper import MappingResult, _PNL_ITEM_BELOW, _SAME_IS_LINE, _SUMMARY_LINE, _make_item, _side_rules
from core.pdf_reader import ExtractedRow, ExtractionResult
from core.routing import schedule_for

_SECTION = re.compile(r"^\[(\d{6})\]\s*(.+)$")
_DATE = re.compile(r"^(\d{2})/(\d{2})/(\d{4})$")
_UNIT = re.compile(r"all monetary values are in\s+(?:(millions|lakhs|thousands|crores|hundreds)\s+of\s+)?INR", re.I)
_UNITS = {"millions": 1_000_000, "lakhs": 100_000, "crores": 10_000_000, "thousands": 1_000, "hundreds": 100, None: 1}
_NUMBER = re.compile(r"^\(?-?[\d,]*\d(?:\.\d+)?\)?$")
_MARK = re.compile(r"^\(([A-Z]{1,2})\)$")
_NOT_A_ROW = re.compile(r"\[(?:abstract|table|axis|member|line items|text\s*block|see below)\]|^whether\b|^textual information"
                        r"|^description\b|^\[inr/shares\]|\[(?:inr/shares|pure)\]", re.I)
_PAGE_HEADER = re.compile(r"standalone financial statements for period|consolidated financial statements for period", re.I)
TOL = 0.011
# section codes: Ind AS taxonomy, then the older Indian GAAP (Schedule III) taxonomy
_BS_CODES, _PNL_CODES = ("110000", "100100"), ("210000", "100200")
_NOTE_CODES, _EQUITY_CODES = ("500100", "300500"), ("400200", "200200")
_PNL_MUST = r"^revenue from|^total income|^total revenue"


def is_xbrl(pages):
    """True for the MCA XBRL rendering: a '[110000] Balance sheet' section and the 'for period' page header."""
    text = [l for p in pages for l in p.lines[:4]]
    codes = {m.group(1) for p in pages for l in p.lines if (m := re.match(r"^\[(\d{6})\]", l))}
    return (any(_PAGE_HEADER.search(l) for l in text)
            and bool(codes & set(_BS_CODES)) and bool(codes & set(_PNL_CODES)))


@dataclass
class Row:
    label: str
    values: list                      # [(value, footnote letter | None)] left to right
    heading: str                      # the nearest "[Abstract]" heading above
    page: int
    is_total: bool = False


@dataclass
class Table:
    columns: list                     # [("I" | "D", "FY24")] instant / duration, left to right
    rows: list = field(default_factory=list)
    footnotes: dict = field(default_factory=dict)   # letter -> [(name, value)]

    def values_of(self, row, kinds=None):
        """{FY: value} of the first columns of the given kinds (a row omits columns it has no value for)."""
        cols = [c for c in self.columns if kinds is None or c[0] in kinds]
        if kinds is None and len(row.values) != len(self.columns):
            durations = [c for c in self.columns if c[0] == "D"]
            if len(row.values) == len(durations):
                cols = durations
        out = {}
        for (value, _), (_, fy) in zip(row.values, cols):
            out.setdefault(fy, value)
        return out

    def marks_of(self, row):
        out = {}
        for (_, mark), (_, fy) in zip(row.values, self.columns):
            if mark:
                out.setdefault(fy, mark)
        return out


def _fy(day, month, year):
    return f"FY{(year if month <= 3 else year + 1) % 100:02d}"


def _number(tok):
    t = tok.replace(",", "")
    neg = t.startswith("(") and t.endswith(")")
    t = t.strip("()")
    try:
        v = float(t)
    except ValueError:
        return None
    return -v if neg else v


def _split_values(line):
    """'Other non-current assets (A) 36.61 (B) 36.47' -> ('Other non-current assets', [(36.61, 'A'), (36.47, 'B')])"""
    tokens = line.split()
    vals = []
    while tokens:
        t = tokens[-1]
        if _NUMBER.match(t) and _number(t) is not None:
            vals.append([_number(t), None])
            tokens.pop()
        elif _MARK.match(t) and vals and vals[-1][1] is None:
            vals[-1][1] = _MARK.match(t).group(1)
            tokens.pop()
        else:
            break
    vals.reverse()
    return " ".join(tokens), [tuple(v) for v in vals]


def _content_lines(pages):
    """All lines with the per-page header (page number, company / period line) removed: [(page, line)]."""
    out = []
    for p in pages:
        lines = list(p.lines)
        while lines and (re.fullmatch(r"\d{1,4}", lines[0].strip()) or _PAGE_HEADER.search(lines[0])):
            lines.pop(0)
        out += [(p.page, l.strip()) for l in lines if l.strip()]
    return out


def _parse_footnotes(text):
    """'(A) Income tax assets : 36.61 (B) ...' -> {'A': [('Income tax assets', 36.61)], ...}"""
    out = {}
    parts = re.split(r"(?:^|\s)\(([A-Z]{1,2})\)\s", " " + text)
    for letter, body in zip(parts[1::2], parts[2::2]):
        pairs = re.findall(r"\s*(.+?)\s*:\s*(-?[\d,]*\.?\d+)(?=\s|$)", body)
        out[letter] = [(name.strip(" ,;"), _number(value)) for name, value in pairs if _number(value) is not None]
    return out


def _sections(lines):
    """{code: (title, [Table])}"""
    sections, code = {}, None
    table, mode, header, foot = None, None, [], []

    def close_footnotes():
        nonlocal foot
        if table is not None and foot:
            table.footnotes.update(_parse_footnotes(" ".join(foot)))
        foot = []

    pending = None                        # a label wrapped over two lines
    for page, line in lines:
        m = _SECTION.match(line)
        if m:
            close_footnotes()
            code = m.group(1)
            sections.setdefault(code, (m.group(2), []))
            table, mode, pending = None, None, None
            continue
        if code is None:
            continue
        if re.match(r"^unless otherwise specified", line, re.I):
            close_footnotes()
            table, mode, header, pending = Table([]), "header", [], None
            sections[code][1].append(table)
            continue
        if table is None:
            continue
        if mode == "header":
            tokens = line.split()
            # "31/03/2024 31/03/2023" on one line, or "01/04/2023" / "to" / "31/03/2024" one per line
            if tokens and all(_DATE.match(t) or (t.lower() == "to" and (header or k)) for k, t in enumerate(tokens)):
                header += tokens
                continue
            if not header:
                continue                  # axis / member names above the dates
            cols, k = [], 0
            while k < len(header):
                if k + 2 < len(header) and header[k + 1].lower() == "to":
                    d, mth, y = map(int, _DATE.match(header[k + 2]).groups())
                    cols.append(("D", _fy(d, mth, y)))
                    k += 3
                else:
                    d, mth, y = map(int, _DATE.match(header[k]).groups())
                    cols.append(("I", _fy(d, mth, y)))
                    k += 1
            table.columns, mode = cols, "rows"
        if line.lower() == "footnotes":
            mode = "footnotes"
            continue
        if mode == "footnotes":
            foot.append(line)
            continue
        if line.lower().endswith("[abstract]"):
            heading = re.sub(r"\s*\[abstract\]$", "", line, flags=re.I)
            table.rows.append(Row(heading, [], heading, page))       # a heading marker (no values)
            pending = None
            continue
        if _NOT_A_ROW.search(line):
            pending = None
            continue
        label, values = _split_values(line)
        if pending is not None and line[:1].islower():
            label = f"{pending} {label}".strip()
        pending = None
        if not values:
            pending = label
            continue
        heading = next((r.heading for r in reversed(table.rows) if not r.values), "")
        table.rows.append(Row(label, values, heading, page, bool(re.match(r"^total\b", label, re.I))))
    close_footnotes()
    return sections


def _main_table(sections, codes, must):
    for code in codes:
        for t in sections.get(code, ("", []))[1]:
            if any(re.search(must, r.label, re.I) for r in t.rows if r.values):
                return t
    return None


def _breakdown(table, row, fy_columns):
    """{name: {fy: value}} from the row's footnotes, when they add up to the row in every year it has figures."""
    marks = table.marks_of(row)
    values = table.values_of(row)
    if not marks:
        return None
    out = {}
    for fy in fy_columns:
        v = values.get(fy)
        if not v:
            continue
        pairs = table.footnotes.get(marks.get(fy, ""), [])
        if not pairs or abs(sum(x for _, x in pairs) - v) > TOL + 0.005 * len(pairs):
            return None
        for name, x in pairs:
            out.setdefault(name, {})[fy] = x
    return {n: v for n, v in out.items() if any(v.values())} or None


# [500100] group headings -> the face line they itemise
_PNL_GROUPS = [
    (r"disclosure of revenue from operations$", r"^revenue from operations", r"^total revenue from operations$"),
    (r"disclosure of other income", r"^other income", r"^total other income"),
    (r"disclosure of finance cost", r"^finance costs?", r"^total finance costs?"),
    (r"^employee benefit expense$", r"^employee benefits? expense", r"^total employee benefit expense"),
    (r"^depreciation, depletion and amortisation expense$", r"^depreciation", r"^total depreciation"),
    (r"^breakup of other expenses$", r"^other expenses", r"^total other expenses"),
]


def _note_groups(table):
    """{face pattern: [Row]} -- the item rows of each [500100] group."""
    out = {}
    if table is None:
        return out
    rows = table.rows
    for head, face, total in _PNL_GROUPS:
        start = next((k for k, r in enumerate(rows) if not r.values and re.search(head, r.label, re.I)), None)
        if start is None:
            continue
        items, since = [], []
        for r in rows[start + 1:]:
            if r.values and re.search(total, r.label, re.I):
                break
            if not r.values:
                since = []
                continue
            if r.is_total:
                if not since:
                    items.append(r)          # a sub-heading with only its total
                continue
            items.append(r)
            since.append(r)
        out[face] = items
    return out


def read(pages):
    """-> (ExtractionResult, MappingResult) for an XBRL rendering."""
    lines = _content_lines(pages)
    sections = _sections(lines)
    bs = _main_table(sections, _BS_CODES, r"^total assets")
    pnl = _main_table(sections, _PNL_CODES, _PNL_MUST)
    if bs is None or pnl is None:
        raise ValueError("This looks like an AOC-4 XBRL statement, but its Balance Sheet or Profit & Loss could not be read.")
    unit = next((m for _, l in lines if (m := _UNIT.search(l))), None)
    key = unit.group(1).lower() if unit and unit.group(1) else None
    mult = _UNITS.get(key, 1)
    unit_label = {1_000_000: "Rs. in millions", 100_000: "Rs. in lakhs", 10_000_000: "Rs. in crores",
                  1_000: "Rs. in '000", 100: "Rs. in hundreds", 1: "Rs. (whole rupees)"}[mult]
    fy_columns = [fy for kind, fy in pnl.columns if kind == "D"][:2]
    company = next((re.split(_PAGE_HEADER, p.lines[1])[0].strip() for p in pages if len(p.lines) > 1
                    and _PAGE_HEADER.search(p.lines[1])), "")

    rows = []
    for statement, table in (("BS", bs), ("P&L", pnl)):
        for r in table.rows:
            if r.values:
                values = {fy: v for fy, v in table.values_of(r).items() if fy in fy_columns}
                rows.append(ExtractedRow(r.page, statement, r.heading, "", "", r.label, values, r.label, r.is_total))
    for p in pages:
        if any(re.match(rf"^\[(?:{'|'.join(_BS_CODES)})\]", l) for l in p.lines):
            p.page_type = "BS"
        elif any(re.match(rf"^\[(?:{'|'.join(_PNL_CODES)})\]", l) for l in p.lines):
            p.page_type = "P&L"
    extraction = ExtractionResult(rows, fy_columns, mult, unit_label, pages,
                                  checks=[f"AOC-4 XBRL statements{(' of ' + company) if company else ''}: read by section "
                                          "(Balance sheet, Profit and loss, notes on income and expenses, "
                                          "changes in equity / reserves)."],
                                  warnings=[], whole_units=False)
    extraction.xbrl = sections
    return extraction, map_xbrl(extraction, mult)


def _soce(sections, fy_columns, pat):
    """(opening Other equity of the earliest year | None, {fy: movements other than profit})"""
    table = _main_table(sections, _EQUITY_CODES, r"other equity at end of period|reserves at end of period")
    if table is None:
        return None, {}
    end = next((r for r in table.rows if r.values and re.search(r"(?:other equity|reserves) at end of period", r.label, re.I)), None)
    change = next((r for r in table.rows if r.values and re.search(r"total increase \(decrease\) in equity|total changes in reserves",
                                                                    r.label, re.I)), None)
    profit = next((r for r in table.rows if r.values and re.search(r"^profit \(loss\) for period", r.label, re.I)), None)
    earliest = min(fy_columns, key=lambda f: int(f[2:]))
    before = f"FY{int(earliest[2:]) - 1:02d}"
    opening = table.values_of(end).get(before) if end is not None else None
    if opening is None and end is not None and change is not None:
        # no balance printed for the day before the earliest year: closing less the year's change
        closing, changed = table.values_of(end).get(earliest), table.values_of(change, kinds=("D",)).get(earliest)
        if closing is not None and changed is not None:
            opening = closing - changed
    moves = {}
    if change is not None:
        d_change = table.values_of(change, kinds=("D",))
        d_profit = table.values_of(profit, kinds=("D",)) if profit is not None else pat
        for fy in fy_columns:
            if d_change.get(fy) is not None:
                moves[fy] = d_change[fy] - (d_profit.get(fy) or 0.0)
    return opening, moves


def map_xbrl(extraction, multiplier=None):
    mult = extraction.unit_multiplier if multiplier is None else multiplier
    fy = list(extraction.fy_columns)
    sections = extraction.xbrl
    bs = _main_table(sections, _BS_CODES, r"^total assets")
    pnl = _main_table(sections, _PNL_CODES, _PNL_MUST)
    # the section may open with small tables ("Disclosure of other operating revenues"): the breakup is the
    # table holding the most of the group headings
    candidates = [t for code in _NOTE_CODES for t in sections.get(code, ("", []))[1]]
    notes = max(candidates, default=None, key=lambda t: sum(
        1 for r in t.rows if not r.values and any(re.search(h, r.label, re.I) for h, _, _ in _PNL_GROUPS)))
    groups = _note_groups(notes)
    result = MappingResult()

    def item(label, values, statement, heading, page, flags=(), source="pdf-face"):
        return _make_item(label, {k: v for k, v in values.items() if k in fy}, statement, heading, source, page,
                          list(flags), fy, mult)

    def add_with_breakdown(table, r, statement, heading, base):
        parts = _breakdown(table, r, fy)
        if not parts or re.search(r"share capital", r.label, re.I):
            result.items.append(base)
            return
        face_key = schedule_for(base)
        for name, values in parts.items():
            it = item(name, values, statement, r.label, r.page, ["from_footnote"], "pdf-note")
            own = schedule_for(it)
            if face_key and own != face_key and not any({own, face_key} <= g for g in _SAME_IS_LINE):
                it.schedule_override = face_key
                it.flags = [f for f in it.flags if f != "unplaced"]
            result.items.append(it)
        result.log.append(f"{r.label}: itemised from its footnote ({len(parts)} lines)")

    # ---- Balance sheet
    assets_side = None
    for r in bs.rows:
        if not r.values:
            if re.match(r"^assets$", r.label, re.I):
                assets_side = True
            elif re.match(r"^equity and liabilities$", r.label, re.I):
                assets_side = False
            continue
        if r.is_total or re.search(r"^other equity$|^reserves and surplus$|^non controlling interest$|equity attributable", r.label, re.I):
            continue
        values = bs.values_of(r)
        if all(not values.get(y) for y in fy):
            continue
        base = _side_rules(item(r.label, values, "BS", r.heading, r.page), "BS", r.label, assets_side)
        add_with_breakdown(bs, r, "BS", r.heading, base)

    # ---- Other equity: opening balance and movements from the statement of changes in equity
    other_equity = next((r for r in bs.rows if r.values and re.match(r"^other equity$|^reserves and surplus$", r.label, re.I)), None)
    pat_row = next((r for r in pnl.rows if r.values and re.match(r"^total profit \(loss\) for period$", r.label, re.I)), None)
    pat = pnl.values_of(pat_row) if pat_row is not None else {}
    opening, moves = _soce(sections, fy, pat)
    earliest = min(fy, key=lambda f: int(f[2:]))
    oe = bs.values_of(other_equity) if other_equity is not None else {}
    if opening is None and oe.get(earliest) is not None and pat.get(earliest) is not None:
        opening = oe[earliest] - pat[earliest] - (moves.get(earliest) or 0.0)
        result.log.append("WARNING: the statement of changes in equity could not be read; the opening Other equity is "
                          "worked back from the balance sheet and the profit.")
    if opening is not None:
        it = item("Other equity - opening balance", {earliest: opening}, "BS", "Other equity", None, source="pdf-note")
        it.statement_tag, it.category, it.is_new, it.confidence = "Opening Reserves", "", False, "Matched"
        it.flags = [f for f in it.flags if f != "unplaced"]
        result.items.append(it)
    if any(abs(v) > 1e-9 for v in moves.values()):
        it = item("Other changes in equity (statement of changes in equity)", moves, "BS", "Other equity", None,
                  source="pdf-note")
        it.statement_tag, it.category, it.is_new, it.confidence = "Reserves Adjustments", "", False, "Matched"
        it.flags = [f for f in it.flags if f != "unplaced"]
        result.items.append(it)
    years = sorted(fy, key=lambda f: int(f[2:]))
    running = opening
    for y in years:
        if running is None or pat.get(y) is None or oe.get(y) is None:
            break
        running = running + pat[y] + (moves.get(y) or 0.0)
        if abs(running - oe[y]) > TOL:
            result.log.append(f"WARNING: Other equity for {y} rolls forward to {running:,.2f} but the balance sheet shows "
                              f"{oe[y]:,.2f}.")

    # ---- Statement of profit and loss
    below_total = False                 # past 'Total expenses': prior period / exceptional items, minority interest
    for r in pnl.rows:
        if not r.values:
            continue
        low = r.label.lower()
        if re.match(r"^total expenses", low):
            below_total = True
            continue
        if re.match(r"^total profit before tax", low):
            below_total = False          # tax lines follow
            continue
        if re.match(r"^total profit \(loss\) for period$", low):
            break
        if re.search(r"minority interest|non.?controlling", low) and not re.search(r"before", low):
            # consolidated (Indian GAAP): the minority's share comes off the group's profit; the owners' profit
            # is what the reserves take, so the share is a line of the P&L (as the difference the PDF prints)
            before = next((x for x in pnl.rows if x.values and re.search(r"before minority interest", x.label, re.I)), None)
            if before is not None and pat:
                b = pnl.values_of(before)
                share = {y: pat[y] - b[y] for y in fy if pat.get(y) is not None and b.get(y) is not None}
                it = item("Share of loss / (profit) of minority interest", share, "P&L", r.heading, r.page)
                it.statement_tag, it.category, it.confidence = "", "Other Income", "Matched"
                it.flags = [f for f in it.flags if f != "unplaced"] + ["minority_interest", "below_total_expenses"]
                result.items.append(it)
            continue
        if r.is_total or (_SUMMARY_LINE.search(r.label) and not _PNL_ITEM_BELOW.search(r.label))                 or "comprehensive income" in low:
            continue
        values = pnl.values_of(r)
        if all(not values.get(y) for y in fy):
            continue
        base = item(r.label, values, "P&L", r.heading, r.page, ["below_total_expenses"] if below_total else [])
        face_values = values
        group = next((rows for pattern, rows in groups.items() if re.search(pattern, r.label, re.I)), None)
        if group:
            sums = {y: sum(notes.values_of(g).get(y) or 0.0 for g in group) for y in fy}
            if all(abs(sums[y] - (face_values.get(y) or 0.0)) <= TOL + 0.005 * len(group) for y in fy):
                face_key = schedule_for(base)
                for g in group:
                    if all(not notes.values_of(g).get(y) for y in fy):
                        continue
                    parts = _breakdown(notes, g, fy)
                    for name, vals in (parts.items() if parts else [(g.label, notes.values_of(g))]):
                        it = item(name, vals, "P&L", g.label if parts else g.heading, g.page,
                                  ["from_footnote"] if parts else [], "pdf-note")
                        own = schedule_for(it)
                        if face_key and own != face_key and not any({own, face_key} <= s for s in _SAME_IS_LINE):
                            it.schedule_override = face_key
                            it.flags = [f for f in it.flags if f != "unplaced"]
                        result.items.append(it)
                result.log.append(f"{r.label}: itemised from the notes on income and expenses [500100]")
                continue
            result.log.append(f"{r.label}: the [500100] breakup does not add up to the statement, so the statement figure is used")
        add_with_breakdown(pnl, r, "P&L", r.heading, base)
    return result
