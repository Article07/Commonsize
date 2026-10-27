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
from core.routing import CHANGE_IN_INVENTORIES, schedule_for

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
_BS_NOTE_CODES, _SHARE_CODES = ("401100", "200600"), ("400100", "200100")
_PNL_MUST = r"^revenue from|^total income|^total revenue"
# in the XBRL statements these are the purchases (Sch. Cost of goods sold)
_PURCHASES = re.compile(r"^\W*(?:cost of (?:raw )?materials? consumed|purchases? of stock.in.trade)", re.I)


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
    path: tuple = ()                  # every heading it sits under, outermost first


@dataclass
class Table:
    columns: list                     # [("I" | "D", "FY24")] instant / duration, left to right
    rows: list = field(default_factory=list)
    footnotes: dict = field(default_factory=dict)   # letter -> [(name, value)]
    open: list = field(default_factory=list)        # headings not yet closed by their total, innermost last

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
            if pending is not None and (not heading or heading[:1].islower()):
                heading = f"{pending} {heading}".strip()     # a heading wrapped before "[Abstract]"
            table.rows.append(Row(heading, [], heading, page, path=tuple(table.open)))   # a heading marker
            table.open.append(heading)
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
        heading = table.open[-1] if table.open else ""
        is_total = bool(re.match(r"^total\b", label, re.I))
        table.rows.append(Row(label, values, heading, page, is_total, tuple(table.open)))
        if is_total and heading and _closes(label, heading):
            # the group's total: lines after it (a "CSR expenditure" below depreciation) are not part of it
            table.open.pop()
    close_footnotes()
    return sections


def _closes(total, heading):
    """'Total depreciation, depletion and amortisation expense' closes 'Depreciation, depletion and amortisation expense'."""
    def words(text):
        return re.findall(r"[a-z]+", re.sub(r"^(?:total|disclosure of|breakup of)\s+", "", text.lower()))
    h, t = words(heading), words(total)
    n = min(2, len(h))
    return n > 0 and t[:n] == h[:n]


def _printed(pages, value, mult):
    """True when the amount appears as a figure anywhere in the document (any digit grouping)."""
    for target in {round(value), round(value * mult)}:
        digits = str(abs(target))
        if len(digits) < 4:
            continue
        for p in pages:
            for line in p.lines:
                if re.search(rf"(?<![\d.]){digits}(?![\d])", line.replace(",", "")):
                    return True
    return False


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
            part = out.setdefault(name, {})
            part[fy] = part.get(fy, 0.0) + x       # a name can be listed twice ("Obsolete Stock : 15842 Obsolete Stock : 0")
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
    """
    {face pattern: [Row]} -- the item rows of each [500100] group. Headings nest ("Managerial remuneration" >
    "Remuneration to directors"); a sub-heading's total counts as an item only when nothing under it was itemised
    ("Total remuneration to manager" alone), never on top of its own items ("Total contract cost" after "Site
    labour" and "Other claims"). A line repeated with the same figures (a total printed without "Total") counts once.
    """
    out = {}
    if table is None:
        return out
    rows = table.rows
    for head, face, total in _PNL_GROUPS:
        start = next((k for k, r in enumerate(rows) if not r.values and re.search(head, r.label, re.I)), None)
        if start is None:
            continue
        items, stack, previous = [], [[rows[start].label, False]], None      # stack: [heading, itemised under it]
        for r in rows[start + 1:]:
            if r.values and re.search(total, r.label, re.I):
                break
            if not r.values:
                if r.label:
                    stack.append([r.label, False])
                continue
            if previous is not None and r.label.lower() == previous.label.lower() and r.values == previous.values:
                continue
            previous = r
            if r.is_total:
                k = next((k for k in range(len(stack) - 1, -1, -1) if _closes(r.label, stack[k][0])), None)
                itemised = stack[k][1] if k is not None else any(h[1] for h in stack)
                if k is not None:
                    del stack[k:]
                if itemised:
                    continue
                items.append(r)                     # a sub-heading with only its total
            else:
                items.append(r)
            for h in stack:
                h[1] = True
        out[face] = items
    return out


def _section_pages(pages, codes):
    """1-based page numbers on which any of the given sections is open."""
    out, current = set(), None
    for p in pages:
        if current in codes:
            out.add(p.page)
        for line in p.lines:
            m = re.match(r"^\[(\d{6})\]", line.strip())
            if m:
                current = m.group(1)
                if current in codes:
                    out.add(p.page)
    return out


def read(pages, pdf_bytes=None):
    """-> (ExtractionResult, MappingResult) for an XBRL rendering. With the PDF itself, the note tables are also
    read by position (Balance Sheet breakups, components of reserves)."""
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
    extraction.xbrl_tables = []
    if pdf_bytes:
        from core.xbrl_tables import read_tables
        codes = set(_BS_NOTE_CODES + _EQUITY_CODES)
        try:
            extraction.xbrl_tables = read_tables(pdf_bytes, _section_pages(pages, codes), codes)
        except Exception as exc:          # the statements are still read; only the extra detail is lost
            extraction.warnings.append(f"The note tables could not be read by position ({exc}); Balance Sheet lines "
                                       "are not itemised.")
    return extraction, map_xbrl(extraction, mult)


def _soce(sections, fy_columns, pat):
    """
    (opening Other equity of the earliest year | None, {fy: movements other than profit} from the closing
    balances, {fy: the same} from the "Total changes" line). The two differ when the filing's total column
    misstates a closing balance or leaves a reserve out of its changes line; the caller keeps the one that
    rolls forward to the balance sheet.
    """
    table = _main_table(sections, _EQUITY_CODES, r"other equity at end of period|reserves at end of period")
    if table is None:
        return None, {}, {}
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
    moves, stated = {}, {}
    if change is not None:
        d_change = table.values_of(change, kinds=("D",))
        d_profit = table.values_of(profit, kinds=("D",)) if profit is not None else pat
        ends = table.values_of(end) if end is not None else {}
        for fy in fy_columns:
            if d_change.get(fy) is not None:
                stated[fy] = d_change[fy] - (d_profit.get(fy) or 0.0)
            prev = f"FY{int(fy[2:]) - 1:02d}"
            if ends.get(fy) is not None and ends.get(prev) is not None:
                # closing less opening: also catches a reserve whose balance is printed without a movement line
                # (a foreign currency translation reserve that is not in "Total changes in reserves")
                moves[fy] = ends[fy] - ends[prev] - (d_profit.get(fy) or 0.0)
            elif d_change.get(fy) is not None:
                moves[fy] = d_change[fy] - (d_profit.get(fy) or 0.0)
    return opening, moves, stated


def _side(label):
    if re.search(r"non.?current|long.?term", label, re.I):
        return "non-current"
    if re.search(r"\bcurrent\b|short.?term", label, re.I):
        return "current"
    return None


def _key(label):
    """'Other current financial liabilities, others' / 'Total trade receivables' -> the line it breaks up."""
    t = re.sub(r"\[.*?\]", " ", label.lower())
    t = re.sub(r"non.?current|long.?term|short.?term|\bcurrent\b", " ", t)
    t = re.sub(r"\b(?:total|gross|others?|notes?|disclosure of|breakup of|details of|subclassification of|"
               r"classification of)\b", " ", t)
    return " ".join(re.findall(r"[a-z]+", t))


def _clean_name(text):
    """A cell's text without footnote markers and 'Textual information (12) [See below]' placeholders."""
    t = re.sub(r"textual information\s*(?:\(\d+\))?\s*(?:\[see below\]|below\])?", " ", text, flags=re.I)
    t = re.sub(r"\((?:[A-Z]{1,2})\)|\[see below\]|below\]", " ", t, flags=re.I)
    t = re.sub(r"\s+", " ", t).strip(" ,;")
    return t if re.search(r"[A-Za-z]{3}", t) else ""


def _member_names(table):
    """Display name of each member: the filing's description of it ("Description ..." / "Nature of ..." rows,
    read down its first column), else the member heading."""
    out = []
    # cell text names a member only where the filing describes its members (a "Description ..." / "Nature of ..."
    # row); other text cells are explanatory paragraphs
    rows = table.rows
    near = set()                      # the describing row, the figures row just above it, its wrapped lines below
    for k, r in enumerate(rows):
        if re.match(r"description|nature of", r.label, re.I):
            near.update({k - 1, k} if k and rows[k - 1].values else {k})
            j = k + 1
            while j < len(rows) and not rows[j].values:
                near.add(j)
                j += 1
    for m in table.members:
        col = m.columns[0]
        texts = [_clean_name(r.texts[col]) for k, r in enumerate(rows) if k in near and col in r.texts
                 and not re.match(r"mode of valuation|type of share|country of", r.label, re.I)]
        texts = [t for t in texts if t]
        name = " ".join(texts) if texts else _clean_name(m.name)
        name = re.sub(r"\s+", " ", name.replace("[Member]", "")).strip(" ,")
        if not texts:
            # an outer current / non-current axis stacked above the member's own heading
            name = re.sub(r"^(?:non.?current|current)\s+(?=\S)", "", name, flags=re.I)
        out.append(name)
    return out


def _note_candidates(tables, fy_columns):
    """
    [(key, side, {name: {fy: value}})] -- every way the note tables break up a Balance Sheet line. Only tables
    with members (an "[Axis]"), not the explanatory text blocks; a breakup into a single part only where the
    filing names that part itself (a "Description ..." / "Nature of ..." row).
    """
    sides = {"current": "current", "noncurrent": "non-current"}
    sets = {}                         # (key, side, row) -> [parts, the parts are named by the filing]
    for t in tables:
        if t.code not in _BS_NOTE_CODES or not t.axis.strip() or not t.members:
            continue
        names = _member_names(t)
        described = any(re.match(r"description|nature of", r.label, re.I) for r in t.rows)
        member_sides = [sides.get(re.sub(r"[^a-z]", "", m.name.lower())) for m in t.members]
        if all(member_sides):
            # the members are the current / non-current split: the rows are the items (provisions)
            for m, side in zip(t.members, member_sides):
                entry = sets.setdefault((_key(t.title), side, "rows"), [{}, True])
                for r in t.rows:
                    if r.values and not r.is_total:
                        for fy, v in t.values(r, m).items():
                            entry[0].setdefault(r.label, {})[fy] = v
            continue
        seen = set()
        for r in t.rows:
            if not r.values or r.label.lower() in seen or re.match(r"description|nature of", r.label, re.I):
                continue
            seen.add(r.label.lower())
            entry = sets.setdefault((_key(r.label), _side(r.label), re.sub(r"[^a-z]", "", r.label.lower())), [{}, described])
            for m, name in zip(t.members, names):
                if name:
                    for fy, v in t.values(r, m).items():
                        entry[0].setdefault(name, {})[fy] = v
    out = []
    for (key, side, _), (parts, described) in sets.items():
        parts = {n: {fy: v for fy, v in vals.items() if fy in fy_columns} for n, vals in parts.items()}
        parts = {n: v for n, v in parts.items() if any(v.values())}
        # a member that is the sum of the others is their total ("Company inventories", the default member)
        for name in list(parts):
            others = [v for n, v in parts.items() if n != name]
            if others and all(abs(parts[name].get(fy, 0.0) - sum(o.get(fy, 0.0) for o in others)) <= TOL
                              for fy in fy_columns):
                del parts[name]
                break
        if len(parts) > 1 or (parts and described):
            out.append((key, side, parts))
    return out


def _note_parts(candidates, label, values, fy_columns):
    """{name: {fy: value}} breaking up a Balance Sheet line, from the first candidate that adds up to it."""
    key, side = _key(label), _side(label)
    if not key or not any(values.get(fy) for fy in fy_columns):
        return None
    for k, s, parts in candidates:
        if k != key or (s is not None and side is not None and s != side):
            continue
        if all(abs(sum(p.get(fy, 0.0) for p in parts.values()) - (values.get(fy) or 0.0)) <= TOL + 0.005 * len(parts)
               for fy in fy_columns):
            if len(parts) == 1 and _key(next(iter(parts))) == key:
                return None               # only the line itself again
            return parts
    return None


# headings that only name a part of the statement; a line directly under one of them gets a row of the schedule
# its section belongs to, a line under any other (more particular) heading gets that heading as a new heading
_SECTION_HEADING = re.compile(
    r"^(?:balance sheet|assets|equity and liabilities|equity|liabilities|shareholders'? funds|non.?current assets|"
    r"current assets|non.?current liabilities|current liabilities|equity attributable to .*|statement of profit and "
    r"loss|income|expenses|tax expense|total .*|.*\[line ?items\].*|.*line items)$", re.I)


def _bs_target(path, assets_side):
    """Schedule for an unrecognised Balance Sheet line, from the section the filing puts it in."""
    for h in reversed(path):
        h = h.lower()
        if re.search(r"non.?current assets", h):
            return "Long-term Loans & Advances"
        if re.search(r"^current assets|current financial assets", h):
            return "Current Asset - Other"
        if re.search(r"liabilities", h) and not re.search(r"equity and liabilities", h):
            return "Current Liability - Other"
        if re.search(r"^equity$|shareholders|equity attributable", h):
            return "Other Reserves"
    return "Current Asset - Other" if assets_side else "Current Liability - Other"


_ASSET_SCHEDULES = {"Fixed Assets", "Investments", "Current Investments", "Cash & Bank", "Trade Receivables", "Inventory",
                    "Long-term Loans & Advances", "Short-term Loans & Advances", "Current Asset - Other",
                    "Deferred Tax Assets"}


def _schedule_of_heading(heading, statement, section_schedule):
    """The Reference file's schedule for a filing heading ("Fixed assets" -> Fixed Assets), on the same side of the
    statement as the line's section; None when the Reference file has no such heading."""
    from core.models import LineItem
    from core.resolver import resolve
    res = resolve(heading, "", statement)
    if not res.resolved or res.via == "heading":
        return None
    key = schedule_for(LineItem(statement, heading, res.tag, {}, category=res.category))
    if key is None or key in ("Reserves and Surplus", "Reserves Adjustments", CHANGE_IN_INVENTORIES):
        return None
    if statement == "BS" and (key in _ASSET_SCHEDULES) != (section_schedule in _ASSET_SCHEDULES):
        return None                       # a heading word on the other side of the balance sheet
    return key


def _new_heading(path):
    """The filing's own heading for a line, when it is more particular than a part of the statement."""
    for h in reversed(path):
        if h and not _SECTION_HEADING.match(h.strip()):
            return h.strip()
    return None


def _equity_components(tables):
    """(securities premium {fy: balance} | None, other reserves (name, {fy: balance}) | None): each component's
    balance at the end of each year, from the changes in equity / reserves tables."""
    balances = {}
    for t in tables:
        if t.code not in _EQUITY_CODES:
            continue
        end = next((r for r in t.rows if r.values and re.search(r"(?:other equity|reserves|equity) at end of period",
                                                                   r.label, re.I)), None)
        if end is None:
            continue
        for m in t.members:
            for fy, v in t.values(end, m).items():
                balances.setdefault(re.sub(r"\s+", " ", m.name).strip(), {})[fy] = v
    premium = next((v for n, v in balances.items() if re.search(r"securities premium", n, re.I)), None)
    others = {n: v for n, v in balances.items()
              if not re.search(r"securities premium|^equity|attributable|^reserves$|retained|surplus|comprehensive|"
                               r"share capital|controlling|^total|other equity|^$|profit and loss", n, re.I)}
    # a component that is the total of the others is their heading ("Other reserves" over "General reserve")
    for name in list(others):
        rest = [v for n, v in others.items() if n != name]
        years = set(others[name])
        if rest and years and all(abs(others[name][y] - sum(v.get(y, 0.0) for v in rest)) <= TOL for y in years):
            del others[name]
            break
    other = None
    if others:
        years = {fy for v in others.values() for fy in v}
        other = (", ".join(others), {fy: sum(v.get(fy, 0.0) for v in others.values()) for fy in years})
    return premium, other


def _recover_left_out(result, item, fy_columns, pnl, pat, left_out, mult):
    """
    Safety net for a P&L line this reader does not know yet. When the profit read falls short of the profit the
    filing prints, and one of the lines it left out as a summary (a "share of profit of ...", an unusual
    exceptional item) accounts for the difference exactly in every year, that line is brought in -- as printed,
    in its place (before or after tax) -- and the log names it.
    """
    from core.statements import compute_statements
    years = sorted(fy_columns, key=lambda f: int(f[2:]))
    if not left_out or not pat or any(pat.get(y) is None for y in years):
        return
    tol = max(1.0, 0.011 * mult)
    statements = compute_statements(result.items, years)[0]
    short = {y: pat[y] * mult - statements[y]["pat"] for y in years}
    if all(abs(d) <= tol for d in short.values()):
        return
    for r, below, after in left_out:
        values = pnl.values_of(r)
        for sign in (1, -1):
            if all(abs(sign * (values.get(y) or 0.0) * mult - short[y]) <= tol for y in years):
                it = item(r.label, {y: sign * (values.get(y) or 0.0) for y in years}, "P&L", r.heading, r.page)
                it.face_line, it.face_sign = r.label, sign
                it.statement_tag, it.category, it.confidence = "", "Other Income", "Matched"
                it.flags = [f for f in it.flags if f != "unplaced"]
                if below or after:
                    it.flags.append("below_total_expenses")
                if after:
                    it.flags.append("minority_interest")       # after tax: kept out of the profit-before-tax check
                result.items.append(it)
                result.log.append(f"WARNING: '{r.label}' was not a line this reader knows, but the profit the filing "
                                  "prints includes it exactly, so it is included (as other income"
                                  f"{', after tax' if after else ''}). Check its placement.")
                return


def _anchor_reserves(result, item, fy_columns, other_equity, pat, mult):
    """
    Reserves end every year at the Other equity the balance sheet prints. The statement of changes in equity
    is used for the opening balance and the named movements, but filings vary (a misstated total column, a
    restated opening, the non-controlling interest inside or outside it, components counted twice): whatever
    it leaves unexplained becomes one labelled movement of the year, worked out with the profit the P&L
    prints -- so a P&L that is read wrongly still shows as a P&L difference, it is never absorbed here.
    """
    from core.statements import compute_statements
    years = sorted(fy_columns, key=lambda f: int(f[2:]))
    if not other_equity or not pat or any(other_equity.get(y) is None or pat.get(y) is None for y in years):
        return
    tol = max(1.0, 0.011 * mult)
    added = {}
    for _ in years:
        statements = compute_statements(result.items, years)[0]
        p_and_l_error = 0.0               # what the reserves would be off by if only the P&L were read wrongly
        fix = None
        for y in years:
            p_and_l_error += pat[y] * mult - statements[y]["pat"]
            gap = other_equity[y] * mult - (statements[y]["reserves"] + p_and_l_error)
            if abs(gap) > tol:
                fix = (y, gap)
                break
        if fix is None:
            break
        y, gap = fix
        added[y] = added.get(y, 0.0) + gap
        it = item("Other movements in reserves (to the balance sheet's Other equity)", {y: gap / mult}, "BS",
                  "Other equity", None, [], "pdf-note")
        it.face_line = next((i.face_line for i in result.items if i.statement_tag == "Opening Reserves"), "Other equity")
        it.statement_tag, it.category, it.is_new, it.confidence = "Reserves Adjustments", "", False, "Matched"
        it.flags = [f for f in it.flags if f != "unplaced"]
        result.items.append(it)
    for y, gap in added.items():
        result.log.append(f"WARNING: Other equity {y}: the statement of changes in equity does not explain "
                          f"{gap / mult:,.2f} of the year's change in reserves; it is shown as 'Other movements in "
                          "reserves' so that reserves equal the balance sheet's Other equity.")


def _group_checks(table, fy_columns):
    """
    Every printed total that closes a heading ("Total current assets", "Total expenses", "Total tax expense") is
    the sum of what sits under that heading. -> (rows whose sign must turn so their group adds up -- only where
    turning one row makes every year add up --, notes on groups that do not add up at all, naming them).
    """
    rows = [r for r in table.rows if r.values]
    flips, notes = [], []
    for t in rows:
        if not t.is_total or not t.path or not _closes(t.label, t.path[-1]) or re.search(r"comprehensive", t.label, re.I):
            continue                      # (total comprehensive income is the profit plus OCI, not its group's sum)
        depth = len(t.path)
        below = [r for r in rows if r is not t and len(r.path) >= depth and r.path[:depth] == t.path]
        parts = [r for r in below if len(r.path) == depth and not (r.is_total and _closes(r.label, t.path[-1]))]
        for child in dict.fromkeys(r.path[depth] for r in below if len(r.path) > depth):
            inner = [r for r in below if len(r.path) > depth and r.path[depth] == child]
            sub_total = next((r for r in inner if r.is_total and len(r.path) == depth + 1 and _closes(r.label, child)), None)
            parts += [sub_total] if sub_total else [r for r in inner if len(r.path) == depth + 1 and not r.is_total]
        if not parts:
            continue
        total = table.values_of(t)

        def adds_up(flip=None):
            for y in fy_columns:
                if total.get(y) is None:
                    continue
                got = sum((-1 if r is flip else 1) * (table.values_of(r).get(y) or 0.0) for r in parts)
                if abs(got - total[y]) > TOL + 0.005 * len(parts):
                    return False
            return True

        if adds_up():
            continue
        fix = [r for r in parts if not r.is_total and any(table.values_of(r).get(y) for y in fy_columns) and adds_up(r)]
        if len(fix) == 1:
            flips.append(fix[0])
            notes.append(f"{fix[0].label}: printed with the sign the filing's own '{t.label}' does not use; that "
                         "total's sign is used.")
        else:
            y = next(y for y in fy_columns if total.get(y) is not None)
            got = sum(table.values_of(r).get(y) or 0.0 for r in parts)
            notes.append(f"WARNING: the lines under '{t.path[-1]}' add up to {got:,.2f} in {y}, not to the "
                         f"{total[y]:,.2f} the filing prints as '{t.label}'.")
    return flips, notes


def _deductions_to_flip(pnl, fy_columns):
    """
    Revenue lines printed with the sign their total does not use. The MCA form computes revenue as gross less
    duties ("Other duties taxes collected", excise duty): a duty entered as a negative figure is added by the
    filing's own total. -> the deduction rows whose sign has to be turned so the lines add up to the total
    the filing prints (none when they already add up).
    """
    rows = [r for r in pnl.rows if r.values]
    end = next((k for k, r in enumerate(rows) if re.match(r"^total revenue from operations", r.label, re.I)), None)
    if end is None:
        return []
    lines = [r for r in rows[:end] if not r.is_total]
    total = pnl.values_of(rows[end])
    duties = [r for r in lines if re.search(r"dut(?:y|ies)|taxes collected|excise|deductions? (?:from|on) revenue", r.label, re.I)]
    if not duties:
        return []

    def adds_up(flip):
        for y in fy_columns:
            if total.get(y) is None:
                continue
            got = sum((-1 if any(r is f for f in flip) else 1) * (pnl.values_of(r).get(y) or 0.0) for r in lines)
            if abs(got - total[y]) > TOL + 0.005 * len(lines):
                return False
        return True

    return [] if adds_up([]) or not adds_up(duties) else duties


def _other_comprehensive_income(sections):
    """{fy: other comprehensive income} from the statement of changes in equity (its first, total, column)."""
    table = _main_table(sections, _EQUITY_CODES, r"other equity at end of period|reserves at end of period")
    if table is None:
        return {}
    row = next((r for r in table.rows if r.values and re.search(r"^changes in comprehensive income components|"
                                                                  r"^other comprehensive income", r.label, re.I)), None)
    return table.values_of(row, kinds=("D",)) if row is not None else {}


def _share_capital_labels(lines):
    """('9,46,91,773 equity shares of Rs. 5 each', '15,70,00,000 equity shares of Rs. 5 each') from the share
    capital note, else (None, None)."""
    text, on = [], False
    for _, line in lines:
        m = _SECTION.match(line)
        if m:
            on = m.group(1) in _SHARE_CODES
            continue
        if on:
            text.append(line)
    joined = " ".join(text)
    issued = re.search(r"Number of shares issued\s*\[shares\]\s*([\d,]+)", joined)
    authorised = re.search(r"Number of shares authorised\s*\[shares\]\s*([\d,]+)", joined)
    par = re.search(r"(?:Par|Face|Nominal) value per share\s*\[INR/shares\]\s*([\d.]+)", joined)
    if not issued or not par:
        return None, None, None
    rupees = f"{float(par.group(1)):g}"
    issued_text = f"{issued.group(1)} equity shares of Rs. {rupees} each"
    authorised_text = f"{authorised.group(1)} equity shares of Rs. {rupees} each" if authorised else None
    return issued_text, authorised_text, float(issued.group(1).replace(",", "")) * float(par.group(1))


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

    def place_by_position(it, path, schedule, how):
        """
        A line the taxonomy and aliases do not know: placed where the filing classifies it, never dropped. Under a
        heading the Reference file has ("Fixed assets") it goes into that schedule; under one it does not have
        ("Current financial liabilities") into a new heading of that name, within the schedule its section belongs to.
        """
        if schedule_for(it) is not None:
            return it
        it.flags = [f for f in it.flags if f != "unplaced"] + ["placed_by_position"]
        it.is_new = True
        heading = _new_heading(path) if schedule not in ("Other Reserves",) else None
        known = _schedule_of_heading(heading, it.statement, schedule) if heading else None
        if known:
            schedule, heading = known, None
            how = f"under the filing's heading {_new_heading(path)!r}, which the Reference file has"
        it.schedule_override = schedule
        if heading:
            it.flags.append("new_heading")
            it.pdf_heading = heading
        result.log.append(f"WARNING: '{it.line_item}' is not a line the app knows; placed {how}"
                          f"{f' under a new heading {heading!r}' if heading else ''} (highlighted yellow) -- check it.")
        return it

    def item(label, values, statement, heading, page, flags=(), source="pdf-face"):
        return _make_item(label, {k: v for k, v in values.items() if k in fy}, statement, heading, source, page,
                          list(flags) + ["own_row"], fy, mult)

    nci = None
    equity_extra = []                     # equity lines besides share capital and other equity (placed by position)
    tables = getattr(extraction, "xbrl_tables", None) or []
    note_sets = _note_candidates(tables, fy)
    issued_text, authorised_text, issued_value = _share_capital_labels(_content_lines(extraction.pages))
    inventory_parts = None

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
    bs_flips, bs_notes = _group_checks(bs, fy)
    result.log += bs_notes
    assets_side = None
    starts = []                           # (first item index, face line, sign): lines created for that face line
    for r in bs.rows:
        if r.values:
            starts.append((len(result.items), r.label, -1 if any(r is f for f in bs_flips) else 1))
        if not r.values:
            if re.match(r"^assets$", r.label, re.I):
                assets_side = True
            elif re.match(r"^equity and liabilities$", r.label, re.I):
                assets_side = False
            continue
        if r.is_total or re.search(r"^other equity$|^reserves and surplus$|equity attributable", r.label, re.I):
            continue
        values = bs.values_of(r)
        if any(r is f for f in bs_flips):
            values = {y: -v for y, v in values.items()}
        if all(not values.get(y) for y in fy):
            continue
        base = _side_rules(item(r.label, values, "BS", r.heading, r.page), "BS", r.label, assets_side)
        target = _bs_target(r.path, assets_side)
        base = place_by_position(base, r.path, target, f"in {target} (as the filing classifies it)")
        if target == "Other Reserves" and "placed_by_position" in base.flags:
            base.statement_tag, base.schedule_override = "Other Reserves", ""
            equity_extra.append(base)
        if re.match(r"^non.?controlling interests?$", r.label, re.I):
            nci = base                    # decided at the end: some filings already carry it in their reserves
            nci.face_line = r.label
            continue
        latest = max(fy, key=lambda f: int(f[2:]))
        if re.search(r"share capital", r.label, re.I) and issued_text and base.fy_values.get(latest) \
                and abs(issued_value - base.fy_values[latest]) <= 0.01 * base.fy_values[latest]:
            # the wording only when shares x face value is the share capital shown (one class of shares)
            base.line_item = issued_text
            base.flags.append("label_row")
            base.pdf_heading = f"Authorised: {authorised_text}" if authorised_text else base.pdf_heading
        parts = None if re.search(r"share capital", r.label, re.I) else _note_parts(note_sets, r.label, values, fy)
        if parts:
            face_key = schedule_for(base)
            for name, vals in parts.items():
                it = item(name, {y: vals.get(y, 0.0) for y in fy if values.get(y) is not None}, "BS", r.label,
                          r.page, ["from_note"], "pdf-note")
                it.line_item = name                      # as the note names it ("Non-current lease liabilities")
                if face_key:
                    it.schedule_override = face_key      # a part stays under the line it breaks up
                    it.flags = [f for f in it.flags if f != "unplaced"]
                result.items.append(it)
            if re.match(r"^inventor", r.label, re.I):
                inventory_parts = parts
            result.log.append(f"{r.label}: itemised from the notes on liabilities and assets ({len(parts)} lines)")
            continue
        add_with_breakdown(bs, r, "BS", r.heading, base)

    reserves_from = len(result.items)
    # ---- Other equity: opening balance and movements from the statement of changes in equity
    other_equity = next((r for r in bs.rows if r.values and re.match(r"^other equity$|^reserves and surplus$", r.label, re.I)), None)
    pat_row = next((r for r in pnl.rows if r.values and re.match(r"^total profit \(loss\) for period$", r.label, re.I)), None)
    pat = pnl.values_of(pat_row) if pat_row is not None else {}
    opening, moves, stated_moves = _soce(sections, fy, pat)
    earliest = min(fy, key=lambda f: int(f[2:]))
    oe = bs.values_of(other_equity) if other_equity is not None else {}
    running = opening
    for y in sorted(fy, key=lambda f: int(f[2:])):
        if running is None or pat.get(y) is None or oe.get(y) is None:
            break
        if (abs(running + pat[y] + (moves.get(y) or 0.0) - oe[y]) > TOL and y in stated_moves
                and abs(running + pat[y] + stated_moves[y] - oe[y]) <= TOL):
            # the total column's closing balance is misstated; its "Total changes" line and the balance sheet agree
            result.log.append(f"Other equity {y}: the reserves note's closing balance does not add up; its 'Total "
                              f"changes' line ({stated_moves[y] + pat[y]:,.2f}) does and is used.")
            moves[y] = stated_moves[y]
        running = oe[y]
    if opening is None and oe.get(earliest) is not None and pat.get(earliest) is not None:
        opening = oe[earliest] - pat[earliest] - (moves.get(earliest) or 0.0)
        result.log.append("WARNING: the statement of changes in equity could not be read; the opening Other equity is "
                          "worked back from the balance sheet and the profit.")
    elif opening is not None and oe.get(earliest) is not None and pat.get(earliest) is not None:
        implied = oe[earliest] - pat[earliest] - (moves.get(earliest) or 0.0)
        if abs(implied - opening) > TOL and _printed(extraction.pages, implied, mult):
            # the filing's opening does not roll forward to its balance sheet, while the reserves note's own text
            # prints the opening that does ("As per last Balance Sheet ..."): last year's closing was restated
            result.log.append(f"WARNING: the opening Other equity in the filing's tables ({opening:,.2f}) does not roll "
                              f"forward to the balance sheet's {earliest} figure; the opening its reserves note prints, "
                              f"{implied:,.2f}, is used.")
            opening = implied
    restated = {}
    years = sorted(fy, key=lambda f: int(f[2:]))
    running = opening
    for y in years:
        if running is None or pat.get(y) is None or oe.get(y) is None:
            break
        expected = running + pat[y] + (moves.get(y) or 0.0)
        if abs(expected - oe[y]) > TOL and y != earliest and _printed(extraction.pages, running + oe[y] - expected, mult):
            # the year opens on a restated balance the reserves note prints, not on last year's closing
            restated[y] = oe[y] - expected
            result.log.append(f"WARNING: {y} opens Other equity at {running + restated[y]:,.2f} as its reserves note "
                              f"prints, not at last year's closing {running:,.2f}; the difference is shown as a "
                              f"separate movement.")
        elif abs(expected - oe[y]) > TOL:
            result.log.append(f"WARNING: Other equity for {y} rolls forward to {expected:,.2f} but the balance sheet "
                              f"shows {oe[y]:,.2f}.")
        running = oe[y] if y in restated else expected
    premium, other = _equity_components(tables)
    if (premium and any(premium.get(y) is None for y in years)) or (other and any(other[1].get(y) is None for y in years)):
        premium = other = None            # a balance missing for a year: keep Other equity in one piece
    before = f"FY{int(earliest[2:]) - 1:02d}"

    def components(y):
        """securities premium + other reserves at the end of year y (the part of Other equity that is not surplus)"""
        if y not in years and not ((premium and y in premium) or (other and y in other[1])):
            y = earliest                  # no balance printed for the day before: taken as unchanged
        return (premium or {}).get(y, 0.0) + (other[1] if other else {}).get(y, 0.0)

    if premium or other:
        for y, prev in zip(years, [before] + years[:-1]):
            change = components(y) - components(prev)
            if abs(change) > TOL:
                moves[y] = (moves.get(y) or 0.0) - change    # the surplus moves by the rest
        if opening is not None:
            opening -= components(before)
        for label, tag, values in (("Securities premium", "Securities Premium", premium),
                                   (other[0] if other else "", "Other Reserves", other[1] if other else None)):
            if values and any(values.get(y) for y in years):
                it = item(label, {y: values[y] for y in years}, "BS", "Other equity", None, ["label_row"], "pdf-note")
                it.statement_tag, it.category, it.is_new, it.confidence = tag, "", False, "Matched"
                it.flags = [f for f in it.flags if f != "unplaced"]
                result.items.append(it)
        result.log.append("Other equity: securities premium and other reserves shown as balances of their own, from the "
                          "statement of changes in equity; the surplus rolls forward with the profit.")
    if opening is not None:
        # later years open on last year's balance sheet figure (the roll-forward uses the previous closing, so
        # these only matter when a year is taken on its own: a later year of a multi-file combination)
        openings = {earliest: opening}
        for prev, y in zip(years, years[1:]):
            if oe.get(prev) is not None:
                openings[y] = oe[prev] - components(prev)
        it = item("Other equity - opening balance", openings, "BS", "Other equity", None, source="pdf-note")
        it.statement_tag, it.category, it.is_new, it.confidence = "Opening Reserves", "", False, "Matched"
        it.flags = [f for f in it.flags if f != "unplaced"]
        result.items.append(it)
    oci = _other_comprehensive_income(sections)
    moves_label = "Other changes in equity (statement of changes in equity)"
    if oci and moves and all(abs((moves.get(y) or 0.0) - (oci.get(y) or 0.0)) <= TOL for y in years):
        moves_label = "Other comprehensive income (net of tax)"
    for label, values in ((moves_label, moves),
                          ("Opening reserves restated (differs from last year's closing)", restated)):
        if any(abs(v) > 1e-9 for v in values.values()):
            it = item(label, values, "BS", "Other equity", None, ["label_row"] if values is moves else [], "pdf-note")
            it.statement_tag, it.category, it.is_new, it.confidence = "Reserves Adjustments", "", False, "Matched"
            it.flags = [f for f in it.flags if f != "unplaced"]
            result.items.append(it)

    equity_label = other_equity.label if other_equity is not None else "Other equity"
    starts.append((reserves_from, equity_label, 1))
    # ---- Statement of profit and loss
    flipped = _deductions_to_flip(pnl, fy)
    for r in flipped:
        result.log.append(f"{r.label}: a deduction in the revenue (gross less duties); the filing's own total "
                          f"takes it with the opposite sign to the one printed, so that sign is used.")
    pnl_flips, pnl_notes = _group_checks(pnl, fy)
    for r, note in zip(pnl_flips, [n for n in pnl_notes if not n.startswith("WARNING")]):
        if not any(r is f for f in flipped):
            flipped.append(r)
            result.log.append(note)
    result.log += [n for n in pnl_notes if n.startswith("WARNING")]
    below_total = False                 # past 'Total expenses': prior period / exceptional items, minority interest
    after_tax = False                   # past 'Total profit before tax'
    after_income = False                # past 'Total income'
    left_out = []                       # (row, below total expenses, after tax): lines read as summaries
    for r in pnl.rows:
        if not r.values:
            continue
        starts.append((len(result.items), r.label, None))     # sign settled below, once 'flipped' is final
        low = r.label.lower()
        if re.match(r"^total (?:income|revenue)$", low):
            after_income = True
        if re.match(r"^total expenses", low):
            below_total = True
            continue
        if re.match(r"^total profit before tax", low):
            below_total, after_tax = False, True          # tax lines follow
            continue
        if re.match(r"^total profit \(loss\) for period$", low):
            break
        if re.search(r"minority interest|non.?controlling", low) and not re.search(r"before", low):
            # consolidated (Indian GAAP): after the group's profit come the minority's share and the share of
            # associates; the owners' profit is what the reserves take, so both are lines of the P&L (together the
            # difference the PDF prints between the two profits)
            before = next((x for x in pnl.rows if x.values and re.search(r"before minority interest", x.label, re.I)), None)
            if before is not None and pat:
                b = pnl.values_of(before)
                assoc_row = next((x for x in pnl.rows if x.values and re.search(r"share of profit.{0,12}associates?", x.label, re.I)), None)
                assoc = pnl.values_of(assoc_row) if assoc_row is not None else {}
                total = {y: pat[y] - b[y] for y in fy if pat.get(y) is not None and b.get(y) is not None}
                lines = (("Share of profit / (loss) of associates", {y: assoc.get(y) or 0.0 for y in total}),
                         ("Share of loss / (profit) of minority interest", {y: total[y] - (assoc.get(y) or 0.0) for y in total}))
                for name, share in lines:
                    if all(abs(v) < TOL for v in share.values()):
                        continue
                    it = item(name, share, "P&L", r.heading, r.page)
                    it.face_line = assoc_row.label if assoc_row is not None and name.startswith("Share of profit") else r.label
                    it.statement_tag, it.category, it.confidence = "", "Other Income", "Matched"
                    it.flags = [f for f in it.flags if f != "unplaced"] + ["minority_interest", "below_total_expenses"]
                    result.items.append(it)
            continue
        if re.match(r"^share of (?:profit|loss|\(?profit).{0,25}associates?", low) and not any(
                it.line_item.startswith("Share of profit / (loss) of associates") for it in result.items):
            # Ind AS consolidated: the share of associates / joint ventures comes after tax and is part of the
            # profit for the period (the Indian GAAP layout is handled with the minority interest above)
            values = pnl.values_of(r)
            if any(values.get(y) for y in fy):
                it = item("Share of profit / (loss) of associates and joint ventures", values, "P&L", r.heading, r.page)
                it.statement_tag, it.category, it.confidence = "", "Other Income", "Matched"
                it.flags = [f for f in it.flags if f != "unplaced"] + ["minority_interest", "below_total_expenses"]
                result.items.append(it)
            continue
        if re.match(r"^(?:exceptional|extraordinary) items", low):
            # added to the profit before them (a gain positive, a loss negative), like prior period items
            values = pnl.values_of(r)
            if any(values.get(y) for y in fy):
                it = item(r.label, values, "P&L", r.heading, r.page, ["below_total_expenses"])
                it.statement_tag, it.category, it.confidence = "", "Other Income", "Matched"
                it.flags = [f for f in it.flags if f != "unplaced"]
                result.items.append(it)
            continue
        if r.is_total or (_SUMMARY_LINE.search(r.label) and not _PNL_ITEM_BELOW.search(r.label))                 or "comprehensive income" in low:
            if not r.is_total and not re.search(r"per (?:equity )?share|attributable|comprehensive|before tax|"
                                                r"for period|\[inr/shares\]", r.label, re.I):
                left_out.append((r, below_total, after_tax))
            continue
        values = pnl.values_of(r)
        if any(r is f for f in flipped):
            values = {y: -v for y, v in values.items()}
        if all(not values.get(y) for y in fy):
            continue
        base = item(r.label, values, "P&L", r.heading, r.page, ["below_total_expenses"] if below_total else [])
        if schedule_for(base) is None:
            if after_tax:
                where = ("Tax Expense", "with tax") if re.search(r"tax", " ".join(r.path), re.I) else \
                    ("Other Income", "after tax")
            elif below_total:
                where = ("Other Income", "after total expenses")
            elif after_income:
                where = ("Administrative Expenses", "among the expenses")
            else:
                where = ("Revenue", "in revenue") if re.search(r"revenue from operations", " ".join(r.path), re.I) \
                    else ("Other Income", "in other income")
            base = place_by_position(base, r.path, where[0], where[1])
            if after_tax and where[0] == "Other Income":
                base.flags += ["below_total_expenses", "minority_interest"]
        if _PURCHASES.match(r.label):
            # in the XBRL statements these are the purchases: Sch. Cost of goods sold, "Cost of materials consumed"
            # below "Purchase of traded goods" (one line each, not itemised)
            base.schedule_override = "COGS-Purchase"
            base.flags = [f for f in base.flags if f != "unplaced"]
            result.items.append(base)
            continue
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
    if inventory_parts:
        goods = [v for n, v in inventory_parts.items() if re.search(r"finished|work.in.progress|stock.in.trade|traded", n, re.I)]
        change = next((it for it in result.items if schedule_for(it) == CHANGE_IN_INVENTORIES), None)
        if goods and change is not None:
            for y in fy:
                if change.fy_values.get(y) is not None and all(g.get(y) is not None for g in goods):
                    closing = sum(g[y] for g in goods) * mult
                    change.stock_levels[y] = (closing + change.fy_values[y], closing)
            if change.stock_levels:
                result.log.append("Change in inventories: opening and closing stock (finished goods, work-in-progress, "
                                  "stock-in-trade) from the inventories note.")
    # a "Cost of materials consumed" filed inside another line's breakup (Other expenses) is purchases too
    for it in result.items:
        if it.statement == "P&L" and _PURCHASES.match(it.line_item) and it.schedule_override != "COGS-Purchase":
            it.schedule_override = "COGS-Purchase"
            it.flags = [f for f in it.flags if f != "unplaced"]
            result.log.append(f"{it.line_item}: placed under purchases (Cost of goods sold)")
    for k, (first, face_label, sign) in enumerate(starts):
        nxt = starts[k + 1][0] if k + 1 < len(starts) else len(result.items)
        if sign is None:
            sign = -1 if any(f.label == face_label for f in flipped) else 1
        for it in result.items[first:nxt]:
            if not it.face_line:
                it.face_line, it.face_sign = face_label, sign
    _recover_left_out(result, item, fy, pnl, pat, left_out, mult)
    extra = {y: sum((e.fy_values.get(y) or 0.0) for e in equity_extra) / mult for y in fy}
    target_equity = {y: v + extra[y] for y, v in oe.items()} if oe else oe
    _anchor_reserves(result, item, fy, target_equity, pat, mult)
    if nci is not None and any(nci.fy_values.get(y) for y in fy):
        # Ind AS consolidated: the non-controlling interest is part of total equity. Where the statement of
        # changes in equity already rolls it into the reserves column (Krishi Rasayan) the balance sheet tallies
        # without it; where that column is the owners' alone (Indofil) it is short by exactly this amount.
        from core.statements import compute_statements
        years = sorted(fy, key=lambda f: int(f[2:]))
        gaps = compute_statements(result.items, years)[0]
        if all(abs(gaps[y]["gap"] + (nci.fy_values.get(y) or 0.0)) <= 0.5 * mult + 1 for y in years):
            result.items.append(nci)
            result.log.append("Non-controlling interest: shown under other liabilities (part of total equity, not in the "
                              "reserves column of the statement of changes in equity).")
        else:
            result.log.append("Non-controlling interest: already included in the reserves the statement of changes in "
                              "equity rolls forward, so not added again.")
    return result
