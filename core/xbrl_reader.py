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
            table.rows.append(Row(heading, [], heading, page))       # a heading marker (no values)
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
        table.rows.append(Row(label, values, heading, page, is_total))
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
        ends = table.values_of(end) if end is not None else {}
        for fy in fy_columns:
            prev = f"FY{int(fy[2:]) - 1:02d}"
            if ends.get(fy) is not None and ends.get(prev) is not None:
                # closing less opening: also catches a reserve whose balance is printed without a movement line
                # (a foreign currency translation reserve that is not in "Total changes in reserves")
                moves[fy] = ends[fy] - ends[prev] - (d_profit.get(fy) or 0.0)
            elif d_change.get(fy) is not None:
                moves[fy] = d_change[fy] - (d_profit.get(fy) or 0.0)
    return opening, moves


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
    other = None
    if others:
        years = {fy for v in others.values() for fy in v}
        other = (", ".join(others), {fy: sum(v.get(fy, 0.0) for v in others.values()) for fy in years})
    return premium, other


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

    def item(label, values, statement, heading, page, flags=(), source="pdf-face"):
        return _make_item(label, {k: v for k, v in values.items() if k in fy}, statement, heading, source, page,
                          list(flags) + ["own_row"], fy, mult)

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
            continue
        values = pnl.values_of(r)
        if all(not values.get(y) for y in fy):
            continue
        base = item(r.label, values, "P&L", r.heading, r.page, ["below_total_expenses"] if below_total else [])
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
    return result
