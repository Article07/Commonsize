"""
Reads the member tables of an MCA XBRL rendering by position (pdfplumber words), for the note sections whose
plain text cannot be tied to columns: a cell left empty shifts the figures after it ("Provision gratuity 12" --
non-current FY23 or current FY24?). Each figure is placed in the column whose date it sits under.

A table:
    Other current liabilities, others [Table] ..(1)
    Unless otherwise specified, all monetary values are in Millions of INR
    Other current liabilities, others [Axis]   Advance from customers    Statutory liabilities      <- members
                                               01/04/2023  01/04/2022    01/04/2023  01/04/2022     <- columns
                                               to          to            to          to
                                               31/03/2024  31/03/2023    31/03/2024  31/03/2023
    Other current liabilities, others                   3           4            17          17      <- a row
    Description of other current liabilities, others   Advance from customers ...               <- a text row

A member is a run of columns, one per year; it starts again when a year repeats. Its name is the header text
above its columns (the last "[Member]" phrase when several axes stack), or the row "Description ..." / "Nature
..." gives it per column.
"""

import re
from dataclasses import dataclass, field

_DATE = re.compile(r"^(\d{2})/(\d{2})/(\d{4})$")
_NUMBER = re.compile(r"^\(?-?[\d,]*\d(?:\.\d+)?\)?$")
_SECTION = re.compile(r"^\[(\d{6})\]")
_PAGE_HEADER = re.compile(r"(?:standalone|consolidated) financial statements for period", re.I)


@dataclass
class Column:
    fy: str
    kind: str                 # "D" period / "I" instant
    x1: float                 # right edge of its end date: figures are right-aligned under it
    center: float


@dataclass
class Member:
    name: str
    columns: list             # column indexes


@dataclass
class TRow:
    label: str
    values: dict              # column index -> value
    texts: dict               # column index -> text
    is_total: bool
    page: int


@dataclass
class MTable:
    code: str
    title: str
    axis: str
    columns: list
    members: list = field(default_factory=list)
    rows: list = field(default_factory=list)

    def member_of(self, col):
        return next(m for m in self.members if col in m.columns)

    def values(self, row, member):
        """{fy: value} of one member of a row."""
        return {self.columns[c].fy: v for c, v in row.values.items() if c in member.columns}

    def text(self, row, member):
        return " ".join(row.texts[c] for c in member.columns if c in row.texts).strip()


def _fy(month, year):
    return f"FY{(year if month <= 3 else year + 1) % 100:02d}"


def _number(text):
    t = text.replace(",", "")
    neg = t.startswith("(") and t.endswith(")")
    try:
        v = float(t.strip("()"))
    except ValueError:
        return None
    return -v if neg else v


def _lines(page):
    words = page.extract_words(x_tolerance=1.5, y_tolerance=2, keep_blank_chars=False)
    words.sort(key=lambda w: (w["top"], w["x0"]))
    lines, current, top = [], [], None
    for w in words:
        if top is not None and w["top"] - top > 2.5:
            lines.append(sorted(current, key=lambda w: w["x0"]))
            current = []
        if not current:
            top = w["top"]
        current.append(w)
    if current:
        lines.append(sorted(current, key=lambda w: w["x0"]))
    return lines


def _text(words):
    return " ".join(w["text"] for w in words)


def _columns(date_words):
    """Date words of a header -> [Column] left to right (a period stacks 'from', 'to', 'end')."""
    clusters = []
    for w in sorted(date_words, key=lambda w: (w["x0"] + w["x1"]) / 2):
        c = (w["x0"] + w["x1"]) / 2
        if clusters and abs(clusters[-1][0] - c) < 14:
            clusters[-1][1].append(w)
        else:
            clusters.append([c, [w]])
    out = []
    for c, ws in clusters:
        ws.sort(key=lambda w: w["top"])
        end = ws[-1]
        day, month, year = map(int, _DATE.match(end["text"]).groups())
        out.append(Column(_fy(month, year), "D" if len(ws) > 1 else "I", end["x1"], c))
    return out


def _members(columns, header_words):
    """Runs of columns named by the header text above them: within a member the years go back (FY24, FY23,
    FY22), so a new member starts when the year does not."""
    runs = []
    for i, col in enumerate(columns):
        if not runs or int(col.fy[2:]) >= int(columns[runs[-1][-1]].fy[2:]):
            runs.append([])
        runs[-1].append(i)
    spans = [(columns[r[0]].center, columns[r[-1]].center) for r in runs]
    bounds = []
    for k, (lo, hi) in enumerate(spans):
        left = (spans[k - 1][1] + lo) / 2 if k else lo - 80
        right = (hi + spans[k + 1][0]) / 2 if k + 1 < len(spans) else hi + 80
        bounds.append((left, right))
    out = []
    for run, (left, right) in zip(runs, bounds):
        words = [w for w in header_words if left <= (w["x0"] + w["x1"]) / 2 < right]
        words.sort(key=lambda w: (round(w["top"]), w["x0"]))
        text = _text(words)
        parts = [p.strip() for p in text.split("[Member]") if p.strip()]
        name = parts[-1] if "[Member]" in text and parts else text.strip()
        out.append(Member(name, run))
    return out


def read_tables(pdf_bytes, page_numbers, codes):
    """Member tables of the given sections ({'401100', ...}) on the given pages (1-based) -> [MTable]."""
    import io

    import pdfplumber

    tables, code, table = [], None, None
    state, header, dates, title, pending = None, [], [], "", None
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for number in sorted(page_numbers):
            if number < 1 or number > len(pdf.pages):
                continue
            for line in _lines(pdf.pages[number - 1]):
                text = _text(line)
                if _PAGE_HEADER.search(text) or re.fullmatch(r"\d{1,4}", text):
                    continue
                m = _SECTION.match(text)
                if m:
                    code, table, state = m.group(1), None, None
                    continue
                if code not in codes:
                    continue
                if text.lower().startswith("unless otherwise specified"):
                    state, header, dates, pending = "header", [], [], None
                    table = None
                    continue
                if "[Table]" in text:
                    title = re.sub(r"\s*\[Table\].*$", "", text)
                    continue
                if state == "header":
                    date_line = all(_DATE.match(w["text"]) or w["text"].lower() == "to" for w in line)
                    if date_line:
                        dates += [w for w in line if _DATE.match(w["text"])]
                        continue
                    if not dates:
                        header += line              # axis / member names above the dates
                        continue
                    columns = _columns(dates)
                    zone = min(c.x1 for c in columns) - 55
                    axis = _text([w for w in header if w["x0"] < zone and "[Axis]" not in w["text"]])
                    names = [w for w in header if w["x0"] >= zone]
                    table = MTable(code, title, axis, columns, _members(columns, names))
                    table.zone = zone
                    tables.append(table)
                    state = "rows"
                if table is None:
                    continue
                label = [w for w in line if w["x1"] <= table.zone]
                cells = [w for w in line if w["x1"] > table.zone]
                values, texts = {}, {}
                # a description cell's words are text, numbers included ("maturity more than 12 months")
                describing = re.match(r"description|nature of", _text(label) or (pending or "") or
                                      (table.rows[-1].label if table.rows and not label else ""), re.I)
                for w in cells:
                    if not describing and _NUMBER.match(w["text"]) and _number(w["text"]) is not None:
                        col = min(range(len(table.columns)), key=lambda i: abs(table.columns[i].x1 - w["x1"]))
                        values[col] = _number(w["text"])
                    else:
                        c = (w["x0"] + w["x1"]) / 2
                        col = min(range(len(table.columns)), key=lambda i: abs(table.columns[i].center - c))
                        texts[col] = f"{texts.get(col, '')} {w['text']}".strip()
                name = _text(label)
                if re.search(r"\[(?:abstract|line|axis|member)", name, re.I) or name.lower() == "items]":
                    pending = None
                    continue
                if not name and (values or texts) and table.rows and pending is None:
                    last = table.rows[-1]          # a wrapped text cell, or figures under a wrapped label
                    if values and not last.values:
                        last.values.update(values)
                    for col, t in texts.items():
                        last.texts[col] = f"{last.texts.get(col, '')} {t}".strip()
                    continue
                if pending is not None:
                    name = f"{pending} {name}".strip()
                    pending = None
                if name and not values and not texts:
                    pending = name                 # a label wrapped over lines
                    continue
                table.rows.append(TRow(name, values, texts, bool(re.match(r"total\b", name, re.I)), number))
    return tables
