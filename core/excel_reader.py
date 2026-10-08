"""
Reads a company's financial statements from an Excel workbook (Balance Sheet,
Profit & Loss and Notes / Schedules sheets laid out like the printed statements:
Particulars | Note | current year | previous year ...).

The sheets are turned into the same text lines a digital PDF yields and then
go through the same parsing, note-tie checks and mapping as a PDF. The values
come straight from the cells (cached results of any formulas), so there is no
OCR risk.

Real workbooks vary a lot, so the reader goes by content, not by sheet names:
  * a sheet is a Balance Sheet / P&L / Notes sheet if its title lines say so;
  * year columns are found from their headings (dates, "AS AT 31.03.2026",
    "F.Y. 2024-25" ...), possibly split over two rows;
  * extra columns (branch / state-wise figures, side calculations, an old 2014
    column) are ignored: only the years shown on the face statements are read,
    from the consolidated columns.
"""

import datetime
import io
import re

import openpyxl

from core.pdf_reader import PageText, classify_page, extract_from_pages

_UNIT_CELL = re.compile(r"amount\s+in|amt\.?\s*in|in\s+rs|in\s+lakhs|in\s+crores|in\s+thousands|'000", re.I)
_NIL = {"nil", "-", "--", "—", "–", "n.a.", "na", ""}
_HEADER_WORDS = re.compile(r"^(?:particulars|notes?|note\s*no\.?|no\.?|sr\.?\s*no\.?|as\s*at|as\s*on|year\s*ended|for\s+the\s+year"
                           r"(?:\s+ended)?|amount|amt\.?|rs\.?|rupees|\(rupees\)|current\s+year|previous\s+year)\.?$", re.I)
_ENUMERATOR = re.compile(r"^(?:\(?(?:[a-zA-Z]|[ivxIVX]{1,4})\)\s*|(?:[ivxIVX]{1,4}|\d{1,2})\.\s*|[a-zA-Z]\.\s+)+")
_MONTHS = {m: i for i, m in enumerate(("jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"), 1)}
_PREFERRED_TITLE = re.compile(r"^\s*(bs|balance|b\s*/\s*s|p\s*&\s*l|pnl|profit|income|cash\s*flow|cf|notes?)\s*(sheet|statement)?\s*$",
                              re.I)


def _fy_from_date(year, month):
    """Indian financial year label of a balance-sheet date: 31 Mar 2026 -> FY26, 30 Sep 2025 -> FY26."""
    return f"FY{(year if month <= 3 else year + 1) % 100:02d}"


def _fy_of(value):
    """FY label if the cell names a financial-year column, else None."""
    if isinstance(value, (datetime.datetime, datetime.date)):
        return _fy_from_date(value.year, value.month)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f"FY{int(value) % 100:02d}" if 1990 <= value <= 2100 and float(value).is_integer() else None
    if not isinstance(value, str):
        return None
    text = value.strip()
    m = re.search(r"\b(\d{1,2})[./-](\d{1,2})[./-]((?:19|20)\d\d)\b", text)            # 31.03.2026
    if m:
        return _fy_from_date(int(m.group(3)), int(m.group(2)))
    m = re.search(r"\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*[\s,.'-]*((?:19|20)\d\d)\b", text, re.I)
    if m:
        return _fy_from_date(int(m.group(2)), _MONTHS[m.group(1).lower()])
    m = re.search(r"\b(?:19|20)\d\d\s*[-/]\s*(\d\d)\b", text)                           # 2024-25
    if m:
        return f"FY{m.group(1)}"
    return None


def _is_number(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _is_branch_name(v):
    """A short column title such as a branch or state name ("MAHARASHTRA", "AP"), not a heading or a unit."""
    if not isinstance(v, str):
        return False
    t = v.strip()
    return (2 <= len(t) <= 30 and not re.search(r"\d", t) and not _HEADER_WORDS.match(t) and not _UNIT_CELL.search(t)
            and t.lower() not in _NIL and _fy_of(t) is None)


def _find_columns(rows):
    """
    -> (header row index, {FY: column}, first data column) or (None, {}, None).
    A column headed by a year whose heading block also names a branch / state is not a consolidated year
    column and is skipped. The first row with two or more year columns wins; else the first with one.
    """
    fallback = None
    for idx, row in enumerate(rows[:40]):
        cand = {c: fy for c, v in enumerate(row) if (fy := _fy_of(v))}
        if not cand:
            continue
        block = rows[max(0, idx - 2): idx + 3]
        named = {c for r in block for c, v in enumerate(r) if c >= 2 and _is_branch_name(v)}
        kept = {}
        for c in sorted(cand):
            if c not in named:
                kept.setdefault(cand[c], c)
        if not kept:
            continue
        data_start = min(list(named) + list(kept.values()))
        if len(kept) >= 2:
            return idx, kept, data_start
        if fallback is None:
            fallback = (idx, kept, data_start)
    return fallback or (None, {}, None)


def _number_text(value):
    text = f"{abs(value):,.2f}"
    return f"({text})" if value < 0 else text


def _cell_value(v):
    """Number, 0.0 for nil markers, None for blank / error / text."""
    if _is_number(v):
        return float(v)
    if isinstance(v, str) and v.strip().lower() in _NIL - {""}:
        return 0.0
    return None


def _sheet_lines(rows, fy_order):
    """Text lines for one sheet; values only for `fy_order` (the face statements' years), in that order."""
    header_idx, fy_cols, data_start = _find_columns(rows)
    if header_idx is None:
        return [], {}
    lines, unit_seen = [], False
    for idx, row in enumerate(rows):
        if idx < header_idx:
            for v in row:
                if isinstance(v, str) and v.strip():
                    if _UNIT_CELL.search(v):
                        if unit_seen:
                            continue    # a second unit cell is usually a side calculation, not the statement unit
                        unit_seen = True
                    lines.append(v.strip())
            continue
        if idx == header_idx:
            years = [fy for fy in (fy_order or sorted(fy_cols, reverse=True)) if fy in fy_cols]
            lines.append("Particulars Notes " + " ".join(f"31 Mar 20{fy[2:]}" for fy in years))
            continue
        head = row[:data_start]
        label = " ".join(str(v).strip() for v in head if isinstance(v, str) and v.strip()
                         and v.strip().lower() not in _NIL and not v.strip().startswith("#"))
        label = _ENUMERATOR.sub("", label).strip()      # "a) Share Capital", "(iii) ...", "1.Shareholders Fund"
        small_ints = [int(v) for v in head if _is_number(v) and float(v).is_integer() and 0 < v < 100]
        cols = [fy_cols.get(fy) for fy in (fy_order or sorted(fy_cols, reverse=True))]
        values = [(_cell_value(row[c]) if c is not None and c < len(row) else None) for c in cols]
        if all(v is None for v in values):
            if not label:
                continue
            if small_ints and len(small_ints) == 1 and not re.match(r"^note\b", label, re.I):
                lines.append(f"{small_ints[0]}. {label}")        # a note heading: "4 | Reserves And Surplus"
            else:
                lines.append(label)
            continue
        if not label:
            continue
        note = f" {small_ints[0]}" if small_ints else ""
        numbers = " ".join(_number_text(v) if v is not None else "-" for v in values)
        lines.append(f"{label}{note} {numbers}")
    return lines, fy_cols


def read_excel_pages(source):
    """source: path or bytes. One PageText per sheet that is a Balance Sheet, P&L, Cash Flow or Notes sheet."""
    handle = io.BytesIO(source) if isinstance(source, (bytes, bytearray)) else source
    wb = openpyxl.load_workbook(handle, data_only=True, read_only=True)
    sheets = []
    for number, ws in enumerate(wb.worksheets, start=1):
        if ws.sheet_state != "visible":
            continue
        rows = [list(r) for r in ws.iter_rows(max_row=600, max_col=40, values_only=True)]
        lines, fy_cols = _sheet_lines(rows, None)
        if not lines:
            continue
        kind = classify_page(lines)
        if kind == "OTHER" and re.search(r"note|sch", ws.title, re.I) and any(re.match(r"^\W*note\s*\d", l, re.I) for l in lines):
            kind = "NOTES"
        if kind != "OTHER":
            sheets.append(dict(number=number, title=ws.title, kind=kind, rows=rows, fy_cols=fy_cols))

    chosen = []
    for kind in ("BS", "P&L", "CF"):
        options = [s for s in sheets if s["kind"] == kind]
        preferred = [s for s in options if _PREFERRED_TITLE.match(s["title"])]
        if options:
            chosen.append((preferred or options)[0])
    notes = [s for s in sheets if s["kind"] == "NOTES"]
    exact = [s for s in notes if re.fullmatch(r"\s*notes?\s*", s["title"], re.I)]
    chosen += exact or notes
    chosen.sort(key=lambda s: s["number"])

    # the years to read: those on the face statements, in their column order (latest first as printed)
    face = next((s for s in chosen if s["kind"] in ("BS", "P&L")), None)
    if face is None:
        return []
    fy_order = sorted(face["fy_cols"], key=lambda fy: face["fy_cols"][fy])
    pages = []
    for s in chosen:
        lines, _ = _sheet_lines(s["rows"], fy_order)
        pages.append(PageText(s["number"], lines, "text"))
    return pages


def extract_excel(source):
    pages = read_excel_pages(source)
    if not pages:
        raise ValueError("No Balance Sheet, Profit & Loss or Notes sheet with year columns was found in this workbook.")
    result = extract_from_pages(pages)
    for row in result.rows:   # scan-repair flags mean nothing for cell values
        row.flags = [f for f in row.flags if f not in ("bracket_noise", "decimal_repaired")]
    result.warnings.append("Read from Excel: figures are the cell values as saved in the workbook.")
    return result
