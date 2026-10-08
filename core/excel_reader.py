"""
Reads a company's financial statements from an Excel workbook (Balance Sheet,
Profit & Loss and Notes sheets laid out like the printed statements:
Particulars | Note | current year | previous year ...).

The sheets are turned into the same text lines a digital PDF yields and then
go through the same parsing, note-tie checks and mapping as a PDF. The values
come straight from the cells (cached results of any formulas), so there is no
OCR risk; only the columns under a year heading are read, so side calculations
and comments typed in other columns are ignored.
"""

import datetime
import io
import re

import openpyxl

from core.pdf_reader import PageText, classify_page, extract_from_pages

_UNIT_CELL = re.compile(r"amount\s+in|in\s+rs|in\s+lakhs|in\s+crores|in\s+thousands|'000", re.I)


def _number_text(value):
    text = f"{abs(value):,.2f}"
    return f"({text})" if value < 0 else text


def _year_header(value):
    """A header cell that names a financial year -> text the PDF logic understands, else None."""
    if isinstance(value, (datetime.datetime, datetime.date)):
        return f"{value.strftime('%d %b %Y')}"
    if isinstance(value, (int, float)) and 1990 <= value <= 2100 and float(value).is_integer():
        return f"31 Mar {int(value)}"
    if isinstance(value, str) and re.search(r"(?:19|20)\d\d\s*[-/]\s*\d\d|(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*[\s,.'-]*(?:19|20)\d\d",
                                            value, re.I):
        return value.strip()
    return None


def _find_header(rows):
    """(row index, {column index: header text}) of the first row with at least two year headings."""
    for idx, row in enumerate(rows[:25]):
        cols = {c: t for c, v in enumerate(row) if (t := _year_header(v))}
        if len(cols) >= 2:
            return idx, cols
    return None, {}


def _sheet_lines(rows):
    header_idx, year_cols = _find_header(rows)
    if header_idx is None:
        return [], False
    first_year_col = min(year_cols)
    lines = []
    unit_seen = False
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
            lines.append("Particulars Notes " + " ".join(year_cols[c] for c in sorted(year_cols)))
            continue
        left = [v for v in row[:first_year_col] if v is not None and str(v).strip() != ""]
        label = " ".join(str(v).strip() for v in left if isinstance(v, str))
        small_ints = [int(v) for v in row[:first_year_col]
                      if isinstance(v, (int, float)) and not isinstance(v, bool) and float(v).is_integer() and 0 < v < 100]
        values = [row[c] if c < len(row) else None for c in sorted(year_cols)]
        numeric = [v if isinstance(v, (int, float)) and not isinstance(v, bool) else None for v in values]
        if all(v is None for v in numeric):
            if not label and small_ints:
                continue
            if label and small_ints and len(small_ints) == 1:
                lines.append(f"{small_ints[0]}. {label}")        # a note heading: "4 | Reserves And Surplus"
            elif label:
                lines.append(label)
            continue
        if not label:
            continue
        note = f" {small_ints[0]}" if small_ints else ""
        numbers = " ".join(_number_text(v) if v is not None else "-" for v in numeric)
        lines.append(f"{label}{note} {numbers}")
    return lines, True


_STATEMENT_SHEET = re.compile(r"^\s*(bs|balance|b\s*/\s*s|p\s*&\s*l|pnl|profit|income|cash\s*flow|cf|notes?)\s*(sheet|statement)?\s*$", re.I)


def read_excel_pages(source):
    """source: path or bytes. One PageText per sheet that looks like a statement or a notes sheet."""
    handle = io.BytesIO(source) if isinstance(source, (bytes, bytearray)) else source
    wb = openpyxl.load_workbook(handle, data_only=True, read_only=True)
    pages = []
    for number, ws in enumerate(wb.worksheets, start=1):
        if ws.sheet_state != "visible" or not _STATEMENT_SHEET.match(ws.title):
            continue
        rows = [list(r) for r in ws.iter_rows(max_row=400, max_col=40, values_only=True)]
        lines, ok = _sheet_lines(rows)
        if not ok:
            continue
        kind = classify_page(lines)
        if kind == "OTHER" and re.search(r"note", ws.title, re.I):
            kind = "NOTES"
        if kind == "OTHER":
            continue
        pages.append(PageText(number, lines, "text"))
    return pages


def extract_excel(source):
    pages = read_excel_pages(source)
    if not pages:
        raise ValueError("No Balance Sheet, Profit & Loss or Notes sheet with year columns was found in this workbook.")
    result = extract_from_pages(pages)
    result.warnings.append("Read from Excel: figures are the cell values as saved in the workbook.")
    return result
