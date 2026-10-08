"""
Input loading.

Two paths:
  - load_structured(): reads the Statement/Line Item/Statement Tag/FY.. schema
    from an .xlsx or .csv file into a list of LineItem records.
  - extract_pdf_tables_for_review(): best-effort table extraction from a PDF
    Annual Report via pdfplumber. This does NOT attempt to classify or map
    anything -- it just dumps every table it finds into a workbook (one sheet
    per table) so the user can manually convert the relevant rows into the
    structured schema above. Auto-mapping PDF tables to line items reliably
    needs an LLM and is deferred to a later phase.
"""

import csv
import re
from pathlib import Path

import openpyxl

from core.models import LineItem

FY_COLUMN_PATTERN = re.compile(r"^FY\s*\d{2,4}$", re.IGNORECASE)


def _find_fy_columns(headers):
    return [h for h in headers if h and FY_COLUMN_PATTERN.match(str(h).strip())]


def _parse_amount(raw):
    if raw is None or str(raw).strip() == "":
        return None
    try:
        return float(str(raw).replace(",", "").strip())
    except ValueError:
        return None


def load_structured(path):
    """Load a Statement/Line Item/Statement Tag/FY.. sheet into LineItem records."""
    path = Path(path)
    if path.suffix.lower() == ".csv":
        return _load_structured_csv(path)
    return _load_structured_xlsx(path)


def _load_structured_xlsx(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    headers = [str(h).strip() if h is not None else "" for h in rows[0]]
    fy_columns = _find_fy_columns(headers)
    if not fy_columns:
        raise ValueError(
            "No FY.. columns found in header row. Expected columns like "
            "FY22, FY23, FY24 alongside Statement / Line Item / Statement Tag."
        )
    items = []
    for row in rows[1:]:
        record = dict(zip(headers, row))
        line_item_name = record.get("Line Item")
        if not line_item_name or not str(line_item_name).strip():
            continue
        fy_values = {fy: _parse_amount(record.get(fy)) for fy in fy_columns}
        items.append(
            LineItem(
                statement=str(record.get("Statement") or "").strip().upper(),
                line_item=str(line_item_name).strip(),
                statement_tag=str(record.get("Statement Tag") or "").strip(),
                fy_values=fy_values,
            )
        )
    return items


def _load_structured_csv(path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        headers = [h.strip() for h in (reader.fieldnames or [])]
        fy_columns = _find_fy_columns(headers)
        if not fy_columns:
            raise ValueError(
                "No FY.. columns found in header row. Expected columns like "
                "FY22, FY23, FY24 alongside Statement / Line Item / Statement Tag."
            )
        items = []
        for record in reader:
            line_item_name = record.get("Line Item")
            if not line_item_name or not line_item_name.strip():
                continue
            fy_values = {fy: _parse_amount(record.get(fy)) for fy in fy_columns}
            items.append(
                LineItem(
                    statement=(record.get("Statement") or "").strip().upper(),
                    line_item=line_item_name.strip(),
                    statement_tag=(record.get("Statement Tag") or "").strip(),
                    fy_values=fy_values,
                )
            )
    return items


def extract_pdf_tables_for_review(pdf_path, output_dir):
    """
    Extract every table pdfplumber can find in the PDF and dump it into a
    review workbook (one sheet per table). Returns the path to that workbook.
    Known limitation: no attempt is made to identify which table is the P&L,
    which is the Balance Sheet, or to map row labels to the structured
    schema -- that requires reasoning this no-API prototype doesn't have.
    """
    import pdfplumber

    pdf_path = Path(pdf_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    review_path = output_dir / f"{pdf_path.stem}_extracted_tables_for_review.xlsx"

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    table_count = 0
    with pdfplumber.open(pdf_path) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            for table_index, table in enumerate(page.extract_tables(), start=1):
                if not table:
                    continue
                table_count += 1
                sheet_name = f"p{page_number}_t{table_index}"[:31]
                ws = wb.create_sheet(title=sheet_name)
                for row in table:
                    ws.append(["" if c is None else c for c in row])

    if table_count == 0:
        ws = wb.create_sheet(title="No tables found")
        ws.append(["pdfplumber did not detect any tables in this PDF."])

    wb.save(review_path)
    return review_path, table_count
