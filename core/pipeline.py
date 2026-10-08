"""
Glue between the pieces, shared by the web app and any script: load input,
map it to LineItems, apply the user's review edits, check the balance, write
the workbook. No Streamlit in here, so it can be tested on its own.
"""

import re
import tempfile
from pathlib import Path

from config.schedule_map import FY_COLUMN_MAP
from core import classifier, io_loader
from core.balance_check import balance_check
from core.pdf_mapper import map_extraction
from core.pdf_reader import extract_financials
from core.render_writer import write_workbook
from core.resolver import resolve

TEMPLATE_PATH = Path(__file__).resolve().parent.parent / "template" / "commonsize_template.xlsx"
_PAT_LABEL = re.compile(r"profit\s*(?:/\s*\(?loss\)?\s*)?(?:for\s+the\s+(?:year|period)|after\s+tax)", re.I)


def read_pdf(pdf_bytes, progress=None):
    """-> (ExtractionResult, MappingResult)"""
    extraction = extract_financials(pdf_bytes, progress=progress)
    return extraction, map_extraction(extraction)


def read_structured(path):
    """Statement / Line Item / Statement Tag / FY.. sheet -> classified LineItems (values taken as rupees)."""
    items = io_loader.load_structured(path)
    classifier.classify_all(items)
    for item in items:
        needs_help = (item.statement == "P&L" and item.category.startswith("Uncertain")) or \
                     (item.statement == "BS" and not item.statement_tag)
        if not needs_help:
            continue
        res = resolve(item.line_item, "", item.statement)
        if res.resolved:
            item.statement_tag, item.category = res.tag, res.category
            item.matched_item, item.is_new = res.matched_item, not res.known
            item.confidence = "Matched" if res.known else "Uncertain (new item)"
    return items


def reported_pat(extraction):
    """Profit after tax exactly as the financials print it, in rupees {fy: value}."""
    for row in extraction.rows:
        if row.statement == "P&L" and _PAT_LABEL.search(row.label) and not re.search(r"before", row.label, re.I):
            if all(row.values.get(fy) is not None for fy in extraction.fy_columns):
                return {fy: row.values[fy] * extraction.unit_multiplier for fy in extraction.fy_columns}
    # the face line is sometimes unreadable in a scan; the reserves note repeats the same profit
    for row in extraction.rows:
        if row.statement == "NOTE" and re.search(r"net\s*profit|profit\s*for\s*the", row.label, re.I)                 and re.search(r"reserve|surplus", row.heading or "", re.I):
            if all(row.values.get(fy) is not None for fy in extraction.fy_columns):
                return {fy: row.values[fy] * extraction.unit_multiplier for fy in extraction.fy_columns}
    return {}


def tolerance(unit_multiplier):
    """Rounding in the financials' own unit (two decimals) is not a failure."""
    return max(1.0, 0.02 * unit_multiplier)


def check(items, fy_columns, unit_multiplier):
    fys = [fy for fy in sorted(fy_columns) if fy in FY_COLUMN_MAP]
    return balance_check(items, fys, tolerance(unit_multiplier))


def build_workbook(items, fy_columns, company, equity_shares=None, reported=None):
    """-> (xlsx bytes, WriteReport)"""
    fys = [fy for fy in sorted(fy_columns) if fy in FY_COLUMN_MAP]
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "commonsize.xlsx"
        report = write_workbook(TEMPLATE_PATH, out, company, items, fys, equity_shares=equity_shares,
                                reported_pat=reported)
        return out.read_bytes(), report


# ----------------------------------------------------------------------------- reconciliation with the input
_FACE_LINES = [
    # (key, label shown, statement, regex on the printed label)
    ("revenue", "Revenue from operations", "P&L", r"^revenue\s*from\s*operations"),
    ("total_income", "Total income", "P&L", r"^total\s*(income|revenue)"),
    ("total_expenses", "Total expenses", "P&L", r"^total\s*expenses"),
    ("pbt", "Profit before tax", "P&L", r"^profit\s*before\s*tax"),
    ("tax", "Tax expense", "P&L", r"^total\s*tax|^tax\s*expense"),
]


def _face_total(extraction, statement, pattern, fy, skip=0):
    hits = [r for r in extraction.rows if r.statement == statement and re.search(pattern, r.label, re.I)
            and r.values.get(fy) is not None]
    return hits[skip].values[fy] * extraction.unit_multiplier if len(hits) > skip else None


def reconcile(extraction, items, fy_columns):
    """
    Compares what the app computed with the figures the input itself prints.
    -> list of dict(year, line, computed, reported, diff, ok). Lines the input does not show are left out.
    """
    fys = [fy for fy in sorted(fy_columns) if fy in FY_COLUMN_MAP]
    result = check(items, fys, extraction.unit_multiplier)
    tol = tolerance(extraction.unit_multiplier)
    pat = reported_pat(extraction)
    out = []
    for fy in fys:
        s = result.statements[fy]
        computed = {
            "revenue": s["revenue"],
            "total_income": s["revenue"] + s["other_income"],
            "total_expenses": s["revenue"] + s["other_income"] - (s["operating_pbt"] + s["other_income"]),
            "pbt": s["operating_pbt"] + s["other_income"],
            "tax": s["tax"],
        }
        for key, label, statement, pattern in _FACE_LINES:
            reported = _face_total(extraction, statement, pattern, fy)
            if reported is not None:
                out.append(dict(year=fy, line=label, computed=computed[key], reported=reported))
        if fy in pat:
            out.append(dict(year=fy, line="Profit for the year", computed=s["pat"], reported=pat[fy]))
        liab, assets = (_face_total(extraction, "BS", r"^total$", fy, skip=k) for k in (0, 1))
        total_liab = s["liabilities_side"] + s["current_liabilities"] + max(-s["deferred_tax_net"], 0.0)
        if liab is not None:
            out.append(dict(year=fy, line="Total equity and liabilities", computed=total_liab, reported=liab))
        if assets is not None:
            out.append(dict(year=fy, line="Total assets", computed=s["total_assets"], reported=assets))
    for row in out:
        row["diff"] = row["computed"] - row["reported"]
        row["ok"] = abs(row["diff"]) <= tol
    return out


def source_balance(extraction, fy_columns):
    """{fy: (reported liabilities total, reported assets total)} for years where the input prints both."""
    out = {}
    for fy in fy_columns:
        liab = _face_total(extraction, "BS", r"^total$", fy, skip=0)
        assets = _face_total(extraction, "BS", r"^total$", fy, skip=1)
        if liab is not None and assets is not None:
            out[fy] = (liab, assets)
    return out


def suggest_fixes(extraction, items, fy_columns):
    """
    Where a face figure disagrees with its own note, test whether using the note's figure makes the year's
    profit agree with the profit the financials print. Only fixes that do are returned, so a suggestion is
    backed by evidence, never a guess: list of (item, fy, current, suggested).
    """
    pat = reported_pat(extraction)
    tol = tolerance(extraction.unit_multiplier)
    found = []
    for fy in fy_columns:
        if fy not in pat:
            continue
        base = check(items, fy_columns, extraction.unit_multiplier).statements[fy]["pat"]
        if abs(base - pat[fy]) <= tol:
            continue
        for item in items:
            alt = item.alt_values.get(fy)
            if alt is None or item.fy_values.get(fy) is None:
                continue
            original = item.fy_values[fy]
            item.fy_values[fy] = alt
            try:
                trial = check(items, fy_columns, extraction.unit_multiplier).statements[fy]["pat"]
            finally:
                item.fy_values[fy] = original
            if abs(trial - pat[fy]) <= tol:
                found.append((item, fy, original, alt))
    return found


def read_excel(xlsx_bytes):
    """Excel financial statements -> (ExtractionResult, MappingResult)"""
    from core.excel_reader import extract_excel
    extraction = extract_excel(xlsx_bytes)
    return extraction, map_extraction(extraction)


def remap(extraction, unit_multiplier):
    """Re-run the mapping with a different unit (no re-reading of the file)."""
    return map_extraction(extraction, unit_multiplier)
