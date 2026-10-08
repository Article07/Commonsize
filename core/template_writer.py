"""
Populates the real Reference file template (IS/BS/CF/Sch./Ratios) via Excel
COM automation instead of writing a standalone workbook. Populates both the
P&L and Balance Sheet schedules on the "Sch." sheet (Phase 2 + Phase 3 of the
project plan) -- IS/BS/CF/Ratios then compute themselves entirely through
the template's own pre-existing formulas.

Row insertion for a brand-new line item uses Excel's own "insert copied
cells" behaviour (Rows(source).Copy() then Rows(target).Insert(...)):
Excel adjusts every relative formula reference itself (the per-row % of
Revenue / YoY formulas already built into rows 195-247, the schedule's own
Total SUM() range, and every cross-sheet reference from IS/BS/CF/Ratios into
Sch.). This is why COM automation was chosen over openpyxl for this step --
openpyxl's insert_rows() does not rewrite formulas elsewhere in the workbook.
"""

import copy
import re
import shutil
from pathlib import Path

import win32com.client as win32

from config.classification_rules import (
    COGS_MANUFACTURING_KEYWORDS,
    COGS_PURCHASE_KEYWORDS,
    REVENUE_SERVICE_KEYWORDS,
)
from config.schedule_map import ALL_SCHEDULE_ORDER, ALL_SCHEDULES, FY_COLUMN_MAP, SHEET_NAME

XL_SHIFT_DOWN = -4121


def _normalize(text):
    return re.sub(r"[^a-z0-9 ]", "", text.lower()).strip()


BS_TAG_TO_SCHEDULE = {
    "Trade Receivables": "Trade Receivables",
    "Inventory": "Inventory",
    "Trade Payables": "Trade Payables",
    "Cash & Bank": "Cash & Bank",
    "Fixed Assets": "Fixed Assets",
    "Investments": "Investments",
    "Current Asset - Other": "Current Asset - Other",
    "Current Liability - Other": "Current Liability - Other",
    "Borrowings - Long Term": "Borrowings - Long Term",
    "Borrowings - Short Term": "Borrowings - Short Term",
}


def _route_schedule_key(line_item, log):
    tag = line_item.statement_tag

    if line_item.statement == "P&L":
        if tag == "Revenue":
            return "Revenue"
        if tag == "Depreciation":
            return "Depreciation"
        if tag == "Tax Expense":
            return "Tax Expense"

        category = line_item.category
        if category == "Direct Expenses / COGS":
            name_lower = line_item.line_item.lower()
            if any(kw in name_lower for kw in COGS_PURCHASE_KEYWORDS):
                return "COGS-Purchase"
            if any(kw in name_lower for kw in COGS_MANUFACTURING_KEYWORDS):
                return "COGS-Manufacturing"
            return "COGS-Manufacturing"  # matched COGS overall but neither sub-list -- safer default
        if category == "Administrative Expenses":
            return "Administrative Expenses"
        if category == "Selling & Distribution Expenses":
            return "Selling & Distribution Expenses"
        if category == "Finance Costs":
            return "Finance Costs"
        if category == "Other Income":
            return "Other Income"
        return None  # Uncertain -- Needs Review: never force a row into the template

    if line_item.statement == "BS":
        if tag == "Paid-up Share Capital":
            return "Share Capital"
        if tag == "Opening Reserves":
            return "Reserves and Surplus"
        if tag == "Borrowings":
            log.append(
                f"WARNING: '{line_item.line_item}' uses the generic 'Borrowings' tag -- "
                "routing to Short Term. Use 'Borrowings - Long Term'/'Borrowings - Short Term' "
                "for precise placement."
            )
            return "Borrowings - Short Term"
        if tag in BS_TAG_TO_SCHEDULE:
            return BS_TAG_TO_SCHEDULE[tag]
        return None  # untagged BS row, or a tag with no Sch. destination in this phase

    return None  # CF rows: no Sch. destination, CF is fully formula-derived


def _revenue_target_row(line_item, state):
    name_lower = line_item.line_item.lower()
    sched = state["Revenue"]
    if any(kw in name_lower for kw in REVENUE_SERVICE_KEYWORDS):
        return sched["first_item_row"] + 1  # "Sale of services"
    return sched["first_item_row"]  # "Sale of goods"


def _write_share_capital(ws, li, fy_columns, state):
    row = state["Share Capital"]["first_item_row"]
    # Keep the template's own '"x" equity shares of Rs. 10 each' label --
    # only the FY value columns are the company's paid-up capital amount.
    _write_values(ws, row, li, fy_columns, overwrite_label=False)


def _write_opening_reserves(ws, li, fy_columns, state, log):
    """
    Only the EARLIEST FY's Opening balance cell is a free input in the
    template -- every later FY's Opening balance is already a formula
    chaining to the prior year's Closing balance (D17=C21, E17=D21, ...).
    Writing into those would destroy the roll-forward, so any values
    supplied for later FYs are intentionally ignored (with a log note).
    """
    if not fy_columns:
        return
    row = state["Reserves and Surplus"]["first_item_row"]
    earliest_fy = fy_columns[0]
    template_col = FY_COLUMN_MAP.get(earliest_fy)
    value = li.fy_values.get(earliest_fy)
    if template_col and value is not None:
        ws.Range(f"{template_col}{row}").Value = value

    ignored_fys = [fy for fy in fy_columns[1:] if li.fy_values.get(fy) is not None]
    if ignored_fys:
        log.append(
            f"NOTE: 'Opening Reserves' values for {', '.join(ignored_fys)} were ignored -- "
            "the template auto-chains each year's opening balance from the prior year's "
            "closing balance; only the earliest FY's opening balance is a free input."
        )


def _apply_equity_share_count(ws, state, equity_shares):
    if equity_shares is None:
        return
    row = state["Share Capital"]["first_item_row"]
    ws.Range(f"B{row}").Value = f"{equity_shares} equity shares of Rs. 10 each"


def _find_matching_row(ws, first_row, last_row, item_name):
    norm_item = _normalize(item_name)
    if first_row == last_row:
        raw_labels = [ws.Range(f"B{first_row}").Value]
    else:
        raw_labels = [row[0] for row in ws.Range(f"B{first_row}:B{last_row}").Value]

    for offset, label in enumerate(raw_labels):
        if not label:
            continue
        norm_label = _normalize(str(label))
        if len(norm_label) < 4:
            continue
        if norm_label in norm_item or norm_item in norm_label:
            return first_row + offset
    return None


def _shift_after(state, from_key, delta=1):
    idx = ALL_SCHEDULE_ORDER.index(from_key)
    for key in ALL_SCHEDULE_ORDER[idx + 1 :]:
        for field in ("title_row", "first_item_row", "last_item_row", "total_row"):
            state[key][field] += delta


def _insert_row(ws, schedule_key, state):
    sched = state[schedule_key]
    target_row = sched["last_item_row"]
    source_row = sched["first_item_row"]

    ws.Rows(source_row).Copy()
    ws.Rows(target_row).Insert(Shift=XL_SHIFT_DOWN)
    ws.Application.CutCopyMode = False

    sched["last_item_row"] += 1
    if sched["pattern"] == "total_last":
        sched["total_row"] += 1
    _shift_after(state, schedule_key, delta=1)

    return target_row  # the newly inserted (now-blank-ish, copied-formula) row


def _write_values(ws, row, line_item, fy_columns, overwrite_label=True):
    if overwrite_label:
        ws.Range(f"B{row}").Value = line_item.line_item
    for fy in fy_columns:
        template_col = FY_COLUMN_MAP.get(fy)
        if not template_col:
            continue
        value = line_item.fy_values.get(fy)
        if value is not None:
            ws.Range(f"{template_col}{row}").Value = value


def _verify_totals(ws, state, log):
    for key, sched in state.items():
        if sched["pattern"] == "total_last":
            formula = ws.Range(f"C{sched['total_row']}").Formula
            expected_fragment = f"{sched['first_item_row']}:C{sched['last_item_row']}"
            if str(sched["first_item_row"]) not in str(formula) and expected_fragment not in str(formula):
                log.append(
                    f"WARNING: '{key}' Total row {sched['total_row']} formula does not "
                    f"appear to reference the full item range ({formula!r}). Please verify manually."
                )


def _write_analysis_sheet_com(wb, company_name, sections, irl_rows):
    """
    Adds an Analysis+IRL sheet via COM, in the same session/save as the Sch.
    schedule edits. Deliberately NOT done with openpyxl afterwards: reopening
    a formula-heavy workbook with openpyxl (data_only=False) discards every
    cached formula result on re-save, which would blank out IS/BS/CF/Ratios.
    """
    existing_names = [s.Name for s in wb.Sheets]
    if "Analysis" in existing_names:
        wb.Sheets("Analysis").Delete()
    ws = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
    ws.Name = "Analysis"

    row = 1
    ws.Range(f"A{row}").Value = f"{company_name} -- Analysis"
    row += 2

    for title in [
        "Executive Summary",
        "YoY Variance Analysis",
        "Common Size Analysis",
        "Ratio Commentary",
        "Data Integrity",
    ]:
        ws.Range(f"A{row}").Value = title
        row += 1
        findings = sections.get(title, [])
        if not findings:
            ws.Range(f"A{row}").Value = "No findings for this section."
            row += 2
            continue
        for finding in findings:
            ws.Range(f"A{row}").Value = finding.area
            ws.Range(f"B{row}").Value = finding.observation
            ws.Range(f"C{row}").Value = finding.reason_for_concern
            ws.Range(f"D{row}").Value = finding.info_required
            row += 1
        row += 1

    ws.Range(f"A{row}").Value = "Information Requirement List (IRL)"
    row += 1
    irl_headers = [
        "Reference Number",
        "Financial Statement Area",
        "Observation",
        "Reason for Concern",
        "Information / Clarification Required",
    ]
    for col_idx, header in enumerate(irl_headers, start=1):
        ws.Cells(row, col_idx).Value = header
    row += 1
    for irl_row in irl_rows:
        for col_idx, key in enumerate(irl_headers, start=1):
            ws.Cells(row, col_idx).Value = irl_row[key]
        row += 1

    used_range = ws.UsedRange
    used_range.Font.Name = "Ebrima"
    used_range.Font.Size = 10
    ws.Columns("A").ColumnWidth = 30
    ws.Columns("B:E").ColumnWidth = 45


def populate_template(
    template_path,
    output_path,
    company_name,
    line_items,
    fy_columns,
    sections=None,
    irl_rows=None,
    equity_shares=None,
    log=None,
):
    """
    Copies the template to output_path, populates its Sch. schedules (P&L
    and, from Phase 3, Balance Sheet) from the classified line items, and
    (if sections/irl_rows are given) appends an Analysis+IRL sheet in the
    same COM session. Returns (log_lines, unmapped_items).
    """
    log = log if log is not None else []
    template_path = Path(template_path)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    shutil.copyfile(template_path, output_path)

    state = copy.deepcopy(ALL_SCHEDULES)
    unmapped_items = []

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    save_changes = False
    try:
        wb = excel.Workbooks.Open(str(output_path.resolve()))
        is_sheet = wb.Sheets("IS")
        ws = wb.Sheets(SHEET_NAME)

        is_sheet.Range("B2").Value = company_name

        items_to_route = [li for li in line_items if li.statement in ("P&L", "BS")]
        for li in items_to_route:
            schedule_key = _route_schedule_key(li, log)
            if schedule_key is None:
                unmapped_items.append(li)
                continue

            if schedule_key == "Revenue":
                row = _revenue_target_row(li, state)
                # Keep the template's own "Sale of goods"/"Sale of services" label --
                # only the FY value columns are the company's data here.
                _write_values(ws, row, li, fy_columns, overwrite_label=False)
                log.append(f"Matched '{li.line_item}' -> Sch. row {row} (Revenue)")
                continue

            if schedule_key == "Share Capital":
                _write_share_capital(ws, li, fy_columns, state)
                log.append(f"Matched '{li.line_item}' -> Sch. row {state['Share Capital']['first_item_row']} (Share Capital)")
                continue

            if schedule_key == "Reserves and Surplus":
                _write_opening_reserves(ws, li, fy_columns, state, log)
                log.append(f"Matched '{li.line_item}' -> Sch. row {state['Reserves and Surplus']['first_item_row']} (Opening Reserves)")
                continue

            sched = state[schedule_key]
            matched_row = _find_matching_row(
                ws, sched["first_item_row"], sched["last_item_row"], li.line_item
            )
            if matched_row is not None:
                _write_values(ws, matched_row, li, fy_columns)
                log.append(f"Matched '{li.line_item}' -> Sch. row {matched_row} ({schedule_key})")
            elif not sched.get("insertable", True):
                # This schedule's Total is a fixed formula, not a SUM(range) -- a new
                # row here wouldn't be picked up. Fall back to the growable sibling
                # bucket instead of silently producing an incomplete Total.
                fallback_key = "COGS-Manufacturing"
                new_row = _insert_row(ws, fallback_key, state)
                _write_values(ws, new_row, li, fy_columns)
                ws.Range(f"B{new_row}:G{new_row}").Font.Name = "Ebrima"
                ws.Range(f"B{new_row}:G{new_row}").Font.Size = 10
                log.append(
                    f"Inserted '{li.line_item}' -> new Sch. row {new_row} ({fallback_key}, "
                    f"fallback from non-insertable {schedule_key})"
                )
            else:
                new_row = _insert_row(ws, schedule_key, state)
                _write_values(ws, new_row, li, fy_columns)
                ws.Range(f"B{new_row}:G{new_row}").Font.Name = "Ebrima"
                ws.Range(f"B{new_row}:G{new_row}").Font.Size = 10
                log.append(f"Inserted '{li.line_item}' -> new Sch. row {new_row} ({schedule_key})")

        _verify_totals(ws, state, log)
        _apply_equity_share_count(ws, state, equity_shares)

        if sections is not None and irl_rows is not None:
            _write_analysis_sheet_com(wb, company_name, sections, irl_rows)

        # Headless COM automation can defer recalculation past the point where the
        # cached formula results get written to the saved file -- force a full
        # recalc so IS/BS/CF/Ratios don't save with stale (zero/blank) cached values.
        excel.CalculateFullRebuild()

        save_changes = True
    finally:
        if wb is not None:
            wb.Close(SaveChanges=save_changes)
        excel.Quit()

    return log, unmapped_items
