"""
Fills the Reference file with the classified line items using openpyxl only
(no Excel/COM), so it runs on a Linux host such as Render.

Why this is safe: the template already has every row the taxonomy needs plus
hidden spare rows inside each growable schedule's SUM range. The writer never
inserts or deletes a row; it only writes values into input cells that exist, so
every formula in IS / BS / CF / Ratios stays untouched and Excel recalculates
them on opening (`fullCalcOnLoad`).

Where an item goes
  1. an existing row of its schedule whose label matches (taxonomy / alias name, then the PDF label);
  2. otherwise the next hidden spare row of that schedule: the row is unhidden, labelled with the PDF's own
     wording, highlighted yellow and given a comment naming the PDF heading and page;
  3. if the spare rows are used up, or the item could not be placed at all, a yellow block at the bottom of
     Sch. -- these are NOT part of any total, and the report says so (the balance check will show the gap).

Values are written in rupees; the template divides by `curr` (INR million).
"""

import re
from copy import copy
from dataclasses import dataclass, field

import openpyxl
from openpyxl.comments import Comment
from openpyxl.formula.translate import Translator
from openpyxl.styles import PatternFill
from openpyxl.utils import column_index_from_string, get_column_letter
from rapidfuzz import fuzz

from config.schedule_map import ALL_SCHEDULES, FY_COLUMN_MAP, SHEET_NAME, spare_row_numbers
from config.taxonomy import GENERIC_FILLERS
from core.routing import CHANGE_IN_INVENTORIES, schedule_for

YELLOW = PatternFill(fill_type="solid", start_color="FFFF00", end_color="FFFF00")
VALUE_COLUMNS = "CDEFG"
SPECIAL_SINGLE = {            # schedules that are a fixed input cell, values are added into it
    "Share Capital", "Securities Premium", "Reserves Adjustments",
    "Deferred Tax Assets", "Deferred Tax Liabilities", "COGS-Purchase",
}
REVENUE_SERVICE_WORDS = ("service", "commission", "consult", "fees", "job work", "labour charges")


@dataclass
class Placement:
    label: str
    schedule: str
    row: int
    kind: str                    # matched | spare | special | unplaced
    is_new: bool = False


@dataclass
class WriteReport:
    log: list = field(default_factory=list)
    placements: list = field(default_factory=list)
    unplaced: list = field(default_factory=list)      # items written to the bottom block (not in totals)

    @property
    def new_rows(self):
        return [p for p in self.placements if p.kind == "spare"]


def _norm(text):
    return re.sub(r"[^a-z0-9]", "", str(text or "").lower())


def _match_row(ws, sched, wanted, used):
    """Row of the schedule whose existing label matches any of `wanted` (best candidates first), else None."""
    rows = []
    for r in range(sched["first_item_row"], sched["last_item_row"] + 1):
        label = ws.cell(r, 2).value
        if label and not ws.row_dimensions[r].hidden and len(_norm(label)) >= 3:
            rows.append((r, _norm(label), str(label)))
    for name in (w for w in wanted if w):
        n = _norm(name)
        if not n:
            continue
        for r, nl, _ in rows:
            if nl == n:
                return r
        if n in GENERIC_FILLERS:
            for r, nl, _ in rows:
                if nl.startswith("other"):
                    return r
            continue
        for r, nl, _ in rows:
            short, long_ = sorted((n, nl), key=len)
            if len(short) >= 6 and short in long_:
                return r
        best = max(rows, key=lambda x: fuzz.token_sort_ratio(name.lower(), x[2].lower()), default=None)
        if best and fuzz.token_sort_ratio(name.lower(), best[2].lower()) >= 88:
            return best[0]
    return None


def _add_values(ws, row, item, fy_columns, report):
    for fy in fy_columns:
        col = FY_COLUMN_MAP.get(fy)
        value = item.fy_values.get(fy)
        if not col or value is None:
            continue
        cell = ws[f"{col}{row}"]
        if isinstance(cell.value, str) and cell.value.startswith("="):
            report.log.append(f"NOTE: {col}{row} holds a formula, so '{item.line_item}' ({fy}) was not written there.")
            continue
        cell.value = (cell.value or 0.0) + value


def _clone_row_formulas(ws, source_row, target_row):
    """Spare rows were emptied when the template was built; restore the % / YoY formulas from a live row."""
    for cell in ws[source_row]:
        if cell.column <= 7 or not (isinstance(cell.value, str) and cell.value.startswith("=")):
            continue
        target = ws.cell(target_row, cell.column)
        if target.value is None:
            target.value = Translator(cell.value, origin=cell.coordinate).translate_formula(target.coordinate)


def _comment(item):
    parts = [f"New line item (not in the taxonomy). PDF heading: {item.pdf_heading or 'n/a'}"]
    if item.pdf_page:
        parts.append(f"Page {item.pdf_page}")
    return Comment("; ".join(parts), "Commonsize app")


def _revenue_row(ws, sched, item):
    names = [item.matched_item, item.line_item]
    row = _match_row(ws, sched, names, set())
    if row:
        return row
    text = item.line_item.lower()
    second = sched["first_item_row"] + 1
    if any(w in text for w in REVENUE_SERVICE_WORDS) and ws.cell(second, 2).value:
        return second
    return sched["first_item_row"]   # "Sale of goods": the template's own label is kept


def _place_in_spare(ws, key, item, fy_columns, report, used_spares):
    sched = ALL_SCHEDULES[key]
    free = [r for r in spare_row_numbers(key) if r not in used_spares]
    if not free:
        return None
    row = free[0]
    used_spares.add(row)
    ws.row_dimensions[row].hidden = False
    ws.cell(row, 2).value = item.line_item
    for col in "BCDEFG":
        ws[f"{col}{row}"].fill = YELLOW
    ws[f"B{row}"].comment = _comment(item)
    _clone_row_formulas(ws, sched["first_item_row"], row)
    _add_values(ws, row, item, fy_columns, report)
    return row


def _write_schedule_extras(ws, extras, fy_columns, report):
    """
    Items that found no free spare row still count: they are listed in a block at the bottom of Sch. and that
    block is added to the schedule's Total formula, so every statement picks them up without inserting a row.
    """
    row = ws.max_row + 3
    for key, items in extras.items():
        sched = ALL_SCHEDULES[key]
        ws.cell(row, 2).value = f"Additional items - {key}"
        for col in "BCDEFG":
            ws[f"{col}{row}"].fill = YELLOW
        first = row + 1
        for item in items:
            row += 1
            ws.cell(row, 2).value = item.line_item
            for col in "BCDEFG":
                ws[f"{col}{row}"].fill = YELLOW
            ws[f"B{row}"].comment = _comment(item)
            _clone_row_formulas(ws, sched["first_item_row"], row)
            _add_values(ws, row, item, fy_columns, report)
            report.placements.append(Placement(item.line_item, key, row, "extra", True))
        for col in VALUE_COLUMNS:
            cell = ws[f"{col}{sched['total_row']}"]
            if isinstance(cell.value, str) and cell.value.startswith("="):
                cell.value = f"{cell.value}+SUM({col}{first}:{col}{row})"
        row += 3


def _write_overflow_block(ws, items, fy_columns, report):
    start = ws.max_row + 3
    ws.cell(start, 2).value = "Items NOT placed in any schedule (not included in the statements above)"
    for col in "BCDEFG":
        ws[f"{col}{start}"].fill = YELLOW
    row = start + 1
    for item, reason in items:
        ws.cell(row, 2).value = item.line_item
        ws.cell(row, 8).value = reason
        for col in "BCDEFG":
            ws[f"{col}{row}"].fill = YELLOW
        _add_values(ws, row, item, fy_columns, report)
        row += 1


def _write_opening_reserves(ws, item, fy_columns, report):
    sched = ALL_SCHEDULES["Reserves and Surplus"]
    row = sched["first_item_row"]
    present = [fy for fy in fy_columns if fy in FY_COLUMN_MAP]
    if not present:
        return
    earliest = present[0]
    value = item.fy_values.get(earliest)
    if value is not None:
        cell = ws[f"{FY_COLUMN_MAP[earliest]}{row}"]
        cell.value = (cell.value if isinstance(cell.value, (int, float)) else 0.0) + value
    ignored = [fy for fy in present[1:] if item.fy_values.get(fy) is not None]
    if ignored:
        report.log.append("NOTE: opening reserves for " + ", ".join(ignored) + " were not written; the template carries "
                          "each year's closing reserves forward, only the earliest year's opening balance is an input.")


def _write_change_in_inventories(ws, items, fy_columns):
    opening = ALL_SCHEDULES["Opening Stock (Change in Inventories)"]["first_item_row"]
    closing = ALL_SCHEDULES["Closing Stock (Change in Inventories)"]["first_item_row"]
    for fy in fy_columns:
        col = FY_COLUMN_MAP.get(fy)
        values = [i.fy_values.get(fy) for i in items if i.fy_values.get(fy) is not None]
        if not col or not values:
            continue
        net = sum(values)                      # template: change = opening - closing
        ws[f"{col}{opening}"].value = max(net, 0.0)
        ws[f"{col}{closing}"].value = max(-net, 0.0)


def _split_column_groups(ws, upto=40):
    """The template stores column widths as ranges (e.g. C:F); split them so one column can be hidden alone."""
    singles = {d.min for d in ws.column_dimensions.values() if (d.min or 0) == (d.max or 0)}
    for key, dim in list(ws.column_dimensions.items()):
        lo = dim.min or column_index_from_string(key)
        hi = dim.max or lo
        if hi == lo:
            continue
        template = copy(dim)
        for i in range(lo, min(hi, upto) + 1):
            letter = get_column_letter(i)
            if i != lo and i in singles:
                continue
            piece = copy(template)
            piece.index, piece.min, piece.max = letter, i, i
            ws.column_dimensions[letter] = piece
        if hi > upto:
            rest = copy(template)
            rest.index, rest.min, rest.max = get_column_letter(upto + 1), upto + 1, hi
            ws.column_dimensions[get_column_letter(upto + 1)] = rest


def _hide_unused_years(wb, fy_columns):
    """
    The template carries reserves forward into every year column, so years with no data would show a spurious
    non-zero balance-sheet check. Years without data are hidden (unhide them, or add the year with the
    'add a new FY' option). A cash flow needs the previous year's balance sheet, so the earliest year's
    cash-flow column is hidden too.
    """
    for name in ("IS", "BS", "CF", "Ratios"):
        _split_column_groups(wb[name])
    letters = "CDEFG"
    years = list(FY_COLUMN_MAP)                      # FY22..FY26, aligned with C..G
    data = [fy for fy in years if fy in fy_columns]
    for i, fy in enumerate(years):
        has_data = fy in data
        has_prior = has_data and i > 0 and years[i - 1] in data
        col = letters[i]
        for sheet, cols in (("IS", [col, "HIJKL"[i]]), ("BS", [col]), ("Ratios", ["DEFGH"[i]])):
            for c in cols:
                wb[sheet].column_dimensions[c].hidden = not has_data
        if i > 0:
            wb["CF"].column_dimensions[letters[i - 1]].hidden = not has_prior


def write_workbook(template_path, output_path, company_name, line_items, fy_columns, equity_shares=None,
                   reported_pat=None):
    """
    Populate a copy of the template. fy_columns ascending. Returns a WriteReport.
    reported_pat {fy: rupees}: the profit the financials themselves show, written to IS 'As per financials'
    so the template's own 'Difference' row reconciles the computed PAT with it.
    """
    report = WriteReport()
    fy_columns = sorted(fy_columns)
    for fy in fy_columns:
        if fy not in FY_COLUMN_MAP:
            report.log.append(f"WARNING: {fy} is outside the template's columns ({', '.join(FY_COLUMN_MAP)}); not written.")

    wb = openpyxl.load_workbook(template_path)
    ws = wb[SHEET_NAME]
    wb["IS"]["B2"].value = company_name

    used_spares, inventory_items, overflow, extras = set(), [], [], {}
    for item in line_items:
        if item.statement not in ("P&L", "BS"):
            continue
        key = schedule_for(item)
        if key is None:
            if any(v for v in item.fy_values.values() if v):
                overflow.append((item, "could not be placed under any schedule"))
                report.placements.append(Placement(item.line_item, "-", 0, "unplaced", True))
            continue
        if key == CHANGE_IN_INVENTORIES:
            inventory_items.append(item)
            report.placements.append(Placement(item.line_item, "Change in Inventories", 0, "special"))
            continue
        if key == "Reserves and Surplus":
            _write_opening_reserves(ws, item, fy_columns, report)
            report.placements.append(Placement(item.line_item, key, ALL_SCHEDULES[key]["first_item_row"], "special"))
            continue
        if key == "COGS-Purchase" and not re.search(r"purchase", item.line_item, re.I):
            key = "COGS-Manufacturing"   # the purchase schedule has a single fixed row

        sched = ALL_SCHEDULES[key]
        if key in SPECIAL_SINGLE:
            row = sched["first_item_row"]
            _add_values(ws, row, item, fy_columns, report)
            report.placements.append(Placement(item.line_item, key, row, "special"))
            continue

        row = _revenue_row(ws, sched, item) if key == "Revenue" else _match_row(
            ws, sched, [item.matched_item, item.line_item], used_spares)
        if row is not None:
            _add_values(ws, row, item, fy_columns, report)
            report.placements.append(Placement(item.line_item, key, row, "matched"))
            continue
        row = _place_in_spare(ws, key, item, fy_columns, report, used_spares)
        if row is not None:
            report.placements.append(Placement(item.line_item, key, row, "spare", True))
            continue
        total_cell = ws[f"C{sched['total_row']}"].value
        if sched.get("insertable", True) and isinstance(total_cell, str) and total_cell.startswith("=SUM("):
            extras.setdefault(key, []).append(item)
        else:
            overflow.append((item, f"no free row left in '{key}'"))
            report.placements.append(Placement(item.line_item, key, 0, "unplaced", True))

    if inventory_items:
        _write_change_in_inventories(ws, inventory_items, fy_columns)
    if extras:
        _write_schedule_extras(ws, extras, fy_columns, report)
        report.log.append(f"NOTE: {sum(len(v) for v in extras.values())} new item(s) did not fit in the template's spare rows; "
                          f"they are listed under 'Additional items' at the bottom of '{SHEET_NAME}' and are included in the totals.")
    if overflow:
        _write_overflow_block(ws, overflow, fy_columns, report)
        report.unplaced = [i for i, _ in overflow]
        report.log.append(f"WARNING: {len(overflow)} item(s) are listed at the bottom of '{SHEET_NAME}' and are not in any "
                          "total; the balance check will show their effect.")

    for fy, value in (reported_pat or {}).items():
        col = FY_COLUMN_MAP.get(fy)
        if col and value is not None and fy in fy_columns:
            wb["IS"][f"{col}24"].value = value / 1_000_000   # IS is in INR million

    if equity_shares:
        ws[f"B{ALL_SCHEDULES['Share Capital']['first_item_row']}"].value = f"{equity_shares} equity shares of Rs. 10 each"

    _hide_unused_years(wb, fy_columns)
    wb.calculation.fullCalcOnLoad = True   # openpyxl drops cached results; Excel recomputes on open
    wb.save(output_path)
    return report
