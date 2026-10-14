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
  2. otherwise the next spare row of that schedule (hidden in the template): labelled with the PDF's own
     wording, highlighted yellow and given a comment naming the PDF heading and page;
  3. if the spare rows are used up, or the item could not be placed at all, a yellow block at the bottom of
     Sch. -- these are NOT part of any total, and the report says so (the balance check will show the gap).

Values are written in rupees; the template divides by `curr` (INR million). The output hides no row or
column: the unused spare rows show blank, and year columns without data are emptied rather than hidden.
"""

import bisect
import re
from copy import copy
from dataclasses import dataclass, field

import openpyxl
from openpyxl.comments import Comment
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.formula.translate import Translator
from openpyxl.packaging.custom import StringProperty
from openpyxl.styles import Border, PatternFill
from openpyxl.utils import column_index_from_string
from rapidfuzz import fuzz

from config.schedule_map import ALL_SCHEDULES, FY_COLUMN_MAP, SHEET_NAME, column_map, spare_row_numbers
from config.taxonomy import GENERIC_FILLERS
from core.routing import CHANGE_IN_INVENTORIES, schedule_for

YELLOW = PatternFill(fill_type="solid", start_color="FFFF00", end_color="FFFF00")
VALUE_COLUMNS = "CDEFG"
SPECIAL_SINGLE = {            # schedules that are a fixed input cell, values are added into it
    "Share Capital", "Securities Premium", "Other Reserves", "Reserves Adjustments",
    "Deferred Tax Assets", "Deferred Tax Liabilities",
}
_MATERIALS_CONSUMED = re.compile(r"materials?\s*consumed|cost\s*of\s*(?:raw\s*)?materials?", re.I)
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
    columns: dict = field(default_factory=lambda: dict(FY_COLUMN_MAP))   # FY -> template column for this workbook

    @property
    def new_rows(self):
        return [p for p in self.placements if p.kind == "spare"]


def _norm(text):
    return re.sub(r"[^a-z0-9]", "", str(text or "").lower())


def _match_row(ws, sched, wanted, used):
    """Row of the schedule whose existing label matches any of `wanted` (best candidates first), else None."""
    if any(w and re.match(r"^\W*less\b", w, re.I) for w in wanted):
        return None   # a "Less: ..." deduction is never merged into the line it is deducted from
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
        col = report.columns.get(fy)
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


def _row_labelled(ws, sched, pattern):
    for r in range(sched["first_item_row"], sched["last_item_row"] + 1):
        label = ws.cell(r, 2).value
        if label and not ws.row_dimensions[r].hidden and re.search(pattern, str(label), re.I):
            return r
    return None


def _revenue_row(ws, sched, item):
    text = item.line_item.lower()
    if re.match(r"^\W*less\b", text):
        return None   # a deduction (branch transfer, inter-branch sales ...) keeps its own row
    if re.search(r"scrap|freight|drawback|export\s*incentive|other\s*operating", text):
        row = _row_labelled(ws, sched, r"other\s*operating")
        if row:
            return row
    row = _match_row(ws, sched, [item.line_item], set())   # own label only: a loose alias must not merge lines
    if row:
        return row
    second = sched["first_item_row"] + 1
    is_service = re.fullmatch(r"\W*(?:sale\s*of\s*services?|services?\s*(?:income|revenue)|income\s*from\s*services?)\W*", text)
    if is_service and ws.cell(second, 2).value:
        return second
    if re.search(r"revenue\s*from\s*operations|^\W*(?:net\s+)?(?:sales|turnover)\b|sale\s*of\s*(?:products|goods)"
                 r"|manufactured\s*goods|sale\s*of\s*manufactured", text):
        return sched["first_item_row"]   # "Sale of goods": the template's own label is kept
    return None   # a specific revenue line (branch transfer, scrap, freight ...) gets its own row


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
    present = [fy for fy in fy_columns if fy in report.columns]
    if not present:
        return
    earliest = present[0]
    value = item.fy_values.get(earliest)
    if value is not None:
        cell = ws[f"{report.columns[earliest]}{row}"]
        cell.value = (cell.value if isinstance(cell.value, (int, float)) else 0.0) + value
    ignored = [fy for fy in present[1:] if item.fy_values.get(fy) is not None]
    if ignored:
        report.log.append("NOTE: opening reserves for " + ", ".join(ignored) + " were not written; the template carries "
                          "each year's closing reserves forward, only the earliest year's opening balance is an input.")


def _write_change_in_inventories(ws, items, fy_columns, columns):
    opening = ALL_SCHEDULES["Opening Stock (Change in Inventories)"]["first_item_row"]
    closing = ALL_SCHEDULES["Closing Stock (Change in Inventories)"]["first_item_row"]
    for fy in fy_columns:
        col = columns.get(fy)
        values = [i.fy_values.get(fy) for i in items if i.fy_values.get(fy) is not None]
        if not col or not values:
            continue
        net = sum(values)                      # template: change = opening - closing
        levels = [i.stock_levels[fy] for i in items if fy in i.stock_levels]
        if len(levels) == 1 and abs(levels[0][0] - levels[0][1] - net) < 1.0:
            ws[f"{col}{opening}"].value, ws[f"{col}{closing}"].value = levels[0]    # the stock the notes give
            continue
        ws[f"{col}{opening}"].value = max(net, 0.0)
        ws[f"{col}{closing}"].value = max(-net, 0.0)


_HEADER_REF = re.compile(r"^=\$?[A-Z]{1,2}\$?6$")      # a schedule's year heading ("=D6")


def _clear_column(ws, letter, first_row):
    """Empties the figures and formulas of one year column, keeping labels and the schedules' year headings."""
    col = column_index_from_string(letter)
    for row in range(first_row, ws.max_row + 1):
        cell = ws.cell(row, col)
        v = cell.value
        if v is None or (isinstance(v, str) and (not v.startswith("=") or _HEADER_REF.match(v))):
            continue
        cell.value = None


def _clear_unused_years(wb, fy_columns, columns):
    """
    Nothing in the workbook is hidden, so a year column without data is emptied instead: the template carries
    reserves forward into every year column, and a year with no figures would otherwise show a spurious
    balance-sheet check and ratios. A cash flow needs the previous year's balance sheet, so the earliest year's
    cash-flow column is emptied too. The year headings stay; a later "add a new FY" rebuilds the workbook from
    the template, so the formulas come back for the year added.
    """
    years = list(columns)                            # the five years of columns C..G
    for i, fy in enumerate(years):
        has_data = fy in fy_columns
        has_prior = has_data and i > 0 and years[i - 1] in fy_columns
        if not has_data:
            for sheet, letters, first_row in (("IS", ["CDEFG"[i], "HIJKL"[i]], 6), ("BS", ["CDEFG"[i]], 6),
                                              ("Ratios", ["DEFGH"[i]], 4),
                                              (SHEET_NAME, ["CDEFG"[i], "IJKLM"[i]] + (["OPQR"[i - 1]] if i else []), 7)):
                for letter in letters:
                    _clear_column(wb[sheet], letter, first_row)
        if i > 0 and not has_prior:
            _clear_column(wb["CF"], "CDEFG"[i - 1], 6)


def _show_everything(wb):
    """The output hides no sheet, row or column (the template keeps its spare schedule rows hidden)."""
    for ws in wb.worksheets:
        ws.sheet_state = "visible"
        for dim in ws.row_dimensions.values():
            dim.hidden = False
        for dim in ws.column_dimensions.values():
            dim.hidden = False


REMOVED_ROWS_PROPERTY = "commonsize_removed_sch_rows"     # template rows of Sch. deleted from this workbook
_REF = re.compile(r"(?P<sheet>(?:'[^']+'|[^'!]+)!)?(?P<c1>\$?[A-Z]{1,3})(?P<d1>\$?)(?P<r1>\d+)"
                  r"(?::(?P<c2>\$?[A-Z]{1,3})(?P<d2>\$?)(?P<r2>\d+))?")


def _clean_borders(ws):
    """Only the 'Particulars' heading rows and the 'Total' rows keep their borders; item rows (including the
    spare rows new items go into) and headings in between carry none."""
    plain = Border()
    for row in ws.iter_rows():
        label = str(ws.cell(row[0].row, 2).value or "").strip().lower()
        if label == "particulars" or label.startswith("total"):
            # the same line under the label as under the figures (the template is uneven here)
            ws.cell(row[0].row, 2).border = copy(ws.cell(row[0].row, 3).border)
            continue
        for cell in row:
            if cell.has_style and cell.border != plain:
                cell.border = plain


def _compact_ranges(rows):
    out, rows = [], sorted(rows)
    for r in rows:
        if out and r == out[-1][1] + 1:
            out[-1][1] = r
        else:
            out.append([r, r])
    return ",".join(f"{a}-{b}" if a != b else f"{a}" for a, b in out)


def expand_ranges(text):
    rows = []
    for part in filter(None, (text or "").split(",")):
        a, _, b = part.partition("-")
        rows += range(int(a), int(b or a) + 1)
    return rows


def _delete_rows(wb, ws, removed):
    """
    Deletes rows of `ws` and renumbers every reference to that sheet in every formula of the workbook (openpyxl's
    own delete_rows moves cells but leaves formulas pointing at the old rows). A range whose first or last row is
    deleted shrinks to the nearest row kept.
    """
    removed = sorted(set(removed))
    gone = set(removed)

    def renumber(r, end=False):
        while r in gone:
            r += -1 if end else 1
        return r - bisect.bisect_left(removed, r)

    def fix_ref(text, sheet_title):
        m = _REF.fullmatch(text)
        if not m:
            return text
        sheet = (m.group("sheet") or "")[:-1].strip("'") or sheet_title
        if sheet != ws.title:
            return text
        prefix, r1 = m.group("sheet") or "", renumber(int(m.group("r1")))
        if not m.group("r2"):
            return f"{prefix}{m.group('c1')}{m.group('d1')}{r1}"
        r2 = max(r1, renumber(int(m.group("r2")), end=True))
        return f"{prefix}{m.group('c1')}{m.group('d1')}{r1}:{m.group('c2')}{m.group('d2')}{r2}"

    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                v = cell.value
                if not (isinstance(v, str) and v.startswith("=")):
                    continue
                tok = Tokenizer(v)
                changed = False
                for t in tok.items:
                    if t.type == "OPERAND" and t.subtype == "RANGE":
                        fixed = fix_ref(t.value, sheet.title)
                        if fixed != t.value:
                            t.value, changed = fixed, True
                if changed:
                    cell.value = tok.render()

    dims = {r: d for r, d in ws.row_dimensions.items()}
    # delete from the bottom up, in contiguous blocks
    blocks = []
    for r in reversed(removed):
        if blocks and blocks[-1][0] == r + 1:
            blocks[-1][0] = r
        else:
            blocks.append([r, r])
    for first, last in blocks:
        ws.delete_rows(first, last - first + 1)
    ws.row_dimensions.clear()
    for r, d in dims.items():
        if r in gone:
            continue
        d.index = renumber(r)
        ws.row_dimensions[d.index] = d


def _remove_empty_rows(wb, ws):
    """Rows of Sch. with no heading, label or figure (unused spare rows, spacing rows) are deleted; which template
    rows were deleted is kept in a workbook property so 'add a new FY' can read the workbook back."""
    removed = [r for r in range(1, ws.max_row + 1) if all(ws.cell(r, c).value is None for c in range(1, 8))]
    if not removed:
        return
    _delete_rows(wb, ws, removed)
    wb.custom_doc_props.append(StringProperty(name=REMOVED_ROWS_PROPERTY, value=_compact_ranges(removed)))


def _relabel_years(wb, columns):
    """Year headings for a moved window: IS / Sch. 'FY22' -> 'FY19', BS '31 Mar 22' -> '31 Mar 19' (CF and Ratios
    take theirs from these)."""
    for fy, col in columns.items():
        yy = fy[2:]
        wb["IS"][f"{col}5"].value = fy
        wb[SHEET_NAME][f"{col}6"].value = fy
        wb["BS"][f"{col}5"].value = f"31 Mar {yy}"


def _cash_flow_share_capital(wb, ws, fy_columns, report):
    """
    The template's cash flow finances only 'Changes in reserves' (reserves less profit); when share capital
    changes between the years written (shares issued, preference shares redeemed) that line takes the change
    in net worth instead, so the cash-flow check still closes.
    """
    row = ALL_SCHEDULES["Share Capital"]["first_item_row"]
    capital = {fy: ws[f"{report.columns[fy]}{row}"].value or 0.0 for fy in fy_columns if fy in report.columns}
    years = [fy for fy in report.columns if fy in capital]
    if all(abs(capital[a] - capital[b]) < 0.5 for a, b in zip(years, years[1:])):
        return
    cf, bs = wb["CF"], wb["BS"]
    reserves_row = next((r for r in range(1, 40) if re.match(r"reserves\s*and\s*surplus", str(bs.cell(r, 2).value or ""), re.I)), None)
    net_worth_row = next((r for r in range(1, 40) if re.match(r"net\s*worth", str(bs.cell(r, 2).value or ""), re.I)), None)
    for r in range(1, 60):
        if re.match(r"changes\s*in\s*reserves", str(cf.cell(r, 2).value or ""), re.I) and reserves_row and net_worth_row:
            cf.cell(r, 2).value = "Changes in share capital and reserves"
            for c in range(3, 8):
                cell = cf.cell(r, c)
                if isinstance(cell.value, str):
                    cell.value = re.sub(rf"(BS!\$?[A-Z]+\$?){reserves_row}\b", rf"\g<1>{net_worth_row}", cell.value)
            report.log.append("NOTE: share capital changes between the years, so the cash flow's 'Changes in reserves' "
                              "line also includes the change in share capital.")
            return


def write_workbook(template_path, output_path, company_name, line_items, fy_columns, equity_shares=None,
                   reported_pat=None):
    """
    Populate a copy of the template. fy_columns ascending. Returns a WriteReport.
    reported_pat {fy: rupees}: the profit the financials themselves show, written to IS 'As per financials'
    so the template's own 'Difference' row reconciles the computed PAT with it.
    """
    report = WriteReport()
    fy_columns = sorted(fy_columns, key=lambda f: int(f[2:]))
    report.columns = column_map(fy_columns)
    for fy in fy_columns:
        if fy not in report.columns:
            report.log.append(f"WARNING: {fy} is outside the workbook's five year columns ({', '.join(report.columns)}); not written.")

    wb = openpyxl.load_workbook(template_path)
    ws = wb[SHEET_NAME]
    wb["IS"]["B2"].value = company_name

    used_spares, inventory_items, overflow, extras, filled = set(), [], [], {}, {}
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
        if key == "COGS-Purchase":
            # two fixed rows: purchase of traded goods, then cost of materials consumed (an XBRL filing's
            # "Cost of materials consumed" is its purchases); anything else is a direct expense
            sched = ALL_SCHEDULES[key]
            if _MATERIALS_CONSUMED.search(item.line_item):
                row = sched["last_item_row"]
            elif re.search(r"purchase", item.line_item, re.I):
                row = sched["first_item_row"]
            else:
                row, key = None, "COGS-Manufacturing"
            if row is not None:
                _add_values(ws, row, item, fy_columns, report)
                report.placements.append(Placement(item.line_item, key, row, "special"))
                continue

        sched = ALL_SCHEDULES[key]
        if key in SPECIAL_SINGLE:
            row = sched["first_item_row"]
            _add_values(ws, row, item, fy_columns, report)
            if "label_row" in item.flags:
                ws.cell(row, 2).value = item.line_item       # the filing's own wording for a fixed row
            if item.pdf_heading.startswith("Authorised: ") and key == "Share Capital":
                ws.cell(ALL_SCHEDULES[key]["title_row"] + 3, 2).value = item.pdf_heading[len("Authorised: "):]
            report.placements.append(Placement(item.line_item, key, row, "special"))
            continue

        # a taxonomy match names the row; a heading / alias match only gives a fallback row, so the item's own
        # label ("Staff Welfare" -> "Workers and staff welfare") is tried first
        wanted = [item.matched_item, item.line_item]
        if not item.match_score and _norm(item.line_item) not in GENERIC_FILLERS:
            wanted.reverse()
        row = _revenue_row(ws, sched, item) if key == "Revenue" else _match_row(ws, sched, wanted, used_spares)
        if row is not None and "own_row" in item.flags and _norm(filled.get(row, item.line_item)) != _norm(item.line_item):
            # a different line already sits in that row: keep each line of the filing on its own row
            spare = _place_in_spare(ws, key, item, fy_columns, report, used_spares)
            if spare is not None:
                filled[spare] = item.line_item
                report.placements.append(Placement(item.line_item, key, spare, "spare", True))
                continue
        if row is not None:
            if "own_row" in item.flags and row not in filled and _match_row(ws, {"first_item_row": row, "last_item_row": row},
                                                                           [item.line_item], set()) is None:
                ws.cell(row, 2).value = item.line_item     # a row taken by meaning only: the filing's own wording
            filled.setdefault(row, item.line_item)
            _add_values(ws, row, item, fy_columns, report)
            report.placements.append(Placement(item.line_item, key, row, "matched"))
            continue
        row = _place_in_spare(ws, key, item, fy_columns, report, used_spares)
        if row is not None:
            filled[row] = item.line_item
            report.placements.append(Placement(item.line_item, key, row, "spare", True))
            continue
        total_cell = ws[f"C{sched['total_row']}"].value
        if sched.get("insertable", True) and isinstance(total_cell, str) and total_cell.startswith("=SUM("):
            extras.setdefault(key, []).append(item)
        else:
            overflow.append((item, f"no free row left in '{key}'"))
            report.placements.append(Placement(item.line_item, key, 0, "unplaced", True))

    if inventory_items:
        _write_change_in_inventories(ws, inventory_items, fy_columns, report.columns)
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
        col = report.columns.get(fy)
        if col and value is not None and fy in fy_columns:
            # the full figure over the workbook's unit cell, so it follows a change of unit like every other line
            number = f"{value:.2f}".rstrip("0").rstrip(".")
            wb["IS"][f"{col}24"].value = f"={number if number not in ('', '-', '-0') else '0'}/curr"

    if equity_shares:
        ws[f"B{ALL_SCHEDULES['Share Capital']['first_item_row']}"].value = f"{equity_shares} equity shares of Rs. 10 each"

    if report.columns != FY_COLUMN_MAP:
        _relabel_years(wb, report.columns)
    _cash_flow_share_capital(wb, ws, fy_columns, report)
    _clear_unused_years(wb, fy_columns, report.columns)
    _show_everything(wb)
    _clean_borders(ws)
    _remove_empty_rows(wb, ws)              # last: every row number above refers to the template's layout
    wb.calculation.fullCalcOnLoad = True   # openpyxl drops cached results; Excel recomputes on open
    wb.save(output_path)
    return report
