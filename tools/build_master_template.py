#!/usr/bin/env python
"""
One-time, LOCAL-ONLY build step. Requires Microsoft Excel on Windows -- never
run by the deployed web app.

Rebuilds `Reference file/Common size format_Mfg-Trading firm.xlsx` IN PLACE
from the ORIGINAL (pre-expansion) file, so the result is always the same no
matter how many times it is run:

  1. the rows identified by cross-referencing the taxonomy in
     `Web app/CS format dropdown.xlsx` against the existing schedules
     (tools/base_layout.NEW_ROWS), and
  2. a few HIDDEN spare rows inside each growable schedule
     (tools/base_layout.SPARE_ROWS). A line item that is not in the taxonomy
     is later written into a spare row (unhidden + highlighted yellow), so it
     sits inside the schedule's SUM() range and flows into every total and
     statement without the deployed app ever having to insert a row.

Rows are inserted with Excel itself (Rows.Copy() + Insert()), so Excel
re-points every formula across IS/BS/CF/Sch./Ratios exactly as if a person
had inserted the row by hand. The final row layout is then written to
config/schedule_layout.json -- generated, never hand-transcribed.

The original file is never modified: it is kept as the `*.backup-*.xlsx`
alongside the Reference file and used as the source for every rebuild.

Run from the commonsize_app directory:
    python tools/build_master_template.py
    python tools/build_master_template.py --source "<path to original>.xlsx"
"""

import argparse
import copy
import json
import shutil
import sys
from datetime import datetime
from pathlib import Path

TOOLS_DIR = Path(__file__).resolve().parent
APP_DIR = TOOLS_DIR.parent
sys.path.insert(0, str(TOOLS_DIR))

import win32com.client as win32

from base_layout import BASE_ORDER, BASE_SCHEDULES, NEW_ROWS, SPARE_ROWS

SHEET_NAME = "Sch."
REFERENCE_DIR = APP_DIR.parent / "Reference file"
TEMPLATE_PATH = REFERENCE_DIR / "Common size format_Mfg-Trading firm.xlsx"
LAYOUT_PATH = APP_DIR / "config" / "schedule_layout.json"
XL_SHIFT_DOWN = -4121
ROW_FIELDS = ("title_row", "first_item_row", "last_item_row", "total_row")


def find_original_source():
    backups = sorted(REFERENCE_DIR.glob("*backup-*.xlsx"))
    if not backups:
        raise FileNotFoundError(
            f"No original backup (*backup-*.xlsx) found in {REFERENCE_DIR}; pass --source explicitly."
        )
    return backups[0]  # oldest = the true original


def shift_after(state, order, key, delta=1):
    for later in order[order.index(key) + 1:]:
        for field in ROW_FIELDS:
            state[later][field] += delta


def insert_row(ws, state, order, key):
    """Insert one row at the schedule's last item row (inside its SUM range); returns the new row."""
    sched = state[key]
    target, source = sched["last_item_row"], sched["first_item_row"]
    ws.Rows(source).Copy()
    ws.Rows(target).Insert(Shift=XL_SHIFT_DOWN)
    ws.Application.CutCopyMode = False
    sched["last_item_row"] += 1
    if sched["pattern"] == "total_last":
        sched["total_row"] += 1
    shift_after(state, order, key)
    return target


def verify_and_repair(ws, state, log):
    """
    Excel extends a SUM range when a row is inserted strictly inside it, but not when the range is a
    single row (insertion at its boundary shifts the range down instead). Rewrite any total that does
    not span its schedule, then fail loudly if one still doesn't.
    """
    problems = []
    for key, s in state.items():
        if s["pattern"] == "single_cell" or not s["insertable"]:
            continue
        for col in "CDEFG":
            expected = f"=SUM({col}{s['first_item_row']}:{col}{s['last_item_row']})"
            cell = ws.Range(f"{col}{s['total_row']}")
            if str(cell.Formula) != expected:
                if col == "C":
                    log.append(f"~ repaired total of {key}: {cell.Formula!r} -> {expected}")
                cell.Formula = expected
        check = str(ws.Range(f"C{s['total_row']}").Formula)
        if check != f"=SUM(C{s['first_item_row']}:C{s['last_item_row']})":
            problems.append(f"{key}: total row {s['total_row']} is {check!r}")
    return problems


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", help="original (pre-expansion) workbook; default = oldest backup")
    args = parser.parse_args()

    source = Path(args.source) if args.source else find_original_source()
    print(f"Source (original, never modified): {source}")
    shutil.copyfile(source, TEMPLATE_PATH)

    state = copy.deepcopy(BASE_SCHEDULES)
    order = list(BASE_ORDER)
    log = []

    # DispatchEx forces a brand-new, separate Excel process: Dispatch can attach to an Excel the
    # user already has open, and hiding/quitting that would disrupt their work.
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb, saved = None, False
    try:
        wb = excel.Workbooks.Open(str(TEMPLATE_PATH.resolve()))
        ws = wb.Sheets(SHEET_NAME)

        for key in order:
            for label in NEW_ROWS.get(key, []):
                row = insert_row(ws, state, order, key)
                ws.Range(f"B{row}").Value = label
                log.append(f"+ {label!r} -> row {row} ({key})")
            for _ in range(SPARE_ROWS.get(key, 0)):
                row = insert_row(ws, state, order, key)
                ws.Range(f"B{row}").ClearContents()
                ws.Rows(row).Hidden = True
            state[key]["spare_rows"] = SPARE_ROWS.get(key, 0)

        problems = verify_and_repair(ws, state, log)
        if problems:
            raise RuntimeError("Totals do not span their schedules:\n  " + "\n  ".join(problems))

        excel.CalculateFullRebuild()
        wb.Save()
        saved = True
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        excel.Quit()

    layout = {
        "sheet": SHEET_NAME,
        "built": datetime.now().isoformat(timespec="seconds"),
        "source": source.name,
        "order": order,
        "schedules": state,
    }
    LAYOUT_PATH.write_text(json.dumps(layout, indent=2), encoding="utf-8")
    # the deployed app ships its own copy of the rebuilt template (the git repo is this folder only)
    deploy_copy = APP_DIR / "template" / "commonsize_template.xlsx"
    deploy_copy.parent.mkdir(exist_ok=True)
    shutil.copyfile(TEMPLATE_PATH, deploy_copy)

    print("\n".join(log))
    print(f"\nRebuilt in place: {TEMPLATE_PATH}  (saved={saved})")
    print(f"Layout written:   {LAYOUT_PATH}")
    total_new = sum(len(v) for v in NEW_ROWS.values())
    total_spare = sum(SPARE_ROWS.values())
    print(f"{total_new} taxonomy rows + {total_spare} hidden spare rows added.")


if __name__ == "__main__":
    main()
