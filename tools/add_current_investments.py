#!/usr/bin/env python
"""
Brings the Reference file's "Current investments" schedule into the shipped template
(template/commonsize_template.xlsx), as the firm added it on 2026-10-14:

    Sch. 156  Current investments          (title)
         157  Particulars                  (year headings)
         158  Current investment           (input row)
         159  Total = SUM(C158)
         160  (blank)
    BS   18   Current and non-current investments = (Sch.!C154 + Sch.!C159) / curr

Excel inserts the five rows (Windows + Excel needed, like tools/build_master_template.py), so every formula that
points below row 155 -- on Sch., IS, BS, CF and Ratios -- is renumbered by Excel itself. Every row number from 156
on in config/schedule_layout.json moves down by 5 and the "Current Investments" schedule is added.

    python tools/add_current_investments.py
"""

import json
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent.parent
TEMPLATE = APP_DIR / "template" / "commonsize_template.xlsx"
LAYOUT = APP_DIR / "config" / "schedule_layout.json"
AT, COUNT = 156, 5


def main():
    import win32com.client

    excel = win32com.client.DispatchEx("Excel.Application")   # its own instance: the user's Excel is untouched
    excel.Visible, excel.DisplayAlerts = False, False
    try:
        wb = excel.Workbooks.Open(str(TEMPLATE))
        sch, bs = wb.Worksheets("Sch."), wb.Worksheets("BS")
        if str(sch.Cells(AT, 2).Value or "").strip().lower() == "current investments":
            print("Already applied.")
            wb.Close(SaveChanges=False)
            return
        assert str(sch.Cells(145, 2).Value).strip() == "Non-current investments", sch.Cells(145, 2).Value
        sch.Rows(f"{AT}:{AT + COUNT - 1}").Insert()
        for src, dst in ((145, 156), (146, 157), (147, 158), (154, 159), (155, 160)):   # same look as the block above
            sch.Rows(src).Copy(sch.Rows(dst))
        sch.Range(f"A{AT}:R{AT + COUNT - 1}").ClearContents()
        sch.Cells(156, 2).Value = "Current investments"
        sch.Cells(157, 2).Value = "Particulars"
        sch.Cells(158, 2).Value = "Current investment"
        sch.Cells(159, 2).Value = "Total"
        for col in "CDEFG":
            sch.Range(f"{col}157").Formula = f"={col}6"
            sch.Range(f"{col}159").Formula = f"=SUM({col}158)"
            bs.Range(f"{col}18").Formula = f"=(Sch.!{col}154+Sch.!{col}159)/curr"
        for r in range(AT, AT + COUNT):
            sch.Rows(r).Hidden = False
        wb.Save()
        wb.Close()
    finally:
        excel.Quit()

    layout = json.loads(LAYOUT.read_text(encoding="utf-8"))
    for sched in layout["schedules"].values():
        for k in ("title_row", "first_item_row", "last_item_row", "total_row"):
            if sched[k] >= AT:
                sched[k] += COUNT
    layout["schedules"]["Current Investments"] = dict(
        title_row=156, first_item_row=158, last_item_row=158, total_row=159, pattern="total_last",
        insertable=False, spare_rows=0)
    order = layout["order"]
    order.insert(order.index("Investments") + 1, "Current Investments")
    LAYOUT.write_text(json.dumps(layout, indent=2) + "\n", encoding="utf-8")
    print("Template and layout updated.")


if __name__ == "__main__":
    main()
