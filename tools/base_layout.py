"""
Schedule layout of the ORIGINAL Reference file (before any Phase 4 expansion),
used only by tools/build_master_template.py as the starting state it applies
insertions to. Row numbers here are the original file's; the build tool
tracks every shift itself and writes the FINAL layout to
config/schedule_layout.json, which is what the app actually reads.

Patterns:
  total_last   item rows, then a Total row (SUM(first:last)) below them
  total_first  a combined title+Total row, then the item rows below it
  single_cell  a fixed input cell, not a growable list
"""


def _s(title, first, last, total, pattern="total_last", insertable=True):
    return {
        "title_row": title,
        "first_item_row": first,
        "last_item_row": last,
        "total_row": total,
        "pattern": pattern,
        "insertable": insertable,
        "spare_rows": 0,
    }


BASE_SCHEDULES = {
    "Share Capital": _s(5, 10, 10, 11, "single_cell", False),
    "Reserves and Surplus": _s(13, 17, 17, 22, "single_cell", False),         # Opening balance
    "Securities Premium": _s(13, 15, 15, 22, "single_cell", False),
    "Reserves Adjustments": _s(13, 19, 19, 22, "single_cell", False),         # "Additions during the year"
    "Long-term Provisions": _s(24, 26, 27, 28),
    "Borrowings - Long Term": _s(30, 32, 33, 34),
    "Trade Payables": _s(36, 38, 40, 41),
    "Borrowings - Short Term": _s(43, 45, 48, 49),
    "Short-term Provisions": _s(51, 53, 54, 55),
    "Current Liability - Other": _s(57, 59, 65, 66),
    "Fixed Assets": _s(68, 71, 76, 77),   # row 70 is a non-summed sub-header
    "Long-term Loans & Advances": _s(79, 81, 81, 82),
    "Investments": _s(84, 86, 87, 88),
    "Trade Receivables": _s(90, 92, 94, 95),
    "Inventory": _s(97, 99, 102, 103),
    "Cash & Bank": _s(105, 107, 109, 110),
    "Deferred Tax Assets": _s(112, 114, 114, 116, "total_last", False),   # Total = DTA - DTL
    "Deferred Tax Liabilities": _s(112, 115, 115, 116, "single_cell", False),
    "Short-term Loans & Advances": _s(118, 120, 122, 123),
    "Current Asset - Other": _s(125, 127, 129, 130),
    "Revenue": _s(132, 134, 135, 136),
    "Other Income": _s(138, 140, 143, 144),
    # Total = C148+C149 (fixed addition) and "Change in inventories" (149) = C150-C151, so this
    # block is not a growable list; opening/closing stock are addressable single cells.
    "COGS-Purchase": _s(146, 148, 148, 152, "total_last", False),
    "Opening Stock (Change in Inventories)": _s(146, 150, 150, 152, "single_cell", False),
    "Closing Stock (Change in Inventories)": _s(146, 151, 151, 152, "single_cell", False),
    "Employee Benefit Expenses": _s(154, 156, 168, 169),
    "Depreciation": _s(171, 173, 174, 175),
    "Finance Costs": _s(177, 180, 182, 183),    # row 179 "Interest expenses" is a non-summed sub-heading
    "Tax Expense": _s(185, 187, 190, 191),
    "COGS-Manufacturing": _s(195, 196, 204, 195, "total_first"),
    "Administrative Expenses": _s(206, 207, 232, 206, "total_first"),
    "Selling & Distribution Expenses": _s(234, 235, 246, 234, "total_first"),
}

# Schedules in top-to-bottom sheet order (ties broken by item row).
BASE_ORDER = sorted(BASE_SCHEDULES, key=lambda k: (BASE_SCHEDULES[k]["first_item_row"]))

# Rows added from the taxonomy in `Web app/CS format dropdown.xlsx` (see project plan, Phase 4).
NEW_ROWS = {
    "Fixed Assets": ["Land", "Building", "Machinery", "Intangible"],
    "Investments": ["Others"],
    "Short-term Provisions": ["Provision for expenses"],
    "COGS-Manufacturing": [
        "Direct labor", "Direct materials", "Manufacturing supplies", "Factory lighting",
        "Factory insurance", "Packing expenses", "Gas, water, and fuel",
        "Cost of raw materials", "Equipment", "Cargo",
    ],
    "Administrative Expenses": ["Other admin expenses"],
    "Revenue": ["Other operating income"],
}

# Hidden blank rows kept INSIDE each schedule's SUM range. A line item that is not in the taxonomy
# is written into one of these (row unhidden + highlighted yellow), so it is included in the
# schedule total and every downstream statement automatically -- no row insertion needed at runtime.
SPARE_ROWS = {
    "Long-term Provisions": 2,
    "Borrowings - Long Term": 3,
    "Trade Payables": 3,
    "Borrowings - Short Term": 3,
    "Short-term Provisions": 3,
    "Current Liability - Other": 4,
    "Fixed Assets": 4,
    "Long-term Loans & Advances": 2,
    "Investments": 2,
    "Trade Receivables": 3,
    "Inventory": 3,
    "Cash & Bank": 2,
    "Short-term Loans & Advances": 3,
    "Current Asset - Other": 4,
    "Revenue": 2,
    "Other Income": 3,
    "Employee Benefit Expenses": 4,
    "Depreciation": 2,
    "Finance Costs": 3,
    "Tax Expense": 2,
    "COGS-Manufacturing": 6,
    "Administrative Expenses": 8,
    "Selling & Distribution Expenses": 6,
}
