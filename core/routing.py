"""
Which Sch. schedule does a line item belong to?

One function, used by the workbook writer, the balance check and the analysis,
so they can never disagree about where an item went. An explicit choice made
in the review table (`schedule_override`) always wins.
"""

import re

CHANGE_IN_INVENTORIES = "CHANGE_IN_INVENTORIES"   # not a schedule: split into Opening/Closing stock cells
UNPLACED = None

_PNL_TAG_TO_SCHEDULE = {
    "Revenue": "Revenue",
    "Depreciation": "Depreciation",
    "Tax Expense": "Tax Expense",
    "Change in Inventories": CHANGE_IN_INVENTORIES,
}

_PNL_CATEGORY_TO_SCHEDULE = {
    "Administrative Expenses": "Administrative Expenses",
    "Selling & Distribution Expenses": "Selling & Distribution Expenses",
    "Finance Costs": "Finance Costs",
    "Other Income": "Other Income",
    "Employee Benefit Expenses": "Employee Benefit Expenses",
}

_BS_TAG_TO_SCHEDULE = {
    "Paid-up Share Capital": "Share Capital",
    "Opening Reserves": "Reserves and Surplus",
    "Securities Premium": "Securities Premium",
    "Reserves Adjustments": "Reserves Adjustments",
    "Deferred Tax Assets": "Deferred Tax Assets",
    "Deferred Tax Liabilities": "Deferred Tax Liabilities",
    "Fixed Assets": "Fixed Assets",
    "Investments": "Investments",
    "Cash & Bank": "Cash & Bank",
    "Trade Receivables": "Trade Receivables",
    "Inventory": "Inventory",
    "Current Asset - Other": "Current Asset - Other",
    "Current Liability - Other": "Current Liability - Other",
    "Trade Payables": "Trade Payables",
    "Short-term Provisions": "Short-term Provisions",
    "Long-term Provisions": "Long-term Provisions",
    "Short-term Loans & Advances": "Short-term Loans & Advances",
    "Long-term Loans & Advances": "Long-term Loans & Advances",
    "Borrowings - Long Term": "Borrowings - Long Term",
    "Borrowings - Short Term": "Borrowings - Short Term",
    "Borrowings": "Borrowings - Short Term",   # generic tag: short term unless told otherwise
}


def schedule_for(item):
    """Schedule key for a LineItem, CHANGE_IN_INVENTORIES, or None if it cannot be placed."""
    if item.schedule_override:
        return item.schedule_override
    if item.statement == "P&L":
        if item.statement_tag in _PNL_TAG_TO_SCHEDULE:
            return _PNL_TAG_TO_SCHEDULE[item.statement_tag]
        if item.category in _PNL_CATEGORY_TO_SCHEDULE:
            return _PNL_CATEGORY_TO_SCHEDULE[item.category]
        if item.category == "Direct Expenses / COGS":
            if re.search(r"purchase", item.line_item, re.I) and not re.search(r"stock|consum", item.line_item, re.I):
                return "COGS-Purchase"
            return "COGS-Manufacturing"
        return UNPLACED
    if item.statement == "BS":
        return _BS_TAG_TO_SCHEDULE.get(item.statement_tag, UNPLACED)
    return UNPLACED


def all_placeable_schedules():
    """Schedules a user may choose in the review table's 'place under' list."""
    keys = list(dict.fromkeys(
        list(_PNL_CATEGORY_TO_SCHEDULE.values())
        + ["Revenue", "Depreciation", "Tax Expense", "COGS-Purchase", "COGS-Manufacturing", CHANGE_IN_INVENTORIES]
        + list(dict.fromkeys(_BS_TAG_TO_SCHEDULE.values()))
    ))
    return keys
