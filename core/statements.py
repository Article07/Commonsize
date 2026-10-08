"""
Python replica of the Reference file's own income-statement and balance-sheet
arithmetic (IS!C6:C22, BS!C8:C36), computed from the routed line items.

This is the single source of truth for: the balance check shown in the app,
the key figures, and the Analysis sheet. It follows the template's layout
exactly, including its quirks:

  * the BS is presented "net": Net worth + Borrowings must equal Fixed assets
    + Deferred tax (net) + Investments + Cash + (Current assets - Current
    liabilities);
  * Reserves = Securities premium + Surplus closing, where Surplus closing =
    Opening + other movements + Profit for the year (the PAT from the IS);
    the Opening balance is an input only for the first year, later years open
    with the previous year's closing.

Because profit flows into reserves, an error anywhere in the P&L shows up as
a Balance Sheet gap -- which is exactly why the tally is a meaningful check.
"""

from core.routing import CHANGE_IN_INVENTORIES, schedule_for

SCHEDULE_NAMES = (
    "Revenue", "Other Income", "COGS-Purchase", "COGS-Manufacturing", "Employee Benefit Expenses",
    "Administrative Expenses", "Selling & Distribution Expenses", "Depreciation", "Finance Costs", "Tax Expense",
    "Share Capital", "Reserves and Surplus", "Securities Premium", "Reserves Adjustments",
    "Borrowings - Long Term", "Borrowings - Short Term", "Trade Payables", "Current Liability - Other",
    "Long-term Provisions", "Short-term Provisions", "Fixed Assets", "Deferred Tax Assets",
    "Deferred Tax Liabilities", "Investments", "Cash & Bank", "Trade Receivables", "Inventory",
    "Long-term Loans & Advances", "Short-term Loans & Advances", "Current Asset - Other",
)


def schedule_totals(line_items, fy_columns):
    """{schedule key: {fy: total}} plus the list of items that could not be placed."""
    totals = {name: {fy: 0.0 for fy in fy_columns} for name in SCHEDULE_NAMES}
    totals["Change in Inventories"] = {fy: 0.0 for fy in fy_columns}
    unplaced = []
    for item in line_items:
        if item.statement not in ("P&L", "BS"):
            continue
        key = schedule_for(item)
        if key is None:
            if any(v for v in item.fy_values.values() if v):
                unplaced.append(item)
            continue
        if key == CHANGE_IN_INVENTORIES:
            key = "Change in Inventories"
        if key not in totals:
            continue
        for fy in fy_columns:
            value = item.fy_values.get(fy)
            if value is not None:
                totals[key][fy] += value
    return totals, unplaced


def compute_statements(line_items, fy_columns):
    """-> (per-FY dict of statement figures, unplaced items). fy_columns must be in ascending order."""
    totals, unplaced = schedule_totals(line_items, fy_columns)
    t = lambda key, fy: totals[key][fy]
    result = {}
    previous_closing = None
    for index, fy in enumerate(fy_columns):
        revenue = t("Revenue", fy)
        cogs = t("COGS-Purchase", fy) + t("Change in Inventories", fy)
        direct = t("COGS-Manufacturing", fy)
        gross_margin = revenue - cogs - direct
        employee = t("Employee Benefit Expenses", fy)
        other_expenses = t("Administrative Expenses", fy) + t("Selling & Distribution Expenses", fy)
        ebitda = gross_margin - employee - other_expenses
        depreciation, finance = t("Depreciation", fy), t("Finance Costs", fy)
        operating_pbt = ebitda - depreciation - finance
        other_income, tax = t("Other Income", fy), t("Tax Expense", fy)
        pat = operating_pbt + other_income - tax

        opening = t("Reserves and Surplus", fy) if previous_closing is None else previous_closing
        surplus_closing = opening + t("Reserves Adjustments", fy) + pat
        reserves = t("Securities Premium", fy) + surplus_closing
        previous_closing = surplus_closing
        share_capital = t("Share Capital", fy)
        net_worth = share_capital + reserves
        borrowings = t("Borrowings - Long Term", fy) + t("Borrowings - Short Term", fy)
        liabilities_side = net_worth + borrowings

        fixed = t("Fixed Assets", fy)
        deferred_tax_net = t("Deferred Tax Assets", fy) - t("Deferred Tax Liabilities", fy)
        investments, cash = t("Investments", fy), t("Cash & Bank", fy)
        receivables, inventory = t("Trade Receivables", fy), t("Inventory", fy)
        current_assets = (receivables + inventory + t("Long-term Loans & Advances", fy)
                          + t("Short-term Loans & Advances", fy) + t("Current Asset - Other", fy))
        current_liabilities = (t("Trade Payables", fy) + t("Current Liability - Other", fy)
                               + t("Long-term Provisions", fy) + t("Short-term Provisions", fy))
        net_current_assets = current_assets - current_liabilities
        assets_side = fixed + deferred_tax_net + investments + cash + net_current_assets

        result[fy] = dict(
            revenue=revenue, cogs=cogs, direct_expenses=direct, gross_margin=gross_margin,
            employee=employee, other_expenses=other_expenses, ebitda=ebitda,
            depreciation=depreciation, finance_cost=finance, operating_pbt=operating_pbt,
            other_income=other_income, tax=tax, pat=pat,
            opening_reserves=opening, reserves=reserves, share_capital=share_capital, net_worth=net_worth,
            borrowings=borrowings, liabilities_side=liabilities_side,
            fixed_assets=fixed, deferred_tax_net=deferred_tax_net, investments=investments, cash=cash,
            receivables=receivables, inventory=inventory, current_assets=current_assets,
            current_liabilities=current_liabilities, net_current_assets=net_current_assets,
            assets_side=assets_side, gap=liabilities_side - assets_side,
            total_assets=fixed + max(deferred_tax_net, 0) + investments + cash + current_assets,
            trade_payables=t("Trade Payables", fy),
            sched=totals,
        )
    return result, unplaced
