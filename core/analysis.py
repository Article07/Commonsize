"""
Rule-based Analysis sheet content: Executive Summary, YoY Variance,
Common Size Analysis, Ratio Commentary, and Data Integrity & Consistency
Checks (Section 7-9 of the instructions doc). Every finding here is a
templated sentence built from computed numbers and threshold comparisons --
there is no AI-written narrative in this no-API prototype. Findings marked
needs_irl feed into irl.py.

Wherever a metric can't be computed because the required Statement Tags
weren't supplied, the corresponding line is emitted as an explicit
"Insufficient data provided" finding rather than silently omitted.
"""

from config import thresholds as T
from core.statements import compute_statements
from core.models import Finding

INSUFFICIENT = "Insufficient data provided"


def _sum_by_category(line_items, fy, category):
    total = None
    for li in line_items:
        if li.statement == "P&L" and li.category == category:
            v = li.fy_values.get(fy)
            if v is not None:
                total = (total or 0.0) + v
    return total


def _sum_by_tag(line_items, fy, tag, statement=None):
    total = None
    for li in line_items:
        if li.statement_tag == tag and (statement is None or li.statement == statement):
            v = li.fy_values.get(fy)
            if v is not None:
                total = (total or 0.0) + v
    return total


def _pct(numerator, denominator):
    if numerator is None or not denominator:
        return None
    return round((numerator / denominator) * 100, 2)


def _yoy_pct(prev, curr):
    if prev in (None, 0) or curr is None:
        return None
    return round(((curr - prev) / prev) * 100, 2)


def _nz(value, present=True):
    """0.0 from an empty schedule means 'not provided'; keep None so the 'Insufficient data' paths still work."""
    return value if (present and value) else None


def compute_metrics(line_items, fy_columns):
    """Per-FY metrics, built on core/statements.py so they agree with the workbook and the balance check."""
    ordered = sorted(fy_columns)
    statements, _ = compute_statements(line_items, ordered)
    metrics = {}
    for fy in fy_columns:
        s = statements[fy]
        sched = s["sched"]
        revenue = _nz(s["revenue"])
        cogs = _nz(s["cogs"] + s["direct_expenses"])
        sd = _nz(sched["Selling & Distribution Expenses"][fy])
        admin = _nz(sched["Administrative Expenses"][fy])
        employee = _nz(s["employee"])
        finance_cost = _nz(s["finance_cost"])
        other_income = _nz(s["other_income"])
        depreciation = _nz(s["depreciation"])
        tax = _nz(s["tax"])
        uncertain = _sum_by_category(line_items, fy, "Uncertain -- Needs Review")

        gross_profit = None if revenue is None else revenue - (cogs or 0.0)
        ebitda = None
        if gross_profit is not None:
            ebitda = gross_profit - (employee or 0.0) - (sd or 0.0) - (admin or 0.0)
        ebit = None if ebitda is None or depreciation is None else ebitda - depreciation
        net_profit = None
        if ebit is not None:
            net_profit = ebit - (finance_cost or 0.0) + (other_income or 0.0) - (tax or 0.0)

        trade_receivables = _nz(s["receivables"])
        inventory = _nz(s["inventory"])
        cash = _nz(s["cash"])
        trade_payables = _nz(s["trade_payables"])
        borrowings = _nz(s["borrowings"])
        equity = _nz(s["net_worth"])
        current_assets = (s["receivables"] + s["inventory"] + s["cash"]
                          + sched["Short-term Loans & Advances"][fy] + sched["Current Asset - Other"][fy]) or None
        current_liabilities = (s["trade_payables"] + sched["Current Liability - Other"][fy]
                               + sched["Short-term Provisions"][fy]) or None
        operating_cf = _sum_by_tag(line_items, fy, "Operating CF", "CF")

        metrics[fy] = dict(
            revenue=revenue,
            cogs=cogs,
            gross_profit=gross_profit,
            gross_margin_pct=_pct(gross_profit, revenue),
            sd=sd,
            admin=admin,
            employee=employee,
            uncertain_expense=uncertain,
            ebitda=ebitda,
            ebitda_margin_pct=_pct(ebitda, revenue),
            depreciation=depreciation,
            ebit=ebit,
            ebit_margin_pct=_pct(ebit, revenue),
            finance_cost=finance_cost,
            other_income=other_income,
            tax=tax,
            net_profit=net_profit,
            net_margin_pct=_pct(net_profit, revenue),
            trade_receivables=trade_receivables,
            inventory=inventory,
            trade_payables=trade_payables,
            cash=cash,
            borrowings=borrowings,
            equity=equity,
            current_assets=current_assets,
            current_liabilities=current_liabilities,
            current_ratio=round(current_assets / current_liabilities, 2)
            if current_assets is not None and current_liabilities
            else None,
            debt_equity=round(borrowings / equity, 2)
            if borrowings is not None and equity
            else None,
            total_assets=_nz(s["total_assets"]),
            operating_cf=operating_cf,
        )
    return metrics


def build_executive_summary(metrics, fy_columns):
    findings = []
    trend_specs = [
        ("Revenue", "revenue", None),
        ("Gross Margin", "gross_margin_pct", "%"),
        ("EBITDA Margin (approx., pre-Finance Cost)", "ebitda_margin_pct", "%"),
        ("EBIT Margin (approx.)", "ebit_margin_pct", "%"),
        ("Net Profit Margin (approx.)", "net_margin_pct", "%"),
        ("Current Ratio", "current_ratio", "x"),
        ("Debt-Equity", "debt_equity", "x"),
        ("Operating Cash Flow", "operating_cf", None),
    ]
    for label, key, unit in trend_specs:
        series = [(fy, metrics[fy].get(key)) for fy in fy_columns]
        available = [(fy, v) for fy, v in series if v is not None]
        if len(available) < 2:
            findings.append(
                Finding(
                    section="Executive Summary",
                    area=label,
                    observation=f"{label} trend: {INSUFFICIENT} (need at least two FY with data).",
                    needs_irl=False,
                )
            )
            continue
        first_fy, first_val = available[0]
        last_fy, last_val = available[-1]
        direction = "increased" if last_val > first_val else ("decreased" if last_val < first_val else "held flat")
        suffix = unit or ""
        findings.append(
            Finding(
                section="Executive Summary",
                area=label,
                observation=(
                    f"{label} {direction} from {first_val}{suffix} in {first_fy} to "
                    f"{last_val}{suffix} in {last_fy}."
                ),
                needs_irl=False,
            )
        )
    return findings


def build_yoy_variance(line_items, fy_columns):
    findings = []
    tracked_categories = [
        "Direct Expenses / COGS",
        "Selling & Distribution Expenses",
        "Administrative Expenses",
        "Employee Benefit Expenses",
        "Finance Costs",
        "Other Income",
    ]
    for prev_fy, curr_fy in zip(fy_columns, fy_columns[1:]):
        for category in tracked_categories:
            prev_val = _sum_by_category(line_items, prev_fy, category)
            curr_val = _sum_by_category(line_items, curr_fy, category)
            change_pct = _yoy_pct(prev_val, curr_val)
            if change_pct is None:
                continue
            if abs(change_pct) >= T.EXPENSE_YOY_PCT_FLAG:
                findings.append(
                    Finding(
                        section="YoY Variance Analysis",
                        area=category,
                        observation=(
                            f"{category} changed by {change_pct}% from {prev_fy} to {curr_fy} "
                            f"({prev_val:,.2f} -> {curr_val:,.2f})."
                        ),
                        reason_for_concern=f"Movement exceeds the {T.EXPENSE_YOY_PCT_FLAG}% materiality threshold.",
                        info_required=f"Management explanation for the {category} movement between {prev_fy} and {curr_fy}.",
                    )
                )
    return findings


def build_common_size_flags(metrics, fy_columns):
    findings = []
    for prev_fy, curr_fy in zip(fy_columns, fy_columns[1:]):
        prev_gm = metrics[prev_fy].get("gross_margin_pct")
        curr_gm = metrics[curr_fy].get("gross_margin_pct")
        if prev_gm is not None and curr_gm is not None:
            delta = round(curr_gm - prev_gm, 2)
            if delta <= -T.GROSS_MARGIN_DECLINE_PT_FLAG:
                findings.append(
                    Finding(
                        section="Common Size Analysis",
                        area="Gross Margin",
                        observation=f"Gross Margin declined by {abs(delta)} pts from {prev_fy} ({prev_gm}%) to {curr_fy} ({curr_gm}%).",
                        reason_for_concern="Decline exceeds materiality threshold; may indicate pricing pressure or rising input costs.",
                        info_required="Breakdown of raw material cost trends and any pricing changes during the period.",
                    )
                )

        for label, key in [
            ("COGS % of Revenue", "cogs"),
            ("Selling & Distribution % of Revenue", "sd"),
            ("Administrative % of Revenue", "admin"),
            ("Finance Cost % of Revenue", "finance_cost"),
        ]:
            revenue_prev = metrics[prev_fy].get("revenue")
            revenue_curr = metrics[curr_fy].get("revenue")
            val_prev = metrics[prev_fy].get(key)
            val_curr = metrics[curr_fy].get(key)
            pct_prev = _pct(val_prev, revenue_prev)
            pct_curr = _pct(val_curr, revenue_curr)
            if pct_prev is not None and pct_curr is not None:
                delta = round(pct_curr - pct_prev, 2)
                if abs(delta) >= T.COMMON_SIZE_PT_MOVE_FLAG:
                    direction = "increased" if delta > 0 else "decreased"
                    findings.append(
                        Finding(
                            section="Common Size Analysis",
                            area=label,
                            observation=f"{label} {direction} by {abs(delta)} pts from {prev_fy} ({pct_prev}%) to {curr_fy} ({pct_curr}%).",
                            reason_for_concern=f"Movement exceeds the {T.COMMON_SIZE_PT_MOVE_FLAG} pt materiality threshold.",
                            info_required=f"Management explanation for the shift in {label} between {prev_fy} and {curr_fy}.",
                        )
                    )
    return findings


def build_ratio_commentary(metrics, fy_columns):
    findings = []
    ratio_specs = [
        ("Current Ratio", "current_ratio", "liquidity"),
        ("Debt-Equity Ratio", "debt_equity", "solvency"),
        ("Gross Margin %", "gross_margin_pct", "profitability"),
        ("EBITDA Margin % (approx.)", "ebitda_margin_pct", "profitability"),
        ("Net Profit Margin % (approx.)", "net_margin_pct", "profitability"),
    ]
    any_data = False
    for label, key, dimension in ratio_specs:
        series = [(fy, metrics[fy].get(key)) for fy in fy_columns if metrics[fy].get(key) is not None]
        if len(series) < 2:
            continue
        any_data = True
        first_fy, first_val = series[0]
        last_fy, last_val = series[-1]
        if last_val > first_val:
            trend = "improving"
        elif last_val < first_val:
            trend = "deteriorating"
        else:
            trend = "stable"
        finding = Finding(
            section="Ratio Commentary",
            area=label,
            observation=f"{label} is {trend}: {first_val} ({first_fy}) -> {last_val} ({last_fy}).",
            needs_irl=(trend == "deteriorating"),
        )
        if trend == "deteriorating":
            finding.reason_for_concern = f"{dimension.capitalize()} concern -- {label} has weakened over the period."
            finding.info_required = f"Management commentary on the drivers behind the {label} trend."
        findings.append(finding)

    if not any_data:
        findings.append(
            Finding(
                section="Ratio Commentary",
                area="Ratios",
                observation=f"Ratio Commentary: {INSUFFICIENT} (requires Balance Sheet Statement Tags).",
                needs_irl=False,
            )
        )
    return findings


def build_data_integrity_checks(line_items, metrics, fy_columns):
    findings = []
    for prev_fy, curr_fy in zip(fy_columns, fy_columns[1:]):
        rev_prev, rev_curr = metrics[prev_fy].get("revenue"), metrics[curr_fy].get("revenue")
        rev_growth = _yoy_pct(rev_prev, rev_curr)

        # COGS increasing faster than Revenue
        cogs_growth = _yoy_pct(metrics[prev_fy].get("cogs"), metrics[curr_fy].get("cogs"))
        if rev_growth is not None and cogs_growth is not None:
            gap = round(cogs_growth - rev_growth, 2)
            if gap >= T.GROWTH_GAP_VS_REVENUE_PT_FLAG:
                findings.append(Finding(
                    section="Data Integrity", area="COGS vs Revenue",
                    observation=f"COGS grew {cogs_growth}% vs Revenue growth of {rev_growth}% ({prev_fy}->{curr_fy}), a gap of {gap} pts.",
                    reason_for_concern="COGS growing materially faster than Revenue erodes gross margin and may signal cost control or pricing issues.",
                    info_required="Explanation for COGS growth outpacing Revenue growth, incl. any one-off cost items.",
                ))

        # Employee cost disproportionate to revenue -- approximated via Admin (no separate Employee Cost tag in v1)
        admin_growth = _yoy_pct(metrics[prev_fy].get("admin"), metrics[curr_fy].get("admin"))
        if rev_growth is not None and admin_growth is not None:
            gap = round(admin_growth - rev_growth, 2)
            if gap >= T.GROWTH_GAP_VS_REVENUE_PT_FLAG:
                findings.append(Finding(
                    section="Data Integrity", area="Administrative Expenses vs Revenue",
                    observation=f"Administrative Expenses grew {admin_growth}% vs Revenue growth of {rev_growth}% ({prev_fy}->{curr_fy}).",
                    reason_for_concern="Administrative cost growth disproportionate to Revenue growth.",
                    info_required="Breakdown of Administrative Expense increase and whether it is one-off or structural.",
                ))

        # Finance costs increasing despite lower/flat borrowings
        finance_growth = _yoy_pct(metrics[prev_fy].get("finance_cost"), metrics[curr_fy].get("finance_cost"))
        borrowings_prev, borrowings_curr = metrics[prev_fy].get("borrowings"), metrics[curr_fy].get("borrowings")
        if finance_growth is not None and finance_growth >= T.FINANCE_COST_YOY_PCT_FLAG:
            if borrowings_prev is not None and borrowings_curr is not None and borrowings_curr <= borrowings_prev:
                findings.append(Finding(
                    section="Data Integrity", area="Finance Costs vs Borrowings",
                    observation=f"Finance Costs rose {finance_growth}% ({prev_fy}->{curr_fy}) while Borrowings did not increase ({borrowings_prev:,.2f} -> {borrowings_curr:,.2f}).",
                    reason_for_concern="Rising finance costs without a corresponding rise in borrowings is inconsistent absent a rate or classification explanation.",
                    info_required="Reconciliation of Finance Cost increase against average borrowing balances and interest rates.",
                ))
            elif borrowings_prev is None or borrowings_curr is None:
                findings.append(Finding(
                    section="Data Integrity", area="Finance Costs vs Borrowings",
                    observation=f"Finance Costs rose {finance_growth}% ({prev_fy}->{curr_fy}); Borrowings data not provided to cross-check.",
                    reason_for_concern="Cannot verify finance cost movement against borrowings without Balance Sheet data.",
                    info_required="Borrowings balances for both years to reconcile against the Finance Cost movement.",
                ))

        # Revenue growth without corresponding receivable growth
        recv_growth = _yoy_pct(metrics[prev_fy].get("trade_receivables"), metrics[curr_fy].get("trade_receivables"))
        if rev_growth is not None and rev_growth >= T.EXPENSE_YOY_PCT_FLAG and recv_growth is not None:
            gap = round(rev_growth - recv_growth, 2)
            if gap >= T.WORKING_CAPITAL_GROWTH_GAP_PT_FLAG:
                findings.append(Finding(
                    section="Data Integrity", area="Revenue vs Trade Receivables",
                    observation=f"Revenue grew {rev_growth}% but Trade Receivables grew only {recv_growth}% ({prev_fy}->{curr_fy}).",
                    reason_for_concern="Revenue growth not mirrored in receivables may warrant review of collection terms or revenue recognition.",
                    info_required="Ageing schedule of Trade Receivables and any change in credit terms.",
                ))

        # Significant inventory build-up
        inv_growth = _yoy_pct(metrics[prev_fy].get("inventory"), metrics[curr_fy].get("inventory"))
        if inv_growth is not None and rev_growth is not None:
            gap = round(inv_growth - rev_growth, 2)
            if gap >= T.WORKING_CAPITAL_GROWTH_GAP_PT_FLAG:
                findings.append(Finding(
                    section="Data Integrity", area="Inventory",
                    observation=f"Inventory grew {inv_growth}% vs Revenue growth of {rev_growth}% ({prev_fy}->{curr_fy}).",
                    reason_for_concern="Inventory build-up disproportionate to Revenue may indicate slow-moving stock or demand mismatch.",
                    info_required="Inventory ageing and slow/non-moving stock analysis.",
                ))

        # Negative operating cash flow despite reported profit
        ocf_curr = metrics[curr_fy].get("operating_cf")
        net_profit_curr = metrics[curr_fy].get("net_profit")
        if ocf_curr is not None and net_profit_curr is not None and ocf_curr < 0 and net_profit_curr > 0:
            findings.append(Finding(
                section="Data Integrity", area="Operating Cash Flow vs Net Profit",
                observation=f"{curr_fy}: Operating Cash Flow is negative ({ocf_curr:,.2f}) despite a reported profit ({net_profit_curr:,.2f}).",
                reason_for_concern="Profit not converting to operating cash is a key quality-of-earnings red flag.",
                info_required="Reconciliation between Net Profit and Operating Cash Flow (working capital movements, non-cash items).",
            ))

        # Significant changes in tax rates (effective tax rate = tax / (ebit - finance_cost + other_income), approx pre-tax profit)
        def _pretax(fy):
            m = metrics[fy]
            if m.get("ebit") is None:
                return None
            return m["ebit"] - (m.get("finance_cost") or 0.0) + (m.get("other_income") or 0.0)

        pretax_prev, pretax_curr = _pretax(prev_fy), _pretax(curr_fy)
        tax_prev, tax_curr = metrics[prev_fy].get("tax"), metrics[curr_fy].get("tax")
        etr_prev = _pct(tax_prev, pretax_prev)
        etr_curr = _pct(tax_curr, pretax_curr)
        if etr_prev is not None and etr_curr is not None:
            delta = round(etr_curr - etr_prev, 2)
            if abs(delta) >= T.TAX_RATE_CHANGE_PT_FLAG:
                findings.append(Finding(
                    section="Data Integrity", area="Effective Tax Rate",
                    observation=f"Effective tax rate moved by {delta} pts from {prev_fy} ({etr_prev}%) to {curr_fy} ({etr_curr}%).",
                    reason_for_concern="Material tax rate movement should be understood before relying on Net Profit trends.",
                    info_required="Tax computation/reconciliation explaining the effective rate change.",
                ))

    # Uncertain expense items are themselves a standing data-integrity/classification concern
    uncertain_items = [li for li in line_items if li.confidence.startswith("Uncertain")]
    for li in uncertain_items:
        findings.append(Finding(
            section="Data Integrity", area="Expense Classification",
            observation=f"'{li.line_item}' could not be confidently classified by the rule-based engine ({li.confidence}).",
            reason_for_concern="Per the classification instructions, items must not be force-classified without evidentiary support.",
            info_required="Management/Annual Report clarification on the nature and function of this expense to confirm classification.",
        ))

    return findings


def run_analysis(line_items, fy_columns):
    metrics = compute_metrics(line_items, fy_columns)
    sections = {
        "Executive Summary": build_executive_summary(metrics, fy_columns),
        "YoY Variance Analysis": build_yoy_variance(line_items, fy_columns),
        "Common Size Analysis": build_common_size_flags(metrics, fy_columns),
        "Ratio Commentary": build_ratio_commentary(metrics, fy_columns),
        "Data Integrity": build_data_integrity_checks(line_items, metrics, fy_columns),
    }
    return metrics, sections
