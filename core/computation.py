"""
Common-size % and YoY variance computation.

- P&L items are sized against total Revenue (Statement Tag == "Revenue").
- BS items are sized against Total Assets, computed as the sum of asset-side
  Statement Tags (Fixed Assets, Investments, Cash & Bank, Trade Receivables,
  Inventory, Current Asset - Other).
- CF items are not common-sized (no single natural denominator) -- YoY change
  is still computed for them.

If Revenue or the asset tags are missing, the corresponding common-size
columns are left as None rather than guessed -- callers should render that
as "Insufficient data provided".
"""

import re

ASSET_TAGS = {
    "Fixed Assets",
    "Investments",
    "Cash & Bank",
    "Trade Receivables",
    "Inventory",
    "Current Asset - Other",
}


def _fy_sort_key(fy_label):
    match = re.search(r"\d+", fy_label)
    return int(match.group()) if match else 0


def ordered_fy_columns(line_items):
    fy_labels = set()
    for li in line_items:
        fy_labels.update(li.fy_values.keys())
    return sorted(fy_labels, key=_fy_sort_key)


def _total_for_tag_set(line_items, statement, tag_set, fy):
    total = None
    for li in line_items:
        if li.statement == statement and li.statement_tag in tag_set:
            val = li.fy_values.get(fy)
            if val is not None:
                total = (total or 0.0) + val
    return total


def compute_revenue_by_fy(line_items, fy_columns):
    return {
        fy: _total_for_tag_set(line_items, "P&L", {"Revenue"}, fy)
        for fy in fy_columns
    }


def compute_total_assets_by_fy(line_items, fy_columns):
    return {
        fy: _total_for_tag_set(line_items, "BS", ASSET_TAGS, fy)
        for fy in fy_columns
    }


def compute_common_size(line_items, fy_columns=None):
    """Populates li.common_size on every item. Returns (revenue_by_fy, total_assets_by_fy)."""
    fy_columns = fy_columns or ordered_fy_columns(line_items)
    revenue_by_fy = compute_revenue_by_fy(line_items, fy_columns)
    total_assets_by_fy = compute_total_assets_by_fy(line_items, fy_columns)

    for li in line_items:
        common_size = {}
        for fy in fy_columns:
            value = li.fy_values.get(fy)
            denominator = revenue_by_fy.get(fy) if li.statement == "P&L" else (
                total_assets_by_fy.get(fy) if li.statement == "BS" else None
            )
            if value is None or not denominator:
                common_size[fy] = None
            else:
                common_size[fy] = round((value / denominator) * 100, 2)
        li.common_size = common_size

    return revenue_by_fy, total_assets_by_fy


def compute_yoy_change(line_items, fy_columns=None):
    """Populates li.yoy_change_pct on every item (percent change vs the prior FY column)."""
    fy_columns = fy_columns or ordered_fy_columns(line_items)
    for li in line_items:
        changes = {}
        for prev_fy, curr_fy in zip(fy_columns, fy_columns[1:]):
            prev_val = li.fy_values.get(prev_fy)
            curr_val = li.fy_values.get(curr_fy)
            if prev_val in (None, 0) or curr_val is None:
                changes[curr_fy] = None
            else:
                changes[curr_fy] = round(((curr_val - prev_val) / prev_val) * 100, 2)
        li.yoy_change_pct = changes
    return line_items
