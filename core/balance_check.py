"""
Balance-check: does the Balance Sheet tally for every financial year?

Built on core/statements.py, so it applies the same arithmetic as the
template. The result is never forced to zero: if the data is incomplete or a
figure was misread, the gap is reported as it is, with the items that could
not be placed.
"""

from dataclasses import dataclass, field

from core.statements import compute_statements


@dataclass
class YearCheck:
    fy: str
    liabilities_side: float   # net worth + borrowings
    assets_side: float        # fixed assets + deferred tax + investments + cash + net current assets
    gap: float                # liabilities_side - assets_side
    balanced: bool
    pat: float


@dataclass
class BalanceResult:
    years: list = field(default_factory=list)
    unplaced: list = field(default_factory=list)
    statements: dict = field(default_factory=dict)

    @property
    def balanced(self):
        return bool(self.years) and all(y.balanced for y in self.years)


def balance_check(line_items, fy_columns, tolerance=1.0):
    """tolerance is in the same unit as the item values (rupees): rounding in the PDF's own units is not a failure."""
    fy_columns = sorted(fy_columns)
    statements, unplaced = compute_statements(line_items, fy_columns)
    years = []
    for fy in fy_columns:
        s = statements[fy]
        years.append(YearCheck(
            fy=fy, liabilities_side=s["liabilities_side"], assets_side=s["assets_side"],
            gap=s["gap"], balanced=abs(s["gap"]) <= tolerance, pat=s["pat"],
        ))
    return BalanceResult(years=years, unplaced=unplaced, statements=statements)
