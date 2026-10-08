"""
Master classification taxonomy, mined from `Web app/CS format dropdown.xlsx`
(Dataset sheet). Every entry is an item classified under a
Type > Category > Sub-category heading in that source file (e.g. Land sits
under Asset > Fixed > Tangible; Advertisement sits under
Expense > Expenses > Other_expenses).

Each entry maps that item to one of TWO things in our own system:
  - `target_kind="category"` + `target`: one of config.classification_rules'
    CATEGORY_KEYWORDS keys, or a PNL_DIRECT_TAGS tag (Revenue/Depreciation/
    Tax Expense) -- for P&L items, which still go through the existing
    category-based routing (including the COGS Purchase-vs-Manufacturing and
    Revenue goods-vs-services sub-routing already in template_writer.py).
  - `target_kind="tag"` + `target`: one of config.classification_rules'
    STATEMENT_TAGS -- for BS items, which are routed by schedule directly.

`target=None` means the source taxonomy has this item, but our template has
nowhere to put it yet (documented gaps: Preference share capital, and
Reserves & Surplus sub-types beyond Securities Premium/Opening). These are
deliberately left unmapped -- they'll be classified as "new/unclassified"
and flagged rather than silently merged into an unrelated row.

`heading` is kept for display purposes (e.g. showing "classified under
Asset > Fixed > Tangible" in the UI) and is never used for matching logic --
matching is purely against `item`. `code` is an internal-only unique key,
also never used for matching or display.
"""

from dataclasses import dataclass


@dataclass(frozen=True)
class TaxonomyEntry:
    item: str
    heading: str          # "Type > Category > Sub-category", for display only
    target_kind: str       # "category" | "tag" | None (unmapped)
    target: str | None
    code: str


def _bs(item, heading, tag, code):
    return TaxonomyEntry(item, heading, "tag", tag, code)


def _pnl(item, heading, category_or_direct_tag, code):
    return TaxonomyEntry(item, heading, "category", category_or_direct_tag, code)


def _unmapped(item, heading, code):
    return TaxonomyEntry(item, heading, None, None, code)


TAXONOMY = [
    # --- Asset > Fixed > Tangible ---
    _bs("Land", "Asset > Fixed > Tangible", "Fixed Assets", "A-F-T-1"),
    _bs("Building", "Asset > Fixed > Tangible", "Fixed Assets", "A-F-T-2"),
    _bs("Machinery", "Asset > Fixed > Tangible", "Fixed Assets", "A-F-T-3"),
    _bs("Office equipments", "Asset > Fixed > Tangible", "Fixed Assets", "A-F-T-4"),
    _bs("Vehicles", "Asset > Fixed > Tangible", "Fixed Assets", "A-F-T-5"),
    _bs("Furniture and fixtures", "Asset > Fixed > Tangible", "Fixed Assets", "A-F-T-6"),
    _bs("Computers", "Asset > Fixed > Tangible", "Fixed Assets", "A-F-T-7"),
    _bs("Intangible", "Asset > Fixed > Tangible", "Fixed Assets", "A-F-T-8"),
    # --- Asset > Fixed > DTA ---
    _bs("DTA", "Asset > Fixed > DTA", "Deferred Tax Assets", "A-F-D"),
    # --- Asset > Fixed > Investments ---
    _bs("Investment in shares", "Asset > Fixed > Investments", "Investments", "A-F-I-1"),
    _bs("Investmnet in mutual fund", "Asset > Fixed > Investments", "Investments", "A-F-I-2"),
    _bs("Others (Investments)", "Asset > Fixed > Investments", "Investments", "A-F-I-3"),
    # --- Asset > Fixed > Cash & cash equivalents ---
    _bs("Cash on hand", "Asset > Fixed > Cash & cash equivalents", "Cash & Bank", "A-F-C-1"),
    _bs("Balances with banks", "Asset > Fixed > Cash & cash equivalents", "Cash & Bank", "A-F-C-2"),
    _bs("Fixed deposit", "Asset > Fixed > Cash & cash equivalents", "Cash & Bank", "A-F-C-3"),
    # --- Asset > Current > Trade receivables ---
    _bs("Outstanding for a period more than six months", "Asset > Current > Trade receivables", "Trade Receivables", "A-C-TR-1"),
    _bs("Outstanding for a period less than six months", "Asset > Current > Trade receivables", "Trade Receivables", "A-C-TR-2"),
    _bs("Advance paid to suppliers", "Asset > Current > Trade receivables", "Trade Receivables", "A-C-TR-3"),
    # --- Asset > Current > Inventories ---
    _bs("Opening stock", "Asset > Current > Inventories", "Inventory", "A-C-I-1"),
    _bs("Closing stock", "Asset > Current > Inventories", "Inventory", "A-C-I-2"),
    # --- Asset > Current > Loans & Advances ---
    _bs("Advance tax", "Asset > Current > Loans & Advances", "Short-term Loans & Advances", "A-C-L-1"),
    _bs("Balance with revenue authorities", "Asset > Current > Loans & Advances", "Short-term Loans & Advances", "A-C-L-2"),
    _bs("Advance to employees", "Asset > Current > Loans & Advances", "Short-term Loans & Advances", "A-C-L-3"),
    _bs("Security deposits", "Asset > Current > Loans & Advances", "Long-term Loans & Advances", "A-C-L-4"),
    # --- Asset > Current > Other current assets ---
    _bs("Prepaid expenses", "Asset > Current > Other current assets", "Current Asset - Other", "A-C-O-1"),
    _bs("Focus in hand", "Asset > Current > Other current assets", "Current Asset - Other", "A-C-O-2"),
    _bs("Interest accrued", "Asset > Current > Other current assets", "Current Asset - Other", "A-C-O-3"),

    # --- Liability > Shareholders_fund > Share capital ---
    _bs("Equity", "Liability > Shareholders_fund > Share capital", "Paid-up Share Capital", "L-S-S-E"),
    _unmapped("Preference", "Liability > Shareholders_fund > Share capital", "L-S-S-P"),
    # --- Liability > Shareholders_fund > Reserves & surplus ---
    _bs("Opening (Reserves)", "Liability > Shareholders_fund > Reserves & surplus", "Opening Reserves", "L-S-R-1"),
    _unmapped("General reserve", "Liability > Shareholders_fund > Reserves & surplus", "L-S-R-2"),
    _bs("Securities premium", "Liability > Shareholders_fund > Reserves & surplus", "Securities Premium", "L-S-R-3"),
    _unmapped("Revaluation reserve", "Liability > Shareholders_fund > Reserves & surplus", "L-S-R-4"),
    _unmapped("CRR", "Liability > Shareholders_fund > Reserves & surplus", "L-S-R-5"),
    # --- Liability > Borrowings ---
    _bs("Bank loan", "Liability > Borrowings > Long_term_borrowings", "Borrowings - Long Term", "L-B-L-B"),
    _bs("Director's loan a/c", "Liability > Borrowings > Long_term_borrowings", "Borrowings - Long Term", "L-B-L-D"),
    _bs("Secured", "Liability > Borrowings > Short_term_borrowings", "Borrowings - Short Term", "L-B-S-S"),
    _bs("Unsecured", "Liability > Borrowings > Short_term_borrowings", "Borrowings - Short Term", "L-B-S-US"),
    # --- Liability > Current_liabilities > Trade_payables ---
    _bs("Sundry creditors for goods", "Liability > Current_liabilities > Trade_payables", "Trade Payables", "L-C-TP-1"),
    _bs("Sundry creditors for other than goods", "Liability > Current_liabilities > Trade_payables", "Trade Payables", "L-C-TP-2"),
    _bs("Advance from customers", "Liability > Current_liabilities > Trade_payables", "Trade Payables", "L-C-TP-3"),
    # --- Liability > Current_liabilities > Other_current_liabilities ---
    _bs("TDS & TCS payable", "Liability > Current_liabilities > Other_current_liabilities", "Current Liability - Other", "L-C-OCL-1"),
    _bs("GST payable", "Liability > Current_liabilities > Other_current_liabilities", "Current Liability - Other", "L-C-OCL-2"),
    _bs("ESIC payable", "Liability > Current_liabilities > Other_current_liabilities", "Current Liability - Other", "L-C-OCL-3"),
    _bs("PF payable", "Liability > Current_liabilities > Other_current_liabilities", "Current Liability - Other", "L-C-OCL-4"),
    _bs("Leave encashment payable", "Liability > Current_liabilities > Other_current_liabilities", "Current Liability - Other", "L-C-OCL-5"),
    _bs("Security taken for service centre", "Liability > Current_liabilities > Other_current_liabilities", "Current Liability - Other", "L-C-OCL-6"),
    _bs("Other payables", "Liability > Current_liabilities > Other_current_liabilities", "Current Liability - Other", "L-C-OCL-7"),
    # --- Liability > Current_liabilities > Provisions ---
    _bs("Provision for gratuity (LT)", "Liability > Current_liabilities > Long-term_provisions", "Long-term Provisions", "L-C-P-1"),
    _bs("Provision for gratuity (ST)", "Liability > Current_liabilities > Short-term_provisions", "Short-term Provisions", "L-C-P-2"),
    _bs("Current year's tax payable", "Liability > Current_liabilities > Short-term_provisions", "Short-term Provisions", "L-C-P-3"),
    _bs("Provision for expenses", "Liability > Current_liabilities > Short-term_provisions", "Short-term Provisions", "L-C-P-4"),

    # --- Income > Revenue > Operating ---
    _pnl("Sale of goods", "Income > Revenue > Operating", "Revenue", "I-O-1"),
    _pnl("Sale of services", "Income > Revenue > Operating", "Revenue", "I-O-2"),
    _pnl("Other operating income", "Income > Revenue > Operating", "Revenue", "I-O-3"),
    # --- Income > Revenue > Non_operating ---
    _pnl("Interest income", "Income > Revenue > Non_operating", "Other Income", "I-NO-1"),
    _pnl("Other non-operating income", "Income > Revenue > Non_operating", "Other Income", "I-NO-2"),
    _pnl("Other misc. receipts", "Income > Revenue > Non_operating", "Other Income", "I-NO-3"),
    _pnl("Sale of fixed assets", "Income > Revenue > Non_operating", "Other Income", "I-NO-4"),

    # --- Expense > Expenses > Purchases ---
    _pnl("Purchase of traded goods", "Expense > Expenses > Purchases", "Direct Expenses / COGS", "E-P"),
    # --- Expense > Expenses > Direct_expenses ---
    _pnl("Direct labor", "Expense > Expenses > Direct_expenses", "Direct Expenses / COGS", "E-D-1"),
    _pnl("Direct materials", "Expense > Expenses > Direct_expenses", "Direct Expenses / COGS", "E-D-2"),
    _pnl("Manufacturing supplies", "Expense > Expenses > Direct_expenses", "Direct Expenses / COGS", "E-D-3"),
    _pnl("Wages for the production staff", "Expense > Expenses > Direct_expenses", "Direct Expenses / COGS", "E-D-4"),
    _pnl("Fuel or power consumption", "Expense > Expenses > Direct_expenses", "Direct Expenses / COGS", "E-D-5"),
    _pnl("Freight inwards", "Expense > Expenses > Direct_expenses", "Direct Expenses / COGS", "E-D-6"),
    _pnl("Factory lighting", "Expense > Expenses > Direct_expenses", "Direct Expenses / COGS", "E-D-7"),
    _pnl("Factory rent", "Expense > Expenses > Direct_expenses", "Direct Expenses / COGS", "E-D-8"),
    _pnl("Factory insurance", "Expense > Expenses > Direct_expenses", "Direct Expenses / COGS", "E-D-9"),
    _pnl("Insurance stock", "Expense > Expenses > Direct_expenses", "Direct Expenses / COGS", "E-D-10"),
    _pnl("Packing expenses (direct)", "Expense > Expenses > Direct_expenses", "Direct Expenses / COGS", "E-D-11"),
    _pnl("Gas, water, and fuel", "Expense > Expenses > Direct_expenses", "Direct Expenses / COGS", "E-D-12"),
    _pnl("Cost of raw materials", "Expense > Expenses > Direct_expenses", "Direct Expenses / COGS", "E-D-13"),
    _pnl("Equipment", "Expense > Expenses > Direct_expenses", "Direct Expenses / COGS", "E-D-14"),
    _pnl("Manufacturing expenses", "Expense > Expenses > Direct_expenses", "Direct Expenses / COGS", "E-D-15"),
    _pnl("Cargo", "Expense > Expenses > Direct_expenses", "Direct Expenses / COGS", "E-D-16"),
    # --- Expense > Expenses > Employee_benefit_expenses ---
    _pnl("Directors remuneration", "Expense > Expenses > Employee_benefit_expenses", "Employee Benefit Expenses", "E-E-1"),
    _pnl("Salaries and wages", "Expense > Expenses > Employee_benefit_expenses", "Employee Benefit Expenses", "E-E-2"),
    _pnl("Co's contribution to PF", "Expense > Expenses > Employee_benefit_expenses", "Employee Benefit Expenses", "E-E-3"),
    _pnl("Co's contribution to ESIC", "Expense > Expenses > Employee_benefit_expenses", "Employee Benefit Expenses", "E-E-4"),
    _pnl("Co's contribution to other funds", "Expense > Expenses > Employee_benefit_expenses", "Employee Benefit Expenses", "E-E-5"),
    _pnl("Health/ Life insurance", "Expense > Expenses > Employee_benefit_expenses", "Employee Benefit Expenses", "E-E-6"),
    _pnl("Workers and staff welfare", "Expense > Expenses > Employee_benefit_expenses", "Employee Benefit Expenses", "E-E-7"),
    _pnl("Bonus paid", "Expense > Expenses > Employee_benefit_expenses", "Employee Benefit Expenses", "E-E-8"),
    _pnl("Gratuity paid", "Expense > Expenses > Employee_benefit_expenses", "Employee Benefit Expenses", "E-E-9"),
    _pnl("Gratuity provision (P&L)", "Expense > Expenses > Employee_benefit_expenses", "Employee Benefit Expenses", "E-E-10"),
    _pnl("Leave salary", "Expense > Expenses > Employee_benefit_expenses", "Employee Benefit Expenses", "E-E-11"),
    # --- Expense > Expenses > Other_expenses (Administrative, per existing Sch. rows 207-232) ---
    _pnl("Annual maintenace charges", "Expense > Expenses > Other_expenses", "Administrative Expenses", "E-OE-1"),
    _pnl("Audit fees", "Expense > Expenses > Other_expenses", "Administrative Expenses", "E-OE-2"),
    _pnl("Additional demand duties and taxes", "Expense > Expenses > Other_expenses", "Administrative Expenses", "E-OE-3"),
    _pnl("BIS & Certification charges", "Expense > Expenses > Other_expenses", "Administrative Expenses", "E-OE-4"),
    _pnl("Donation", "Expense > Expenses > Other_expenses", "Administrative Expenses", "E-OE-5"),
    _pnl("Electricity expenses", "Expense > Expenses > Other_expenses", "Administrative Expenses", "E-OE-6"),
    _pnl("Fees & taxes", "Expense > Expenses > Other_expenses", "Administrative Expenses", "E-OE-7"),
    _pnl("Festival celebration expenses", "Expense > Expenses > Other_expenses", "Administrative Expenses", "E-OE-8"),
    _pnl("Fine & penalties", "Expense > Expenses > Other_expenses", "Administrative Expenses", "E-OE-9"),
    _pnl("Insurance charges", "Expense > Expenses > Other_expenses", "Administrative Expenses", "E-OE-10"),
    _pnl("Loss on sale of fixed asset", "Expense > Expenses > Other_expenses", "Administrative Expenses", "E-OE-11"),
    _pnl("Miscellaneous expenses", "Expense > Expenses > Other_expenses", "Administrative Expenses", "E-OE-12"),
    _pnl("Office expenses", "Expense > Expenses > Other_expenses", "Administrative Expenses", "E-OE-13"),
    _pnl("Printing and stationery expenses", "Expense > Expenses > Other_expenses", "Administrative Expenses", "E-OE-14"),
    _pnl("Postage and courier expenses", "Expense > Expenses > Other_expenses", "Administrative Expenses", "E-OE-15"),
    _pnl("Professional charges", "Expense > Expenses > Other_expenses", "Administrative Expenses", "E-OE-16"),
    _pnl("Rent", "Expense > Expenses > Other_expenses", "Administrative Expenses", "E-OE-17"),
    _pnl("Repair & maintenance", "Expense > Expenses > Other_expenses", "Administrative Expenses", "E-OE-18"),
    _pnl("Software expenses", "Expense > Expenses > Other_expenses", "Administrative Expenses", "E-OE-19"),
    _pnl("Subscription/ Membership fees", "Expense > Expenses > Other_expenses", "Administrative Expenses", "E-OE-20"),
    _pnl("Stamps & duties", "Expense > Expenses > Other_expenses", "Administrative Expenses", "E-OE-21"),
    _pnl("Telephone charges", "Expense > Expenses > Other_expenses", "Administrative Expenses", "E-OE-22"),
    _pnl("Travelling expenses foreign", "Expense > Expenses > Other_expenses", "Administrative Expenses", "E-OE-23"),
    _pnl("Vehicle running expenses", "Expense > Expenses > Other_expenses", "Administrative Expenses", "E-OE-24"),
    # --- Expense > Expenses > Other_expenses (Selling & Distribution, per existing Sch. rows 235-246) ---
    _pnl("Advertisement", "Expense > Expenses > Other_expenses", "Selling & Distribution Expenses", "E-OE-25"),
    _pnl("Business promotion", "Expense > Expenses > Other_expenses", "Selling & Distribution Expenses", "E-OE-26"),
    _pnl("Commission", "Expense > Expenses > Other_expenses", "Selling & Distribution Expenses", "E-OE-27"),
    _pnl("Conference", "Expense > Expenses > Other_expenses", "Selling & Distribution Expenses", "E-OE-28"),
    _pnl("E-commerece support service", "Expense > Expenses > Other_expenses", "Selling & Distribution Expenses", "E-OE-29"),
    _pnl("Freight & cartage (outward)", "Expense > Expenses > Other_expenses", "Selling & Distribution Expenses", "E-OE-30"),
    _pnl("Marketing staff conveyance expesnses", "Expense > Expenses > Other_expenses", "Selling & Distribution Expenses", "E-OE-31"),
    _pnl("Schemes and samples", "Expense > Expenses > Other_expenses", "Selling & Distribution Expenses", "E-OE-32"),
    _pnl("DR/CR Amt. w/o", "Expense > Expenses > Other_expenses", "Selling & Distribution Expenses", "E-OE-33"),
    _pnl("Travellijng exp. within India", "Expense > Expenses > Other_expenses", "Selling & Distribution Expenses", "E-OE-34"),
    _pnl("Warranty & replacement", "Expense > Expenses > Other_expenses", "Selling & Distribution Expenses", "E-OE-35"),
    # --- Expense > Expenses > Other_expenses (new catch-all) ---
    _pnl("Other admin expenses", "Expense > Expenses > Other_expenses", "Administrative Expenses", "E-OE-36"),
    # --- Expense > Expenses > Depreciation ---
    _pnl("Depreciation", "Expense > Expenses > Depreciation", "Depreciation", "E-DEP"),
    # --- Expense > Expenses > Finance_cost ---
    _pnl("Interest to bank", "Expense > Expenses > Finance_cost", "Finance Costs", "E-FC-1"),
    _pnl("Interest to unsecured loans", "Expense > Expenses > Finance_cost", "Finance Costs", "E-FC-2"),
    _pnl("Bank charges", "Expense > Expenses > Finance_cost", "Finance Costs", "E-FC-3"),
    # --- Expense > Expenses > Tax ---
    _pnl("Current tax", "Expense > Expenses > Tax", "Tax Expense", "E-TAX"),
]

# Fast lookup: code -> entry (internal use only, e.g. de-duplication checks).
TAXONOMY_BY_CODE = {e.code: e for e in TAXONOMY}

# Only the items with a real mapping are useful for matching against.
MATCHABLE_TAXONOMY = [e for e in TAXONOMY if e.target is not None]


# ---------------------------------------------------------------------------
# Heading / sub-category aliases
#
# Face statements and notes mostly speak at the sub-category level ("Trade
# receivables", "Other current liabilities") rather than the item level the
# taxonomy lists. These patterns resolve that level, and are also applied to
# the HEADING a line sits under in the PDF, which is how an item that is not
# in the taxonomy still lands under the heading the PDF puts it under.
#
# (regex, target_kind, target)   first match wins, so order matters.
# ---------------------------------------------------------------------------

# (regex, target_kind, target, row_label)
#   row_label = the template row to use when a total has to stand in for a breakup
#   (e.g. a face-statement "Employee Benefit Expenses" total whose note could not be read).
BS_ALIASES = [
    (r"securities\s*premium", "tag", "Securities Premium", ""),
    (r"share\s*capital|equity\s*shares?|paid.?up", "tag", "Paid-up Share Capital", ""),
    (r"def+er+ed\s*tax\s*liabilit", "tag", "Deferred Tax Liabilities", ""),
    (r"def+er+ed\s*tax\s*asset", "tag", "Deferred Tax Assets", ""),
    (r"long.?term\s*provision", "tag", "Long-term Provisions", "Provision for gratuity"),
    (r"short.?term\s*provision|provisions?\s*for", "tag", "Short-term Provisions", "Provision for expenses"),
    (r"long.?term\s*loans?\s*(and|&)\s*advances|security\s*deposit", "tag", "Long-term Loans & Advances", "Security deposits"),
    (r"short.?term\s*loans?\s*(and|&)\s*advances|balance\s*with\s*revenue|advances?\s*to\s*employees?", "tag", "Short-term Loans & Advances", "Advance to employees"),
    (r"long.?term\s*borrowing|\bterm\s*loan(?!s?\s*(and|&)\s*advance)|loans?\s*from\s*(directors?|promoters?)|debentures?", "tag", "Borrowings - Long Term", "Term loan"),
    (r"short.?term\s*borrowing|working\s*capital|cash\s*credit|overdraft", "tag", "Borrowings - Short Term", "Secured loan from banks"),
    (r"trade\s*payable|sundry\s*creditors?|creditors?|msme|micro.*small", "tag", "Trade Payables", "Sundry creditors for goods"),
    (r"other\s*current\s*liabilit|statutory\s*dues|customer\s*deposits?", "tag", "Current Liability - Other", "Other payables"),
    (r"property|plant\s*(and|&)\s*equip|fixed\s*assets?|tangible|capital\s*work", "tag", "Fixed Assets", "Property, plant and equipments"),
    (r"investment", "tag", "Investments", "Investment in shares"),
    (r"inventor|stock.in.trade|work.in.progress|stores\s*(and|&)\s*spares|packing\s*material|raw\s*material|finished\s*goods", "tag", "Inventory", "Finished goods"),
    (r"trade\s*receivable|sundry\s*debtors?|debtors?", "tag", "Trade Receivables", "Outstanding for a period less than six months"),
    (r"cash\s*(and|&)\s*(cash\s*)?(equivalent|bank)|bank\s*balance|cash\s*in\s*hand|balances?\s*with\s*bank|current\s*accounts?", "tag", "Cash & Bank", "Balances with banks"),
    (r"other\s*(current|non.?current)\s*assets?|prepaid|accrued", "tag", "Current Asset - Other", "Prepaid exp."),
]

PNL_ALIASES = [
    (r"revenue\s*from\s*operations?|sale\s*of\s*(products?|goods|services)|^sales?\b|sales\s*manufactur", "tag", "Revenue", "Sale of goods"),
    (r"other\s*income|discounts?\s*received|interest\s*(income|received)|miscellaneous\s*income", "category", "Other Income", "Miscellaneous Income"),
    (r"changes?\s*in\s*inventor", "tag", "Change in Inventories", ""),
    (r"cost\s*of\s*(raw\s*)?materials?|materials?\s*consumed|purchases?\b|opening\s*stock|closing\s*stock", "category", "Direct Expenses / COGS", "Cost of raw materials"),
    (r"employee|salar|wages|staff|bonus|gratuity|provident|esic|contribution\s*to", "category", "Employee Benefit Expenses", "Salaries and wages"),
    (r"finance\s*costs?|interest\s*(on|expense|paid)|bank\s*charges", "category", "Finance Costs", "To bank"),
    (r"depreciation|amorti[sz]ation", "tag", "Depreciation", "Depreciation"),
    (r"tax\s*expense|current\s*tax|deferred\s*tax|previous\s*year\s*tax|income\s*tax\s*(expense|provision)", "tag", "Tax Expense", "Current tax"),
    (r"selling|distribution|marketing|advertis", "category", "Selling & Distribution Expenses", "Business promotion"),
    (r"direct\s*expenses?|manufacturing|production|factory", "category", "Direct Expenses / COGS", "Manufacturing expenses"),
    (r"other\s*expenses?|administrative|general\s*expenses?|miscellaneous", "category", "Administrative Expenses", "Other admin expenses"),
]

# Filler words that are not a "new line item" when they sit under a recognised heading.
GENERIC_FILLERS = {"others", "other", "miscellaneous", "misc", "sundry", "general"}

# Taxonomy item names too generic to match on their own (they only make sense inside their heading).
GENERIC_ITEM_NAMES = {"others", "other", "equity", "opening", "secured", "unsecured", "preference", "dta"}
