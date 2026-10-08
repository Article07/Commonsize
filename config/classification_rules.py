"""
Keyword rules for functional expense classification.

Transcribed directly from Section 3 ("Functional Classification Rules") of
Project_Instructions_and_Preferences.docx. These are NOT a substitute for the
mandatory research-based approach in Sections 1-2 of that document -- they are
the no-API prototype's best-effort stand-in for the judgment an LLM (or a
Chartered Accountant) would otherwise apply. Any item that doesn't clearly
match is left as "Uncertain -- Needs Review" rather than force-classified,
per Section 5/6 of the instructions.
"""

COGS_KEYWORDS = [
    "raw material", "raw materials consumed", "purchase of stock in trade",
    "purchases of traded goods", "manufacturing expense", "factory wages",
    "factory salaries", "production labour", "production labor",
    "power and fuel", "power & fuel", "manufacturing overhead",
    "factory maintenance", "packing material", "freight inward",
    "job work charges", "conversion cost", "production consumables",
    "royalty", "cost of services rendered", "cost of goods sold",
    "changes in inventories", "stores and spares consumed",
    "consumption of raw material", "direct labour", "direct labor",
    "direct expenses",
]

SELLING_DISTRIBUTION_KEYWORDS = [
    "advertisement", "advertising", "sales promotion", "marketing expense",
    "brand promotion", "dealer incentive", "commission on sales",
    "brokerage on sales", "freight outward", "distribution expense",
    "transportation to customer", "export expense", "packaging for distribution",
    "warehousing", "customer support cost", "selling salaries",
    "sales office expense", "channel partner incentive", "trade discount",
    "sales commission", "carriage outward", "delivery expense",
]

ADMINISTRATIVE_KEYWORDS = [
    "office salaries", "corporate salaries", "directors' remuneration",
    "directors remuneration", "audit fee", "legal expense", "professional fee",
    "consultancy charge", "office rent", "insurance", "printing and stationery",
    "printing & stationery", "office maintenance", "general repairs",
    "telephone", "internet expense", "software subscription",
    "it administration", "hr expense", "secretarial expense",
    "registrar expense", "bank charges", "travelling expense",
    "traveling expense", "office electricity", "office security",
    "corporate governance expense", "compliance expense", "rates and taxes",
    "rates & taxes", "donation", "miscellaneous expense",
]

FINANCE_COST_KEYWORDS = [
    "interest expense", "interest on", "lease interest", "borrowing cost",
    "finance charge", "bank interest", "amortization of transaction cost",
    "interest paid", "processing fee on loan",
]

OTHER_INCOME_KEYWORDS = [
    "interest income", "dividend income", "profit on sale of",
    "gain on sale of", "rent received", "other income", "miscellaneous income",
    "foreign exchange gain", "excess provision written back",
]

# Employee costs sit outside the Section-3 functional split in the Reference
# file's own design -- the template carries a single "Employee benefit
# expense" P&L line (IS!C14, Sch. rows 154-169), not split across COGS/S&D/
# Admin by function. Items here come from the taxonomy in
# `Web app/CS format dropdown.xlsx` (Dataset sheet, Employee_benefit_expenses
# sub-category) -- see config/taxonomy.py for the full classification list.
EMPLOYEE_BENEFIT_KEYWORDS = [
    "directors remuneration", "directors' remuneration", "salaries and wages",
    "salary", "wages", "contribution to pf", "provident fund",
    "contribution to esic", "esic", "contribution to other funds",
    "health insurance", "life insurance", "workers and staff welfare",
    "staff welfare", "bonus paid", "bonus", "gratuity paid",
    "gratuity provision", "gratuity", "leave salary", "leave encashment",
]

# Order matters only for the printed rationale; scoring checks every category.
CATEGORY_KEYWORDS = {
    "Direct Expenses / COGS": COGS_KEYWORDS,
    "Selling & Distribution Expenses": SELLING_DISTRIBUTION_KEYWORDS,
    "Administrative Expenses": ADMINISTRATIVE_KEYWORDS,
    "Finance Costs": FINANCE_COST_KEYWORDS,
    "Other Income": OTHER_INCOME_KEYWORDS,
    "Employee Benefit Expenses": EMPLOYEE_BENEFIT_KEYWORDS,
}

UNCERTAIN_LABEL = "Uncertain -- Needs Review"
OTHER_OPERATING_LABEL = "Other Operating Expenses"

# Sub-routing within "Direct Expenses / COGS" for the Reference file template,
# which splits COGS across two separate schedules: "Cost of goods sold"
# (purchase/inventory-movement items) and "Direct expenses" (manufacturing
# overhead items). Used only by core/template_writer.py.
COGS_PURCHASE_KEYWORDS = [
    "purchase of stock in trade", "purchases of traded goods", "raw material",
    "raw materials consumed", "consumption of raw material",
    "changes in inventories", "stores and spares consumed",
]
COGS_MANUFACTURING_KEYWORDS = [
    kw for kw in COGS_KEYWORDS if kw not in COGS_PURCHASE_KEYWORDS
]

# Sub-routing for Revenue within the Reference file template's "Revenue from
# operations" schedule ("Sale of goods" vs "Sale of services").
REVENUE_SERVICE_KEYWORDS = ["service", "services rendered", "job work income"]

# Controlled vocabulary for line items that are identified directly rather
# than run through the Section-3 expense keyword classification -- either
# because they're not an expense (Revenue, BS/CF lines) or because they're a
# standard P&L line that sits outside the 6 functional expense categories
# (Depreciation, Tax Expense).
STATEMENT_TAGS = [
    "Revenue",
    "Depreciation",
    "Tax Expense",
    "Trade Receivables",
    "Inventory",
    "Trade Payables",
    "Cash & Bank",
    "Borrowings",  # generic fallback -- template_writer.py routes this to Short Term with a warning
    "Borrowings - Long Term",
    "Borrowings - Short Term",
    "Equity",  # generic Phase 1 tag -- not used by template_writer.py's Phase 3 Equity handling
    "Paid-up Share Capital",
    "Opening Reserves",
    "Securities Premium",
    "Current Asset - Other",
    "Current Liability - Other",
    "Fixed Assets",
    "Investments",
    "Deferred Tax Assets",
    "Deferred Tax Liabilities",
    "Reserves Adjustments",
    "Change in Inventories",
    "Long-term Loans & Advances",
    "Short-term Loans & Advances",
    "Long-term Provisions",
    "Short-term Provisions",
    "Operating CF",
    "Investing CF",
    "Financing CF",
]

# P&L rows carrying one of these tags are identified directly and skip the
# Section-3 keyword classification in classifier.py.
PNL_DIRECT_TAGS = {"Revenue", "Depreciation", "Tax Expense", "Change in Inventories"}
