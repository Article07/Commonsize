"""
Tunable thresholds for the rule-based Data Integrity & Consistency Checks
(Section 9 of Project_Instructions_and_Preferences.docx). These are
prototype defaults, not calibrated against any specific industry -- adjust
as you see real output and decide what "material" means for your clients.
"""

# Common-size % point movement (of Revenue) that triggers a flag, e.g. COGS
# % of Revenue rising by more than this many percentage points YoY.
COMMON_SIZE_PT_MOVE_FLAG = 2.0

# YoY % change in an expense line that counts as a "sudden increase".
EXPENSE_YOY_PCT_FLAG = 20.0

# Gross margin decline (percentage points) YoY that counts as deterioration.
GROSS_MARGIN_DECLINE_PT_FLAG = 2.0

# YoY growth gap (percentage points) between an expense's growth rate and
# Revenue's growth rate that flags "growing faster than Revenue".
GROWTH_GAP_VS_REVENUE_PT_FLAG = 10.0

# Inventory / Receivables YoY growth outpacing Revenue YoY growth by this
# many percentage points triggers a build-up / deterioration flag.
WORKING_CAPITAL_GROWTH_GAP_PT_FLAG = 15.0

# Finance cost YoY % increase that is flagged if Borrowings did NOT increase.
FINANCE_COST_YOY_PCT_FLAG = 10.0

# Effective tax rate change (percentage points) YoY that is flagged.
TAX_RATE_CHANGE_PT_FLAG = 5.0
