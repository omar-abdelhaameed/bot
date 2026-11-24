"""Where things live in MAZBOOT_Finance.xlsx.

Must match build_workbook.py. If you move tables in the workbook, update these.
"""

TAB_SETTINGS = "الإعدادات"
TAB_TX = "الحركات"
TAB_BUDGET = "الميزانية"

# Settings tab
SET_START_MONTH = "C5"
SET_OPENING_CASH = "C6"
SET_ALERT_PCT = "C9"
SET_PARTNERS = "B13:B16"
SET_CATEGORIES = "B20:F44"          # name, class, budget driver, value, bot keywords

# Ledger tab
TX_HEADER_ROW = 4
TX_FIRST_ROW = 5
TX_LAST_ROW = 10000
TX_COLS = ["id", "date", "category", "details", "qty", "amount", "person",
           "method", "source", "logged_at", "notes"]          # A..K
TX_LAST_COL = "K"

# Budget tab: month headers in D5:O5, planned block rows 13..37 (category name in column B)
BUD_MONTH_ROW = 5
BUD_FIRST_MONTH_COL = "D"
BUD_LAST_MONTH_COL = "O"
BUD_PLANNED_FIRST = 13
BUD_PLANNED_LAST = 37

# Category classes
CLS_INCOME = "إيراد"
CLS_FUNDING = "تمويل"
CLS_WITHDRAW = "سحب"
CLS_COGS = "تكلفة مباشرة"
CLS_OPEX = "تشغيل"
CLS_CAPEX = "تأسيس"
EXPENSE_CLASSES = (CLS_COGS, CLS_OPEX, CLS_CAPEX)
MONEY_IN = (CLS_INCOME, CLS_FUNDING)
MONEY_OUT = EXPENSE_CLASSES + (CLS_WITHDRAW,)

METHODS = ["كاش الشركة", "من جيبه", "فودافون كاش", "إنستاباي", "تحويل بنكي"]
DEFAULT_METHOD = "كاش الشركة"
OUT_OF_POCKET = "من جيبه"
SOURCE_BOT = "بوت"
