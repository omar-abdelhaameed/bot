"""Builds MAZBOOT_Finance.xlsx — the MAZBOOT accounting, budgeting and planning workbook.

Single source of truth: the ledger tab (الحركات). Every other tab is formulas over the
ledger plus the yellow input cells. The Telegram bot appends rows to the ledger only.
All formulas are Excel-2007-era functions so the file works the same in Excel,
LibreOffice and Google Sheets.
"""
import datetime as dt
import sys

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

OUT = sys.argv[1] if len(sys.argv) > 1 else "MAZBOOT_Finance.xlsx"

# ---------------------------------------------------------------- tab names
S_DASH = "لوحة التحكم"
S_SET = "الإعدادات"
S_TX = "الحركات"
S_START = "التأسيس"
S_SAL = "المرتبات"
S_BUD = "الميزانية"
S_PL = "قائمة الدخل"
S_CF = "التدفق النقدي"
S_PART = "الشركاء"
S_MEAL = "تكلفة الوجبات"
S_HELP = "دليل الاستخدام"


def q(name):
    return f"'{name}'"


# ---------------------------------------------------------------- styles
FONT = "Arial"
GREEN = "1F5A3F"      # MAZBOOT street green
DEEP = "123826"
CREAM = "F6E6BD"
MUSTARD = "FFC23D"
RED = "E5463A"
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
HEAD_FILL = PatternFill("solid", fgColor=GREEN)
SUB_FILL = PatternFill("solid", fgColor=CREAM)
TOTAL_FILL = PatternFill("solid", fgColor="E8EFEA")
CARD_FILL = PatternFill("solid", fgColor="F7F4EC")
thin = Side(style="thin", color="C9C2B0")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

F_BASE = Font(name=FONT, size=10)
F_INPUT = Font(name=FONT, size=10, color="0000FF")
F_LINK = Font(name=FONT, size=10, color="008000")
F_HEAD = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_BOLD = Font(name=FONT, size=10, bold=True)
F_TITLE = Font(name=FONT, size=16, bold=True, color=DEEP)
F_NOTE = Font(name=FONT, size=9, italic=True, color="6B6B6B")
F_KPI = Font(name=FONT, size=16, bold=True, color=DEEP)
F_KPI_LABEL = Font(name=FONT, size=9, bold=True, color="486A58")

NUM = '#,##0;(#,##0);"-"'
NUM2 = '#,##0.0;(#,##0.0);"-"'
PCT = '0.0%;(0.0%);"-"'
DATE = "yyyy-mm-dd"
MONTH = "mmm yyyy"

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center", wrap_text=True)
NOWRAP = Alignment(horizontal="right", vertical="center", wrap_text=False)

wb = Workbook()


def sheet(title, first=False):
    ws = wb.active if first else wb.create_sheet()
    ws.title = title
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = GREEN
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return ws


def put(ws, ref, value, font=F_BASE, fill=None, fmt=None, align=None, border=False):
    c = ws[ref]
    c.value = value
    c.font = font
    if fill:
        c.fill = fill
    if fmt:
        c.number_format = fmt
    if align is None and font in (F_TITLE, F_NOTE, F_BOLD) and isinstance(value, str) and len(value) > 28:
        align = NOWRAP
    c.alignment = align or RIGHT
    if border:
        c.border = BORDER
    return c


def inp(ws, ref, value, fmt=None, note=None):
    c = put(ws, ref, value, font=F_INPUT, fill=INPUT_FILL, fmt=fmt, border=True)
    if note:
        c.comment = Comment(note, "MAZBOOT")
    return c


def head(ws, row, col_from, labels, fill=HEAD_FILL, font=F_HEAD, height=30):
    for i, lab in enumerate(labels):
        put(ws, f"{L(col_from + i)}{row}", lab, font=font, fill=fill, align=CENTER, border=True)
    ws.row_dimensions[row].height = height


def title(ws, text, sub=None):
    put(ws, "B2", text, font=F_TITLE)
    ws.row_dimensions[2].height = 28
    if sub:
        put(ws, "B3", sub, font=F_NOTE)


def widths(ws, mapping):
    for col, w in mapping.items():
        ws.column_dimensions[col].width = w


# ---------------------------------------------------------------- constants
START_MONTH = dt.date(2026, 11, 1)
N_MONTHS = 12
TX_FIRST, TX_LAST = 5, 10000
CAT_FIRST, CAT_LAST = 20, 44           # category rows on the settings tab
N_CAT = CAT_LAST - CAT_FIRST + 1
PART_FIRST, PART_LAST = 13, 16         # partner rows on the settings tab

CLASSES = ["إيراد", "تمويل", "سحب", "تكلفة مباشرة", "تشغيل", "تأسيس"]
DRIVERS = ["ثابت شهري", "% من المبيعات", "لكل أوردر", "المبيعات المتوقعة",
           "من تاب المرتبات", "رأس مال الشركاء", "بدون"]
METHODS = ["كاش الشركة", "من جيبه", "فودافون كاش", "إنستاباي", "تحويل بنكي"]
SOURCES = ["يدوي", "بوت"]

# name, class, driver, value, bot keywords, note
CATEGORIES = [
    ("مبيعات", "إيراد", "المبيعات المتوقعة", None,
     "مبيعات,بعنا,بعت,اوردر,اوردرات,أوردر,أوردرات,دخل,ايراد,إيراد,تحصيل,قبضنا", "الكمية = عدد الأوردرات"),
    ("إيرادات أخرى", "إيراد", "بدون", None, "ايرادات اخرى,إيرادات أخرى,رعاية,سبونسر", ""),
    ("رأس مال", "تمويل", "رأس مال الشركاء", None,
     "راس مال,رأس مال,رأسمال,راسمال,حصته,حصتي,مساهمة,تمويل,حط فلوس", "الشخص = الشريك اللي دفع"),
    ("سحب شريك", "سحب", "بدون", None, "سحب,سحبت,سحوبات,سداد شريك,رجعتله", "سحب أرباح أو سداد لشريك"),
    ("خامات وبضاعة", "تكلفة مباشرة", "% من المبيعات", 0.35,
     "بضاعة,بضاعه,خامات,فراخ,صدور,دجاج,لحمة,لحم,مفروم,تونة,رز,أرز,ارز,مكرونة,خضار,فاكهة,بيض,جبنة,جبن,"
     "زبادي,لبن,زيت,زيتون,توابل,بهارات,صوص,صوصات,عيش,خبز,تورتيلا,بطاطس,بطاطا,خس,طماطم,بصل,ثوم,فلفل,"
     "ليمون,شوفان,عسل,سوبر ماركت,سوبرماركت,جملة,مكونات,مخزون", "القيمة = نسبة تكلفة الأكل من المبيعات"),
    ("تغليف", "تكلفة مباشرة", "لكل أوردر", 15,
     "تغليف,علب,علبة,عبوات,عبوة,اكياس,أكياس,كيس,شنط,ليبل,ليبلز,استيكر,ستيكر,بوكس,بوكسات,كراتين,سلوفان,فويل,"
     "معالق,شوك,كاسات", "القيمة = تكلفة التغليف لكل أوردر"),
    ("إعلانات ممولة", "تشغيل", "لكل أوردر", 10,
     "اعلان,إعلان,اعلانات,إعلانات,ممول,ممولة,سبونسرد,ads,ميتا,بوست", "القيمة = ميزانية الإعلان لكل أوردر (قرار 27 أكتوبر)"),
    ("تصوير ومحتوى", "تشغيل", "ثابت شهري", 1000,
     "تصوير,محتوى,مونتاج,مصور,بلوجر,انفلونسر,إنفلونسر,مؤثر,مؤثرين,تذوق", "تقديري"),
    ("إيجار", "تشغيل", "ثابت شهري", 4500, "ايجار,إيجار,الشقة,المطبخ", "تقديري — جزء من الـ 10 آلاف الثابتة"),
    ("كهرباء ومياه وغاز", "تشغيل", "ثابت شهري", 2000,
     "كهرباء,كهربا,نور,مياه,مية,غاز,انبوبة,أنبوبة,عداد", "تقديري"),
    ("إنترنت وموبايل", "تشغيل", "ثابت شهري", 500, "نت,انترنت,إنترنت,رصيد,شحن,باقة,موبايل", "تقديري"),
    ("مرتبات", "تشغيل", "من تاب المرتبات", None, "مرتب,مرتبات,يومية,يوميات,قبض,أجر,اجر,بونص", "بتتحسب من تاب المرتبات"),
    ("توصيل", "تشغيل", "لكل أوردر", 0,
     "توصيل,دليفري,ديليفري,طيار,بنزين,مواصلات,اوبر,أوبر,توكتوك", "لو التوصيل على حساب العميل خليها صفر"),
    ("منظفات ونظافة", "تشغيل", "ثابت شهري", 500,
     "منظفات,صابون,كلور,مناديل,جوانتي,جوانتيات,نظافة,ديتول,اسفنج", "تقديري"),
    ("صيانة", "تشغيل", "ثابت شهري", 500, "صيانة,تصليح,فني,سباك,كهربائي,قطع غيار", "تقديري"),
    ("برامج واشتراكات", "تشغيل", "ثابت شهري", 300,
     "اشتراك,اشتراكات,برنامج,دومين,استضافة,هوستنج,كانفا,canva,سيرفر", "تقديري"),
    ("رسوم وتراخيص", "تشغيل", "بدون", None,
     "ترخيص,تراخيص,سجل تجاري,بطاقة ضريبية,ضرايب,ضرائب,شهادة صحية,رسوم,دمغة,محامي,محاسب", "بتيجي من تاب التأسيس"),
    ("نثريات", "تشغيل", "ثابت شهري", 700, "نثريات,متفرقات,حاجات صغيرة,شاي,قهوة", "تقديري"),
    ("معدات وأدوات", "تأسيس", "بدون", None,
     "ثلاجة,تلاجة,فريزر,ديب فريزر,بوتاجاز,فرن,جريل,شواية,ميزان,خلاط,كبة,حلل,حلة,طاسة,سكاكين,سكينة,معدات,"
     "أدوات,ادوات,طفاية,شفاط,طابعة,ايرفراير,إير فراير,ميكروويف,ترابيزة,رفوف,أرفف", "بتيجي من تاب التأسيس"),
    ("تجهيز المكان", "تأسيس", "بدون", None,
     "تشطيب,دهان,نقاشة,ديكور,تجهيز,مقدم,تأمين الشقة,عفش,سباكة,كهرباء المطبخ,سيراميك", "بتيجي من تاب التأسيس"),
]
assert len(CATEGORIES) <= N_CAT
CAT_ROW = {c[0]: CAT_FIRST + i for i, c in enumerate(CATEGORIES)}

PARTNERS = [("عمر", 0, 0.55), ("أسامة", 150000, 0.25), ("أحمد", 50000, 0.20)]

# ---------------------------------------------------------------- shared ranges
TX = q(S_TX)
R_DATE = f"{TX}!$B${TX_FIRST}:$B${TX_LAST}"
R_CAT = f"{TX}!$C${TX_FIRST}:$C${TX_LAST}"
R_QTY = f"{TX}!$E${TX_FIRST}:$E${TX_LAST}"
R_AMT = f"{TX}!$F${TX_FIRST}:$F${TX_LAST}"
R_PER = f"{TX}!$G${TX_FIRST}:$G${TX_LAST}"
R_MET = f"{TX}!$H${TX_FIRST}:$H${TX_LAST}"
SET = q(S_SET)
R_CATS = f"{SET}!$B${CAT_FIRST}:$B${CAT_LAST}"
R_CLASS = f"{SET}!$C${CAT_FIRST}:$C${CAT_LAST}"
R_DRV = f"{SET}!$D${CAT_FIRST}:$D${CAT_LAST}"
R_VAL = f"{SET}!$E${CAT_FIRST}:$E${CAT_LAST}"
R_PNAMES = f"{SET}!$B${PART_FIRST}:$B${PART_LAST}"

# settings general cells
C_BRAND, C_START, C_OPEN, C_AVG, C_DAYS, C_ALERT, C_RESERVE = (
    f"{SET}!$C$4", f"{SET}!$C$5", f"{SET}!$C$6", f"{SET}!$C$7", f"{SET}!$C$8", f"{SET}!$C$9", f"{SET}!$C$10")
C_TOTCAP = f"{SET}!$C${PART_LAST + 1}"
C_VARPCT, C_VARORD, C_CONTRIB = f"{SET}!$G$5", f"{SET}!$G$6", f"{SET}!$G$7"
SALES_CAT = f"{SET}!$B${CAT_ROW['مبيعات']}"
CAPITAL_CAT = f"{SET}!$B${CAT_ROW['رأس مال']}"
WITHDRAW_CAT = f"{SET}!$B${CAT_ROW['سحب شريك']}"

MCOLS = [L(4 + i) for i in range(N_MONTHS)]       # D..O on month-grid tabs
TOTCOL = L(4 + N_MONTHS)                           # P


def month_filter(col, row=5):
    return f'{R_DATE},">="&{col}${row},{R_DATE},"<="&EOMONTH({col}${row},0)'


def sign_expr():
    """+1 for money in, -1 for money out, per category class (array over settings)."""
    return (f'(({R_CLASS}="إيراد")+({R_CLASS}="تمويل")-({R_CLASS}="تكلفة مباشرة")'
            f'-({R_CLASS}="تشغيل")-({R_CLASS}="تأسيس")-({R_CLASS}="سحب"))')


# ================================================================ 1. Dashboard (filled last, created first)
ws_dash = sheet(S_DASH, first=True)

# ================================================================ 2. Settings
ws = sheet(S_SET)
title(ws, "الإعدادات والمدخلات", "الخانات الصفرا بس هي اللي بتتعدل. كل التابات التانية بتتحسب منها ومن تاب الحركات.")
widths(ws, {"A": 2, "B": 30, "C": 18, "D": 20, "E": 14, "F": 70, "G": 16, "H": 34})
put(ws, "B4", "اسم البراند", F_BOLD); inp(ws, "C4", "MAZBOOT")
put(ws, "B5", "أول شهر في الخطة", F_BOLD)
inp(ws, "C5", START_MONTH, MONTH, "أول يوم في أول شهر. كل جداول الـ 12 شهر بتبدأ منه.")
put(ws, "B6", "الكاش الموجود قبل أي حركة (ج)", F_BOLD)
inp(ws, "C6", 0, NUM, "أي فلوس كانت في الخزنة قبل ما تبدأ تسجل. رأس المال نفسه بيتسجل في الحركات.")
put(ws, "B7", "متوسط قيمة الأوردر (ج)", F_BOLD)
inp(ws, "C7", 240, NUM, "من نموذج البيزنس (15 أكتوبر): الأوردر المباشر 240 ج. عدّله على أسعار المنيو الجديد.")
put(ws, "B8", "أيام الشغل في الشهر", F_BOLD); inp(ws, "C8", 30, NUM)
put(ws, "B9", "تنبيه الميزانية عند", F_BOLD)
inp(ws, "C9", 0.8, PCT, "البوت والداشبورد بينبهوا لما فئة توصل للنسبة دي من ميزانيتها.")
put(ws, "B10", "احتياطي الطوارئ (ج)", F_BOLD)
inp(ws, "C10", 30000, NUM, "من نموذج البيزنس: 30 ألف احتياطي. فلوس متتصرفش غير في الطوارئ.")

put(ws, "F4", "حسابات تلقائية", F_BOLD)
put(ws, "F5", "تكلفة متغيرة كنسبة من المبيعات"); put(ws, "G5", f'=SUMIF({R_DRV},"% من المبيعات",{R_VAL})', fmt=PCT)
put(ws, "F6", "تكلفة متغيرة لكل أوردر (ج)"); put(ws, "G6", f'=SUMIF({R_DRV},"لكل أوردر",{R_VAL})', fmt=NUM)
put(ws, "F7", "اللي بيفضل من كل أوردر بعد التكلفة المتغيرة (ج)")
put(ws, "G7", "=C7-C7*G5-G6", fmt=NUM, font=F_BOLD)

head(ws, 12, 2, ["الشريك", "رأس المال المتفق عليه (ج)", "نسبة الملكية"])
for i, (n, cap, pct) in enumerate(PARTNERS):
    r = PART_FIRST + i
    inp(ws, f"B{r}", n); inp(ws, f"C{r}", cap, NUM); inp(ws, f"D{r}", pct, PCT)
for r in range(PART_FIRST + len(PARTNERS), PART_LAST + 1):
    inp(ws, f"B{r}", None); inp(ws, f"C{r}", None, NUM); inp(ws, f"D{r}", None, PCT)
put(ws, f"B{PART_LAST + 1}", "الإجمالي", F_BOLD, TOTAL_FILL, border=True)
put(ws, f"C{PART_LAST + 1}", f"=SUM(C{PART_FIRST}:C{PART_LAST})", F_BOLD, TOTAL_FILL, NUM, border=True)
put(ws, f"D{PART_LAST + 1}", f"=SUM(D{PART_FIRST}:D{PART_LAST})", F_BOLD, TOTAL_FILL, PCT, border=True)
put(ws, f"E{PART_FIRST}", "النسب 55/25/20 حسب آخر اتفاق. لو غيرتوا التقسيم عدّل هنا والتابات كلها هتتحدث.", F_NOTE)
put(ws, f"E{PART_FIRST + 1}", '=IF(ROUND(D17,4)<>1,"⚠️ مجموع النسب لازم يبقى 100%","✓ مجموع النسب 100%")', F_BOLD)

head(ws, 19, 2, ["الفئة", "التصنيف", "طريقة حساب الميزانية", "القيمة", "كلمات البوت (مفصولة بفاصلة)", "ملاحظة"])
for i in range(N_CAT):
    r = CAT_FIRST + i
    if i < len(CATEGORIES):
        n, cls, drv, val, kw, note = CATEGORIES[i]
    else:
        n = cls = drv = val = kw = note = None
    inp(ws, f"B{r}", n); inp(ws, f"C{r}", cls); inp(ws, f"D{r}", drv)
    inp(ws, f"E{r}", val, PCT if drv == "% من المبيعات" else NUM)
    inp(ws, f"F{r}", kw); put(ws, f"G{r}", note, F_NOTE)
    ws[f"F{r}"].alignment = Alignment(horizontal="right", vertical="center", wrap_text=False)
put(ws, f"B{CAT_LAST + 2}",
    "التصنيف: إيراد / تمويل (رأس مال) / سحب / تكلفة مباشرة (أكل وتغليف) / تشغيل (مصاريف شهرية) / تأسيس (معدات وتجهيز). "
    "ضيف فئة جديدة في الصفوف الفاضية وهتظهر في الميزانية والداشبورد والبوت لوحدها.", F_NOTE)
put(ws, f"B{CAT_LAST + 3}",
    "طريقة الميزانية: ثابت شهري = القيمة كل شهر · % من المبيعات = القيمة × المبيعات المتوقعة · لكل أوردر = القيمة × عدد الأوردرات المتوقع · "
    "أي بند في تاب التأسيس بيتضاف لفئته في شهر صرفه.", F_NOTE)
put(ws, f"B{CAT_LAST + 4}",
    "التقديري: تقسيم الـ 10 آلاف مصاريف ثابتة (قرار 27 أكتوبر) على الفئات تقديري — عدّله على الأرقام الحقيقية.", F_NOTE)

dv = DataValidation(type="list", formula1='"' + ",".join(CLASSES) + '"', allow_blank=True)
ws.add_data_validation(dv); dv.add(f"C{CAT_FIRST}:C{CAT_LAST}")
dv = DataValidation(type="list", formula1='"' + ",".join(DRIVERS) + '"', allow_blank=True)
ws.add_data_validation(dv); dv.add(f"D{CAT_FIRST}:D{CAT_LAST}")
ws.freeze_panes = "C4"

# ================================================================ 3. Ledger (الحركات)
ws = sheet(S_TX)
ws.sheet_properties.tabColor = MUSTARD
title(ws, "الحركات — دفتر كل الفلوس",
      "سطر لكل حركة. البوت بيكتب هنا. المبلغ دايماً موجب، والفئة هي اللي بتحدد داخل ولا خارج.")
TX_HEAD = ["رقم", "التاريخ", "الفئة", "التفاصيل", "الكمية", "المبلغ (ج)", "الشخص", "طريقة الدفع",
           "المصدر", "وقت التسجيل", "ملاحظات"]
head(ws, 4, 1, TX_HEAD)
widths(ws, {"A": 7, "B": 12, "C": 18, "D": 34, "E": 8, "F": 13, "G": 10, "H": 13, "I": 8, "J": 18, "K": 26})
for r in range(TX_FIRST, 3001):
    ws[f"B{r}"].number_format = DATE
    ws[f"F{r}"].number_format = NUM
    ws[f"J{r}"].number_format = "yyyy-mm-dd hh:mm"
dv_cat = DataValidation(type="list", formula1=R_CATS, allow_blank=True,
                        showErrorMessage=True, error="الفئة لازم تكون من قايمة الإعدادات", errorTitle="فئة غلط")
ws.add_data_validation(dv_cat); dv_cat.add(f"C{TX_FIRST}:C{TX_LAST}")
dv = DataValidation(type="list", formula1=R_PNAMES, allow_blank=True)
ws.add_data_validation(dv); dv.add(f"G{TX_FIRST}:G{TX_LAST}")
dv = DataValidation(type="list", formula1='"' + ",".join(METHODS) + '"', allow_blank=True)
ws.add_data_validation(dv); dv.add(f"H{TX_FIRST}:H{TX_LAST}")
dv = DataValidation(type="list", formula1='"' + ",".join(SOURCES) + '"', allow_blank=True)
ws.add_data_validation(dv); dv.add(f"I{TX_FIRST}:I{TX_LAST}")
dv = DataValidation(type="decimal", operator="greaterThan", formula1="0", allow_blank=True,
                    showErrorMessage=True, error="المبلغ لازم يكون رقم موجب", errorTitle="مبلغ غلط")
ws.add_data_validation(dv); dv.add(f"F{TX_FIRST}:F{TX_LAST}")
ws.freeze_panes = "A5"
ws.auto_filter.ref = f"A4:K{TX_LAST}"

# ================================================================ 4. Startup (التأسيس)
ws = sheet(S_START)
title(ws, "مصاريف التأسيس — المعدات والتجهيز والإطلاق",
      "الأسعار تقديرية للبداية (خانات صفرا) — بدّلها بعروض الأسعار الحقيقية. كل بند بيتضاف للميزانية في شهر صرفه.")
widths(ws, {"A": 2, "B": 34, "C": 18, "D": 8, "E": 13, "F": 14, "G": 12, "H": 10, "I": 34})
head(ws, 5, 2, ["البند", "الفئة", "الكمية", "سعر الوحدة (ج)", "الإجمالي (ج)", "شهر الصرف", "اتشترى؟", "ملاحظات"])
ST_FIRST, ST_LAST = 6, 60
START_ITEMS = [
    ("ثلاجة كبيرة", "معدات وأدوات", 1, 25000, ""),
    ("ديب فريزر", "معدات وأدوات", 1, 18000, ""),
    ("بوتاجاز 5 شعلة", "معدات وأدوات", 1, 12000, ""),
    ("صاج/جريل بلانشا", "معدات وأدوات", 1, 8000, "للفراخ والسماش برجر"),
    ("إير فراير كبيرة", "معدات وأدوات", 1, 5000, "للبطاطس"),
    ("ميكروويف", "معدات وأدوات", 1, 4500, ""),
    ("كبة/محضر طعام", "معدات وأدوات", 1, 3500, ""),
    ("ميزان مطبخ ديجيتال", "معدات وأدوات", 3, 400, "Weighed, not guessed — ميزان في كل ستيشن"),
    ("ميزان تغليف", "معدات وأدوات", 1, 1500, ""),
    ("طقم حلل وطاسات", "معدات وأدوات", 1, 5000, ""),
    ("سكاكين وألواح تقطيع", "معدات وأدوات", 1, 2000, ""),
    ("ترابيزة ستانلس", "معدات وأدوات", 2, 4500, ""),
    ("أرفف تخزين", "معدات وأدوات", 2, 2000, ""),
    ("علب حفظ وتخزين", "معدات وأدوات", 1, 2000, ""),
    ("طابعة ليبلز", "معدات وأدوات", 1, 3500, "لليبل الماكروز على العلبة"),
    ("شفاط", "معدات وأدوات", 1, 5000, ""),
    ("طفاية حريق", "معدات وأدوات", 2, 800, "مطلوبة غالباً للترخيص"),
    ("مقدم وتأمين الشقة", "تجهيز المكان", 1, 9000, "شهرين تأمين — حسب العقد"),
    ("تشطيب ودهان المطبخ", "تجهيز المكان", 1, 8000, ""),
    ("كهرباء وسباكة", "تجهيز المكان", 1, 4000, ""),
    ("سجل تجاري وبطاقة ضريبية", "رسوم وتراخيص", 1, 3000, ""),
    ("شهادات صحية وترخيص", "رسوم وتراخيص", 1, 2000, ""),
    ("مخزون خامات أولي", "خامات وبضاعة", 1, 8000, "بضاعة متوازنة مش كتير (خطة الإطلاق)"),
    ("تغليف أولي (علب، شنط، ليبلز)", "تغليف", 1, 6000, ""),
    ("تصوير الإطلاق وتذوق المؤثرين", "تصوير ومحتوى", 1, 5000, ""),
    ("هدايا أول أسبوع (عصير/بطاطس)", "إعلانات ممولة", 1, 3000, "هدية مع كل أوردر أول أسبوع"),
]
for i in range(ST_LAST - ST_FIRST + 1):
    r = ST_FIRST + i
    item = START_ITEMS[i] if i < len(START_ITEMS) else (None, None, None, None, None)
    inp(ws, f"B{r}", item[0]); inp(ws, f"C{r}", item[1]); inp(ws, f"D{r}", item[2], NUM)
    inp(ws, f"E{r}", item[3], NUM)
    put(ws, f"F{r}", f'=IF(B{r}="",0,D{r}*E{r})', fmt=NUM, border=True)
    inp(ws, f"G{r}", f"={C_START}" if item[0] else None, MONTH)
    inp(ws, f"H{r}", "لا" if item[0] else None)
    put(ws, f"I{r}", item[4], F_NOTE, border=True)
    if item[0]:
        ws[f"G{r}"].font = F_LINK
dv = DataValidation(type="list", formula1=R_CATS, allow_blank=True)
ws.add_data_validation(dv); dv.add(f"C{ST_FIRST}:C{ST_LAST}")
dv = DataValidation(type="list", formula1='"نعم,لا"', allow_blank=True)
ws.add_data_validation(dv); dv.add(f"H{ST_FIRST}:H{ST_LAST}")
# summary block K..L
widths(ws, {"J": 3, "K": 36, "L": 16})
head(ws, 5, 11, ["الملخص", "ج"])
summary = [
    ("إجمالي التأسيس المخطط", f"=SUM(F{ST_FIRST}:F{ST_LAST})"),
    ("اتشترى بالفعل", f'=SUMIF(H{ST_FIRST}:H{ST_LAST},"نعم",F{ST_FIRST}:F{ST_LAST})'),
    ("لسه هيتشترى", "=L6-L7"),
    ("رأس المال المتفق عليه", f"={C_TOTCAP}"),
    ("احتياطي الطوارئ", f"={C_RESERVE}"),
    ("المتبقي للتشغيل بعد التأسيس والاحتياطي", "=L9-L6-L10"),
    ("المصروف الفعلي على فئات التأسيس (من الحركات)",
     f'=SUMPRODUCT(SUMIFS({R_AMT},{R_CAT},{R_CATS})*({R_CLASS}="تأسيس"))'),
]
for i, (lab, f) in enumerate(summary):
    r = 6 + i
    put(ws, f"K{r}", lab, F_BOLD if i in (0, 5) else F_BASE, border=True)
    put(ws, f"L{r}", f, F_BOLD if i in (0, 5) else F_BASE, fmt=NUM, border=True)
    if i in (3, 4):
        ws[f"L{r}"].font = F_LINK
put(ws, "K14", "لو الرقم الأخير بالسالب يبقى رأس المال مش مكفّي التأسيس والاحتياطي.", F_NOTE)
put(ws, "K15", "علّم «اتشترى؟ = نعم» لما تشتري البند، وسجّل الفلوس نفسها في الحركات (أو ابعتها للبوت).", F_NOTE)
ws.freeze_panes = "B6"

# ================================================================ 5. Salaries
ws = sheet(S_SAL)
title(ws, "المرتبات", "عمر وأحمد من غير مرتب في الأول. أي موظف جديد: اكتب مرتبه وأول شهر، والميزانية هتحسبه لوحدها.")
SAL_MCOL = [L(7 + i) for i in range(N_MONTHS)]   # G..R
widths(ws, {"A": 2, "B": 26, "C": 22, "D": 12, "E": 12, "F": 12})
for c in SAL_MCOL + [L(7 + N_MONTHS)]:
    ws.column_dimensions[c].width = 11
head(ws, 5, 2, ["الاسم", "الدور", "المرتب الشهري", "أول شهر", "آخر شهر"])
for i, c in enumerate(SAL_MCOL):
    put(ws, f"{c}5", f"=EDATE({C_START},{i})", F_HEAD, HEAD_FILL, MONTH, CENTER, True)
put(ws, f"{L(7 + N_MONTHS)}5", "الإجمالي", F_HEAD, HEAD_FILL, align=CENTER, border=True)
SAL_FIRST, SAL_LAST = 6, 15
STAFF = [
    ("عمر", "إدارة وتشغيل وطبخ", 0, START_MONTH, None),
    ("أحمد", "موقع وتوصيل ومشتريات", 0, START_MONTH, None),
    ("مساعد مطبخ", "لما الأوردرات تزيد", 5000, None, None),
    ("طيار توصيل", "لو احتجنا", 4000, None, None),
]
for i in range(SAL_LAST - SAL_FIRST + 1):
    r = SAL_FIRST + i
    s = STAFF[i] if i < len(STAFF) else (None,) * 5
    inp(ws, f"B{r}", s[0]); inp(ws, f"C{r}", s[1]); inp(ws, f"D{r}", s[2], NUM)
    inp(ws, f"E{r}", s[3], MONTH); inp(ws, f"F{r}", s[4], MONTH)
    for c in SAL_MCOL:
        put(ws, f"{c}{r}",
            f'=IF(OR($D{r}="",$E{r}=""),0,IF(AND({c}$5>=$E{r},OR($F{r}="",{c}$5<=$F{r})),$D{r},0))',
            fmt=NUM, border=True)
    put(ws, f"{L(7 + N_MONTHS)}{r}", f"=SUM({SAL_MCOL[0]}{r}:{SAL_MCOL[-1]}{r})", F_BOLD, fmt=NUM, border=True)
SAL_TOT = SAL_LAST + 1
put(ws, f"B{SAL_TOT}", "إجمالي المرتبات", F_BOLD, TOTAL_FILL, border=True)
for c in SAL_MCOL + [L(7 + N_MONTHS)]:
    put(ws, f"{c}{SAL_TOT}", f"=SUM({c}{SAL_FIRST}:{c}{SAL_LAST})", F_BOLD, TOTAL_FILL, NUM, border=True)
put(ws, f"B{SAL_TOT + 2}", "مساعد المطبخ والطيار مش هيتحسبوا غير لما تكتب «أول شهر». المرتبات تقديرية.", F_NOTE)
ws.freeze_panes = "G6"

# ================================================================ 6. Budget
ws = sheet(S_BUD)
title(ws, "الميزانية — المخطط قصاد الفعلي",
      "المخطط بيتحسب من الإعدادات والتأسيس والمرتبات. الفعلي بيتحسب لوحده من الحركات. عدّل الأوردرات المتوقعة وأيام الشغل (أصفر).")
widths(ws, {"A": 2, "B": 30, "C": 14})
for c in MCOLS:
    ws.column_dimensions[c].width = 12
ws.column_dimensions[TOTCOL].width = 14
head(ws, 5, 2, ["الشهر", ""])
for i, c in enumerate(MCOLS):
    put(ws, f"{c}5", f"={C_START}" if i == 0 else f"=EDATE({MCOLS[0]}5,{i})", F_HEAD, HEAD_FILL, MONTH, CENTER, True)
put(ws, f"{TOTCOL}5", "الإجمالي", F_HEAD, HEAD_FILL, align=CENTER, border=True)

ORDERS_PER_DAY = [6, 15, 19, 23, 28, 32, 36, 40, 45, 49, 53, 58]   # base case 15 -> ~60 (business model)
put(ws, "B6", "أيام الشغل", F_BOLD, border=True)
put(ws, "B7", "الأوردرات المتوقعة في اليوم", F_BOLD, border=True)
put(ws, "B8", "إجمالي الأوردرات المتوقعة", F_BOLD, border=True)
put(ws, "B9", "متوسط قيمة الأوردر", F_BOLD, border=True)
put(ws, "B10", "المبيعات المتوقعة (ج)", F_BOLD, TOTAL_FILL, border=True)
for i, c in enumerate(MCOLS):
    if i == 0:
        inp(ws, f"{c}6", 5, NUM, "الإطلاق آخر نوفمبر، فأيام شغل قليلة في أول شهر.")
    else:
        put(ws, f"{c}6", f"={C_DAYS}", F_LINK, INPUT_FILL, NUM, border=True)
    inp(ws, f"{c}7", ORDERS_PER_DAY[i], NUM)
    put(ws, f"{c}8", f"={c}6*{c}7", fmt=NUM, border=True)
    put(ws, f"{c}9", f"={C_AVG}", F_LINK, fmt=NUM, border=True)
    put(ws, f"{c}10", f"={c}8*{c}9", F_BOLD, TOTAL_FILL, NUM, border=True)
for r in (8, 10):
    put(ws, f"{TOTCOL}{r}", f"=SUM({MCOLS[0]}{r}:{MCOLS[-1]}{r})", F_BOLD, TOTAL_FILL, NUM, border=True)
ws["D7"].comment = Comment("الحالة الأساسية من نموذج البيزنس: 15 أوردر/يوم في ديسمبر لحد ~60 بعد سنة. عدّلها.", "MAZBOOT")

BP0, BA0, BV0 = 13, 42, 71      # first category row of planned / actual / remaining blocks
STARTR = q(S_START)
SALR = q(S_SAL)


def block_header(r, text):
    put(ws, f"B{r}", text, F_HEAD, PatternFill("solid", fgColor=DEEP), border=True)
    put(ws, f"C{r}", "التصنيف", F_HEAD, PatternFill("solid", fgColor=DEEP), align=CENTER, border=True)
    for c in MCOLS + [TOTCOL]:
        put(ws, f"{c}{r}", f"={c}$5" if c != TOTCOL else "الإجمالي",
            F_HEAD, PatternFill("solid", fgColor=DEEP), MONTH, CENTER, True)


block_header(BP0 - 1, "المخطط (الميزانية)")
block_header(BA0 - 1, "الفعلي (من الحركات)")
block_header(BV0 - 1, "المتبقي من الميزانية (المخطط − الفعلي)")
for i in range(N_CAT):
    s = CAT_FIRST + i
    rp, ra, rv = BP0 + i, BA0 + i, BV0 + i
    for r in (rp, ra, rv):
        put(ws, f"B{r}", f'=IF({SET}!$B${s}="","",{SET}!$B${s})', F_LINK, border=True)
        put(ws, f"C{r}", f'=IF({SET}!$C${s}="","",{SET}!$C${s})', F_NOTE, border=True)
    drv, val = f"{SET}!$D${s}", f"{SET}!$E${s}"
    for j, c in enumerate(MCOLS):
        salc = SAL_MCOL[j]
        planned = (
            f'=IF($B{rp}="",0,'
            f'IF({drv}="ثابت شهري",{val},'
            f'IF({drv}="% من المبيعات",{val}*{c}$10,'
            f'IF({drv}="لكل أوردر",{val}*{c}$8,'
            f'IF({drv}="المبيعات المتوقعة",{c}$10,'
            f'IF({drv}="من تاب المرتبات",{SALR}!{salc}${SAL_TOT},'
            f'IF({drv}="رأس مال الشركاء",IF({c}$5={MCOLS[0]}$5,{C_TOTCAP},0),0))))))'
            f'+SUMIFS({STARTR}!$F${ST_FIRST}:$F${ST_LAST},{STARTR}!$C${ST_FIRST}:$C${ST_LAST},$B{rp},'
            f'{STARTR}!$G${ST_FIRST}:$G${ST_LAST},">="&{c}$5,{STARTR}!$G${ST_FIRST}:$G${ST_LAST},"<="&EOMONTH({c}$5,0)))'
        )
        put(ws, f"{c}{rp}", planned, fmt=NUM, border=True)
        put(ws, f"{c}{ra}", f'=IF($B{ra}="",0,SUMIFS({R_AMT},{R_CAT},$B{ra},{month_filter(c)}))', fmt=NUM, border=True)
        put(ws, f"{c}{rv}", f"={c}{rp}-{c}{ra}", fmt=NUM, border=True)
    for r in (rp, ra, rv):
        put(ws, f"{TOTCOL}{r}", f"=SUM({MCOLS[0]}{r}:{MCOLS[-1]}{r})", F_BOLD, fmt=NUM, border=True)

BP_END, BA_END, BV_END = BP0 + N_CAT - 1, BA0 + N_CAT - 1, BV0 + N_CAT - 1


def class_sum(col, r0, r1, classes):
    return "+".join(f'SUMIF($C${r0}:$C${r1},"{k}",{col}{r0}:{col}{r1})' for k in classes)


EXP = ["تكلفة مباشرة", "تشغيل", "تأسيس"]
for (r0, r1, lab) in ((BP0, BP_END, "إجمالي المصاريف المخططة"), (BA0, BA_END, "إجمالي المصاريف الفعلية"),
                      (BV0, BV_END, "إجمالي المتبقي من ميزانية المصاريف")):
    rt = r1 + 1
    put(ws, f"B{rt}", lab, F_BOLD, TOTAL_FILL, border=True)
    put(ws, f"C{rt}", "", fill=TOTAL_FILL, border=True)
    for c in MCOLS + [TOTCOL]:
        put(ws, f"{c}{rt}", "=" + class_sum(c, r0, r1, EXP), F_BOLD, TOTAL_FILL, NUM, border=True)
FIXED_ROW = BP_END + 2
put(ws, f"B{FIXED_ROW}", "التكاليف الثابتة المخططة (تشغيل ثابت + مرتبات)", F_BOLD, border=True)
for j, c in enumerate(MCOLS):
    put(ws, f"{c}{FIXED_ROW}",
        f'=SUMPRODUCT(({R_CLASS}="تشغيل")*({R_DRV}="ثابت شهري"),{R_VAL})+{SALR}!{SAL_MCOL[j]}${SAL_TOT}',
        fmt=NUM, border=True)
BE_ROW = FIXED_ROW + 1
put(ws, f"B{BE_ROW}", "نقطة التعادل (أوردر في اليوم)", F_BOLD, border=True)
for c in MCOLS:
    put(ws, f"{c}{BE_ROW}", f"=IFERROR({c}{FIXED_ROW}/{C_CONTRIB}/{c}$6,0)", fmt=NUM2, border=True)

# conditional formatting on remaining block (expense rows only)
rng = f"{MCOLS[0]}{BV0}:{TOTCOL}{BV_END}"
red_fill = PatternFill("solid", fgColor="F8D3CF")
amber_fill = PatternFill("solid", fgColor="FFE8B0")
ws.conditional_formatting.add(rng, FormulaRule(
    formula=[f'AND(OR($C{BV0}="تكلفة مباشرة",$C{BV0}="تشغيل",$C{BV0}="تأسيس"),{MCOLS[0]}{BV0}<0)'],
    fill=red_fill, font=Font(name=FONT, color="9C0006", bold=True)))
ws.conditional_formatting.add(rng, FormulaRule(
    formula=[f'AND(OR($C{BV0}="تكلفة مباشرة",$C{BV0}="تشغيل",$C{BV0}="تأسيس"),'
             f'{MCOLS[0]}{BA0}>0,{MCOLS[0]}{BA0}>={MCOLS[0]}{BP0}*{C_ALERT})'],
    fill=amber_fill))
ws.freeze_panes = "D6"

# ================================================================ 7. P&L
ws = sheet(S_PL)
title(ws, "قائمة الدخل (الأرباح والخسائر) — فعلي", "كل الأرقام من الحركات. مصاريف التأسيس بتظهر تحت كبند منفصل لأنها استثمار مش مصروف شهري.")
widths(ws, {"A": 2, "B": 36, "C": 4})
for c in MCOLS:
    ws.column_dimensions[c].width = 12
ws.column_dimensions[TOTCOL].width = 14
BUD = q(S_BUD)
head(ws, 5, 2, ["البند", ""])
for c in MCOLS:
    put(ws, f"{c}5", f"={BUD}!{c}5", F_HEAD, HEAD_FILL, MONTH, CENTER, True)
put(ws, f"{TOTCOL}5", "الإجمالي", F_HEAD, HEAD_FILL, align=CENTER, border=True)


def bud_class(c, block0, block1, cls):
    return f'SUMIF({BUD}!$C${block0}:$C${block1},"{cls}",{BUD}!{c}{block0}:{c}{block1})'


def act_row(cat):
    return BA0 + CAT_ROW[cat] - CAT_FIRST


pl_rows = []   # (label, kind, formula_fn) ; kind: section/line/total/pct/memo


def line(label, fn, style="line", fmt=NUM, total="sum"):
    pl_rows.append((label, style, fn, fmt, total))


line("عدد الأوردرات", lambda c: f'=SUMIFS({R_QTY},{R_CAT},{SALES_CAT},{month_filter(c)})', "line")
line("الإيرادات", None, "section")
line("مبيعات", lambda c: f"={BUD}!{c}{act_row('مبيعات')}")
line("إيرادات أخرى", lambda c: f"={BUD}!{c}{act_row('إيرادات أخرى')}")
line("إجمالي الإيرادات", lambda c: "=" + bud_class(c, BA0, BA_END, "إيراد"), "total")
line("متوسط قيمة الأوردر", None, "calc_avg", NUM, "calc")
line("التكلفة المباشرة", None, "section")
line("خامات وبضاعة", lambda c: f"={BUD}!{c}{act_row('خامات وبضاعة')}")
line("تغليف", lambda c: f"={BUD}!{c}{act_row('تغليف')}")
line("إجمالي التكلفة المباشرة", lambda c: "=" + bud_class(c, BA0, BA_END, "تكلفة مباشرة"), "total")
line("مجمل الربح", None, "calc_gp", NUM, "calc")
line("هامش مجمل الربح %", None, "calc_gpm", PCT, "calc")
line("نسبة تكلفة الأكل من المبيعات %", None, "calc_food", PCT, "calc")
line("مصاريف التشغيل", None, "section")
OPEX_CATS = [c[0] for c in CATEGORIES if c[1] == "تشغيل"]
for cat in OPEX_CATS:
    line(cat, (lambda cc: (lambda c: f"={BUD}!{c}{act_row(cc)}"))(cat))
line("إجمالي مصاريف التشغيل", lambda c: "=" + bud_class(c, BA0, BA_END, "تشغيل"), "total")
line("صافي الربح التشغيلي", None, "calc_np", NUM, "calc")
line("هامش صافي الربح %", None, "calc_npm", PCT, "calc")
line("صافي الربح التراكمي", None, "calc_cum", NUM, "calc")
line("بنود للمتابعة", None, "section")
line("مصاريف تأسيس (استثمار)", lambda c: "=" + bud_class(c, BA0, BA_END, "تأسيس"), "line")
line("الصافي بعد التأسيس", None, "calc_after", NUM, "calc")
line("المبيعات المستهدفة (من الميزانية)", lambda c: f"={BUD}!{c}10", "line")
line("الفرق عن المستهدف", None, "calc_vs", NUM, "calc")

PL = {}
r = 6
for label, style, fn, fmt, total in pl_rows:
    PL[label] = r
    r += 1
for label, style, fn, fmt, total in pl_rows:
    r = PL[label]
    if style == "section":
        put(ws, f"B{r}", label, F_BOLD, SUB_FILL, border=True)
        for c in MCOLS + [TOTCOL]:
            put(ws, f"{c}{r}", None, fill=SUB_FILL, border=True)
        continue
    bold = style in ("total",) or label in ("مجمل الربح", "صافي الربح التشغيلي")
    fill = TOTAL_FILL if bold else None
    put(ws, f"B{r}", label, F_BOLD if bold else F_BASE, fill, border=True)
    for j, c in enumerate(MCOLS + [TOTCOL]):
        isT = c == TOTCOL
        P = lambda k: f"{c}{PL[k]}"
        if fn is not None:
            f = f"=SUM({MCOLS[0]}{r}:{MCOLS[-1]}{r})" if isT else fn(c)
        elif style == "calc_avg":
            f = f"=IFERROR({P('مبيعات')}/{P('عدد الأوردرات')},0)"
        elif style == "calc_gp":
            f = f"={P('إجمالي الإيرادات')}-{P('إجمالي التكلفة المباشرة')}"
        elif style == "calc_gpm":
            f = f"=IFERROR({P('مجمل الربح')}/{P('إجمالي الإيرادات')},0)"
        elif style == "calc_food":
            f = f"=IFERROR({P('خامات وبضاعة')}/{P('مبيعات')},0)"
        elif style == "calc_np":
            f = f"={P('مجمل الربح')}-{P('إجمالي مصاريف التشغيل')}"
        elif style == "calc_npm":
            f = f"=IFERROR({P('صافي الربح التشغيلي')}/{P('إجمالي الإيرادات')},0)"
        elif style == "calc_cum":
            npr = PL["صافي الربح التشغيلي"]
            f = f"={TOTCOL}{npr}" if isT else f"=SUM(${MCOLS[0]}{npr}:{c}{npr})"
        elif style == "calc_after":
            f = f"={P('صافي الربح التشغيلي')}-{P('مصاريف تأسيس (استثمار)')}"
        elif style == "calc_vs":
            f = f"={P('مبيعات')}-{P('المبيعات المستهدفة (من الميزانية)')}"
        put(ws, f"{c}{r}", f, F_BOLD if bold else F_BASE, fill, fmt, border=True)
put(ws, f"B{max(PL.values()) + 2}",
    "الإجماليات بتشمل أي فئة جديدة تضيفها في الإعدادات حتى لو مش ظاهرة كسطر هنا.", F_NOTE)
ws.conditional_formatting.add(f"{MCOLS[0]}{PL['صافي الربح التشغيلي']}:{TOTCOL}{PL['صافي الربح التشغيلي']}",
                              CellIsRule(operator="lessThan", formula=["0"], font=Font(name=FONT, color="C00000", bold=True)))
ws.freeze_panes = "D6"

# ================================================================ 8. Cash flow
ws = sheet(S_CF)
title(ws, "التدفق النقدي", "الفعلي من الحركات + المتوقع من الميزانية. أهم رقم: رصيد آخر الشهر — لو قرّب من الاحتياطي يبقى فيه مشكلة.")
widths(ws, {"A": 2, "B": 38, "C": 4})
for c in MCOLS:
    ws.column_dimensions[c].width = 12
ws.column_dimensions[TOTCOL].width = 14
head(ws, 5, 2, ["البند", ""])
for c in MCOLS:
    put(ws, f"{c}5", f"={BUD}!{c}5", F_HEAD, HEAD_FILL, MONTH, CENTER, True)
put(ws, f"{TOTCOL}5", "الإجمالي", F_HEAD, HEAD_FILL, align=CENTER, border=True)


PRE = (f'SUMPRODUCT(SUMIFS({R_AMT},{R_CAT},{R_CATS},{R_DATE},"<"&{C_START})*{sign_expr()})'
       f'+SUMIFS({R_AMT},{R_MET},"من جيبه",{R_DATE},"<"&{C_START})')
cf = [
    ("الفعلي", "section", None),
    ("رصيد أول الشهر", "open", None),
    ("رأس مال دفعه الشركاء", "in", lambda c: bud_class(c, BA0, BA_END, "تمويل")),
    ("مبيعات وإيرادات", "in", lambda c: bud_class(c, BA0, BA_END, "إيراد")),
    ("مصاريف دفعها شريك من جيبه", "in", lambda c: f'SUMIFS({R_AMT},{R_MET},"من جيبه",{month_filter(c)})'),
    ("إجمالي الداخل", "tin", None),
    ("تكلفة مباشرة (أكل وتغليف)", "out", lambda c: bud_class(c, BA0, BA_END, "تكلفة مباشرة")),
    ("مصاريف تشغيل", "out", lambda c: bud_class(c, BA0, BA_END, "تشغيل")),
    ("تأسيس ومعدات", "out", lambda c: bud_class(c, BA0, BA_END, "تأسيس")),
    ("سحوبات الشركاء", "out", lambda c: bud_class(c, BA0, BA_END, "سحب")),
    ("إجمالي الخارج", "tout", None),
    ("صافي حركة الشهر", "net", None),
    ("رصيد آخر الشهر", "close", None),
    ("المتوقع (من الميزانية)", "section", None),
    ("رصيد أول الشهر المتوقع", "popen", None),
    ("رأس المال المتوقع", "pin", lambda c: bud_class(c, BP0, BP_END, "تمويل")),
    ("المبيعات المتوقعة", "pin", lambda c: bud_class(c, BP0, BP_END, "إيراد")),
    ("المصاريف المتوقعة", "pout", lambda c: f"{BUD}!{c}{BP_END + 1}"),
    ("السحوبات المتوقعة", "pout", lambda c: bud_class(c, BP0, BP_END, "سحب")),
    ("رصيد آخر الشهر المتوقع", "pclose", None),
    ("فوق الاحتياطي بكام؟", "pres", None),
]
CF = {lab: 6 + i for i, (lab, _, _) in enumerate(cf)}
for lab, kind, fn in cf:
    r = CF[lab]
    if kind == "section":
        put(ws, f"B{r}", lab, F_HEAD, PatternFill("solid", fgColor=DEEP), border=True)
        for c in MCOLS + [TOTCOL]:
            put(ws, f"{c}{r}", None, fill=PatternFill("solid", fgColor=DEEP), border=True)
        continue
    bold = kind in ("tin", "tout", "close", "pclose", "net")
    fill = TOTAL_FILL if bold else None
    put(ws, f"B{r}", lab, F_BOLD if bold else F_BASE, fill, border=True)
    for j, c in enumerate(MCOLS):
        prev = MCOLS[j - 1] if j else None
        if kind in ("in", "out", "pin", "pout"):
            f = "=" + fn(c)
        elif kind == "open":
            f = f"={C_OPEN}+{PRE}" if j == 0 else f"={prev}{CF['رصيد آخر الشهر']}"
        elif kind == "tin":
            f = f"=SUM({c}{CF['رأس مال دفعه الشركاء']}:{c}{CF['مصاريف دفعها شريك من جيبه']})"
        elif kind == "tout":
            f = f"=SUM({c}{CF['تكلفة مباشرة (أكل وتغليف)']}:{c}{CF['سحوبات الشركاء']})"
        elif kind == "net":
            f = f"={c}{CF['إجمالي الداخل']}-{c}{CF['إجمالي الخارج']}"
        elif kind == "close":
            f = f"={c}{CF['رصيد أول الشهر']}+{c}{CF['صافي حركة الشهر']}"
        elif kind == "popen":
            f = f"={C_OPEN}" if j == 0 else f"={prev}{CF['رصيد آخر الشهر المتوقع']}"
        elif kind == "pclose":
            f = (f"={c}{CF['رصيد أول الشهر المتوقع']}+{c}{CF['رأس المال المتوقع']}+{c}{CF['المبيعات المتوقعة']}"
                 f"-{c}{CF['المصاريف المتوقعة']}-{c}{CF['السحوبات المتوقعة']}")
        elif kind == "pres":
            f = f"={c}{CF['رصيد آخر الشهر المتوقع']}-{C_RESERVE}"
        put(ws, f"{c}{r}", f, F_BOLD if bold else F_BASE, fill, NUM, border=True)
    if kind in ("in", "out", "pin", "pout", "tin", "tout", "net"):
        put(ws, f"{TOTCOL}{r}", f"=SUM({MCOLS[0]}{r}:{MCOLS[-1]}{r})", F_BOLD, fill, NUM, border=True)
    elif kind in ("close", "pclose", "pres"):
        put(ws, f"{TOTCOL}{r}", f"={MCOLS[-1]}{r}", F_BOLD, fill, NUM, border=True)
    else:
        put(ws, f"{TOTCOL}{r}", None, border=True)
for k in ("رصيد آخر الشهر", "رصيد آخر الشهر المتوقع", "فوق الاحتياطي بكام؟"):
    ws.conditional_formatting.add(f"{MCOLS[0]}{CF[k]}:{TOTCOL}{CF[k]}",
                                  CellIsRule(operator="lessThan", formula=["0"], fill=red_fill,
                                             font=Font(name=FONT, color="9C0006", bold=True)))
put(ws, f"B{max(CF.values()) + 2}",
    "رصيد أول شهر بيشمل أي حركات متسجلة قبل بداية الخطة. «من جيبه» = شريك دفع مصروف من فلوسه، فبيتحسب داخل وخارج مع بعض.", F_NOTE)
ws.freeze_panes = "D6"

# ================================================================ 9. Partners
ws = sheet(S_PART)
title(ws, "حسابات الشركاء", "كل رقم هنا من الحركات + النسب من الإعدادات. صافي الحساب = رأس المال + من جيبه + نصيب الربح − السحوبات.")
widths(ws, {"A": 2, "B": 14, "C": 16, "D": 16, "E": 14, "F": 12, "G": 16, "H": 14, "I": 18, "J": 16})
head(ws, 5, 2, ["الشريك", "رأس المال المتفق عليه", "رأس المال المدفوع", "المتبقي عليه", "نسبة الملكية",
                "مصاريف دفعها من جيبه", "السحوبات", "نصيبه من صافي الربح", "صافي حسابه"])
CFQ = q(S_CF)
PLQ = q(S_PL)
NET_PROFIT_ALL = f"{PLQ}!${TOTCOL}${PL['صافي الربح التشغيلي']}"
for i in range(PART_LAST - PART_FIRST + 1):
    r = 6 + i
    s = PART_FIRST + i
    put(ws, f"B{r}", f'=IF({SET}!$B${s}="","",{SET}!$B${s})', F_LINK, border=True)
    put(ws, f"C{r}", f"={SET}!$C${s}", F_LINK, fmt=NUM, border=True)
    put(ws, f"D{r}", f'=IF($B{r}="",0,SUMIFS({R_AMT},{R_CAT},{CAPITAL_CAT},{R_PER},$B{r}))', fmt=NUM, border=True)
    put(ws, f"E{r}", f"=C{r}-D{r}", fmt=NUM, border=True)
    put(ws, f"F{r}", f"={SET}!$D${s}", F_LINK, fmt=PCT, border=True)
    put(ws, f"G{r}", f'=IF($B{r}="",0,SUMIFS({R_AMT},{R_MET},"من جيبه",{R_PER},$B{r}))', fmt=NUM, border=True)
    put(ws, f"H{r}", f'=IF($B{r}="",0,SUMIFS({R_AMT},{R_CAT},{WITHDRAW_CAT},{R_PER},$B{r}))', fmt=NUM, border=True)
    put(ws, f"I{r}", f"={NET_PROFIT_ALL}*F{r}", fmt=NUM, border=True)
    put(ws, f"J{r}", f"=D{r}+G{r}+I{r}-H{r}", F_BOLD, fmt=NUM, border=True)
rt = 6 + PART_LAST - PART_FIRST + 1
put(ws, f"B{rt}", "الإجمالي", F_BOLD, TOTAL_FILL, border=True)
for c in "CDEFGHIJ":
    put(ws, f"{c}{rt}", f"=SUM({c}6:{c}{rt - 1})", F_BOLD, TOTAL_FILL, PCT if c == "F" else NUM, border=True)
put(ws, f"B{rt + 2}", "نصيب الربح محسوب على صافي الربح التشغيلي من أول الخطة (قبل التأسيس). ده حساب دفتري — توزيع الأرباح الفعلي بيتسجل «سحب شريك».", F_NOTE)
put(ws, f"B{rt + 3}", "«مصاريف دفعها من جيبه» = مستحقة له من الشركة لحد ما تترجعله (سجّل الرجوع كـ «سحب شريك»).", F_NOTE)

# ================================================================ 10. Meal costing
ws = sheet(S_MEAL)
title(ws, "تكلفة الوجبات — الـ 7 أطباق",
      "الوصفات والأسعار هنا تقديرية للبداية (أصفر). حط الجرامات الحقيقية من menu master وأسعار الموردين.")
widths(ws, {"A": 2, "B": 30, "C": 10, "D": 9, "E": 12, "F": 3, "G": 2, "H": 26, "I": 12, "J": 9, "K": 30})
# ingredient prices (H..K)
head(ws, 5, 8, ["المكون", "السعر (ج)", "الوحدة", "ملاحظة"])
ING = [
    ("صدور فراخ", 230, "كيلو"), ("لحم مفروم 80:20", 450, "كيلو"), ("رز بسمتي", 70, "كيلو"),
    ("بطاطس", 20, "كيلو"), ("خس", 40, "كيلو"), ("طماطم", 20, "كيلو"), ("بصل", 20, "كيلو"),
    ("فلفل ألوان", 60, "كيلو"), ("ثوم", 80, "كيلو"), ("ليمون", 40, "كيلو"), ("زيت زيتون", 450, "كيلو"),
    ("زيت عادي", 90, "كيلو"), ("زبادي", 70, "كيلو"), ("جبنة شيدر", 400, "كيلو"), ("جبنة بارميزان", 900, "كيلو"),
    ("توابل وبهارات", 300, "كيلو"), ("صوص دايت", 120, "كيلو"), ("عيش برجر", 6, "قطعة"), ("تورتيلا", 7, "قطعة"),
    ("خبز توست (كروتون)", 60, "كيلو"), ("ذرة", 90, "كيلو"),
]
ING_FIRST, ING_LAST = 6, 40
for i in range(ING_LAST - ING_FIRST + 1):
    r = ING_FIRST + i
    it = ING[i] if i < len(ING) else (None, None, None)
    inp(ws, f"H{r}", it[0]); inp(ws, f"I{r}", it[1], NUM); inp(ws, f"J{r}", it[2])
    put(ws, f"K{r}", "تقديري" if it[0] else None, F_NOTE, border=True)
dv = DataValidation(type="list", formula1='"كيلو,قطعة"', allow_blank=True)
ws.add_data_validation(dv); dv.add(f"J{ING_FIRST}:J{ING_LAST}")
R_ING = f"$H${ING_FIRST}:$H${ING_LAST}"
DISHES = [
    ("OG Grilled Chicken Breast", 185, [("صدور فراخ", 250), ("رز بسمتي", 100), ("خس", 40), ("توابل وبهارات", 8), ("زيت زيتون", 10)]),
    ("Spicy Grilled Chicken Breast", 185, [("صدور فراخ", 250), ("رز بسمتي", 100), ("فلفل ألوان", 40), ("توابل وبهارات", 12), ("زيت زيتون", 10)]),
    ("Smash Burger 80:20", 195, [("لحم مفروم 80:20", 180), ("عيش برجر", 1), ("جبنة شيدر", 30), ("خس", 20), ("طماطم", 30), ("صوص دايت", 30)]),
    ("Tortillas Shawarma", 175, [("صدور فراخ", 200), ("تورتيلا", 2), ("زبادي", 50), ("ثوم", 10), ("خس", 30), ("توابل وبهارات", 10)]),
    ("Chicken Fajitas Bowl", 190, [("صدور فراخ", 200), ("رز بسمتي", 100), ("فلفل ألوان", 80), ("بصل", 40), ("ذرة", 30), ("توابل وبهارات", 10)]),
    ("Cubed Chicken & Potatoes", 180, [("صدور فراخ", 200), ("بطاطس", 250), ("زيت زيتون", 15), ("توابل وبهارات", 8), ("صوص دايت", 30)]),
    ("Caesar Salad", 170, [("صدور فراخ", 150), ("خس", 150), ("جبنة بارميزان", 20), ("خبز توست (كروتون)", 30), ("صوص دايت", 40), ("ليمون", 15)]),
]
# summary on top
head(ws, 5, 2, ["الطبق", "التكلفة", "السعر", "تكلفة %"])
SUM_FIRST = 6
N_ING_ROWS = 8
blocks_start = SUM_FIRST + len(DISHES) + 3
PACK_VAL = f"{SET}!$E${CAT_ROW['تغليف']}"
r = blocks_start
for k, (name, price, recipe) in enumerate(DISHES):
    hr = r
    put(ws, f"B{hr}", name, F_HEAD, HEAD_FILL, border=True)
    for c, lab in zip("CDE", ["الكمية", "الوحدة", "التكلفة (ج)"]):
        put(ws, f"{c}{hr}", lab, F_HEAD, HEAD_FILL, align=CENTER, border=True)
    for i in range(N_ING_ROWS):
        rr = hr + 1 + i
        ing = recipe[i] if i < len(recipe) else (None, None)
        inp(ws, f"B{rr}", ing[0]); inp(ws, f"C{rr}", ing[1], NUM)
        put(ws, f"D{rr}", f'=IF(B{rr}="","",IFERROR(INDEX($J${ING_FIRST}:$J${ING_LAST},MATCH(B{rr},{R_ING},0)),"؟"))',
            F_NOTE, border=True)
        put(ws, f"E{rr}",
            f'=IF(OR(B{rr}="",C{rr}=""),0,IFERROR(IF(D{rr}="قطعة",C{rr},C{rr}/1000)*'
            f'INDEX($I${ING_FIRST}:$I${ING_LAST},MATCH(B{rr},{R_ING},0)),0))', fmt=NUM2, border=True)
    dv = DataValidation(type="list", formula1=R_ING, allow_blank=True)
    ws.add_data_validation(dv); dv.add(f"B{hr + 1}:B{hr + N_ING_ROWS}")
    pr = hr + N_ING_ROWS + 1
    put(ws, f"B{pr}", "تكلفة المكونات", F_BOLD, border=True)
    put(ws, f"E{pr}", f"=SUM(E{hr + 1}:E{hr + N_ING_ROWS})", F_BOLD, fmt=NUM2, border=True)
    put(ws, f"B{pr + 1}", "تغليف (من الإعدادات)", border=True)
    put(ws, f"E{pr + 1}", f"={PACK_VAL}", F_LINK, fmt=NUM2, border=True)
    put(ws, f"B{pr + 2}", "إجمالي التكلفة", F_BOLD, TOTAL_FILL, border=True)
    put(ws, f"E{pr + 2}", f"=E{pr}+E{pr + 1}", F_BOLD, TOTAL_FILL, NUM2, border=True)
    put(ws, f"B{pr + 3}", "سعر البيع", F_BOLD, border=True)
    inp(ws, f"E{pr + 3}", price, NUM, "سعر تقديري — حط سعر المنيو الحقيقي")
    put(ws, f"B{pr + 4}", "نسبة تكلفة الأكل (بدون تغليف)", border=True)
    put(ws, f"E{pr + 4}", f"=IFERROR(E{pr}/E{pr + 3},0)", fmt=PCT, border=True)
    put(ws, f"B{pr + 5}", "المكسب من الطبق قبل المصاريف الثابتة", F_BOLD, border=True)
    put(ws, f"E{pr + 5}", f"=E{pr + 3}-E{pr + 2}", F_BOLD, fmt=NUM, border=True)
    sr = SUM_FIRST + k
    put(ws, f"B{sr}", f"=B{hr}", border=True)
    put(ws, f"C{sr}", f"=E{pr + 2}", fmt=NUM, border=True)
    put(ws, f"D{sr}", f"=E{pr + 3}", fmt=NUM, border=True)
    put(ws, f"E{sr}", f"=E{pr + 4}", fmt=PCT, border=True)
    r = pr + 7
avg_r = SUM_FIRST + len(DISHES)
put(ws, f"B{avg_r}", "المتوسط", F_BOLD, TOTAL_FILL, border=True)
put(ws, f"C{avg_r}", f"=AVERAGE(C{SUM_FIRST}:C{avg_r - 1})", F_BOLD, TOTAL_FILL, NUM, border=True)
put(ws, f"D{avg_r}", f"=AVERAGE(D{SUM_FIRST}:D{avg_r - 1})", F_BOLD, TOTAL_FILL, NUM, border=True)
put(ws, f"E{avg_r}", f"=AVERAGE(E{SUM_FIRST}:E{avg_r - 1})", F_BOLD, TOTAL_FILL, PCT, border=True)
put(ws, f"B{avg_r + 1}", "قارن متوسط تكلفة الأكل % بالنسبة اللي في الإعدادات (فئة خامات وبضاعة).", F_NOTE)
ws.conditional_formatting.add(f"E{SUM_FIRST}:E{avg_r}", CellIsRule(
    operator="greaterThan", formula=[f"{SET}!$E${CAT_ROW['خامات وبضاعة']}"], font=Font(name=FONT, color="C00000", bold=True)))

# ================================================================ 1. Dashboard (fill)
ws = ws_dash
ws.sheet_properties.tabColor = RED
title(ws, "MAZBOOT — لوحة التحكم المالية", "اختار الشهر من الخانة الصفرا. كل الأرقام بتتحدث لوحدها من الحركات.")
widths(ws, {"A": 2, "B": 24, "C": 16, "D": 16, "E": 16, "F": 14, "G": 12, "H": 18, "I": 2})
put(ws, "B4", "الشهر", F_BOLD)
inp(ws, "C4", "=DATE(YEAR(TODAY()),MONTH(TODAY()),1)", MONTH,
    "اختار من القايمة. الافتراضي هو الشهر الحالي.")
dv = DataValidation(type="list", formula1=f"{BUD}!$D$5:${MCOLS[-1]}$5", allow_blank=False)
ws.add_data_validation(dv); dv.add("C4")
put(ws, "D4", f'=IF(ISNA(MATCH(C4,{BUD}!$D$5:${MCOLS[-1]}$5,0)),"⚠️ الشهر ده برا الخطة — اختار من القايمة","")', F_NOTE)
put(ws, "H4", f"=IFERROR(MATCH(C4,{BUD}!$D$5:${MCOLS[-1]}$5,0),1)", F_NOTE)   # month index helper
ws["H4"].font = Font(name=FONT, size=8, color="FFFFFF")
IDX = "$H$4"


def at(sheet_name, row):
    return f"INDEX({q(sheet_name)}!$D${row}:${MCOLS[-1]}${row},{IDX})"


CF_CLOSE = CF["رصيد آخر الشهر"]
CASH_NOW = (f'{C_OPEN}+SUMPRODUCT(SUMIFS({R_AMT},{R_CAT},{R_CATS})*{sign_expr()})'
            f'+SUMIFS({R_AMT},{R_MET},"من جيبه")')
cards = [
    # row 6/7
    [("المبيعات", f"={at(S_PL, PL['مبيعات'])}", NUM),
     ("عدد الأوردرات", f"={at(S_PL, PL['عدد الأوردرات'])}", NUM),
     ("متوسط الأوردر", f"={at(S_PL, PL['متوسط قيمة الأوردر'])}", NUM),
     ("تكلفة الأكل %", f"={at(S_PL, PL['نسبة تكلفة الأكل من المبيعات %'])}", PCT),
     ("صافي الربح التشغيلي", f"={at(S_PL, PL['صافي الربح التشغيلي'])}", NUM),
     ("الكاش آخر الشهر", f"={at(S_CF, CF_CLOSE)}", NUM)],
    # row 9/10
    [("المصروف الفعلي", f"={at(S_BUD, BA_END + 1)}", NUM),
     ("ميزانية المصاريف", f"={at(S_BUD, BP_END + 1)}", NUM),
     ("المتبقي من الميزانية", "=C10-B10", NUM),
     ("نقطة التعادل (أوردر/يوم)", f"={at(S_BUD, BE_ROW)}", NUM2),
     ("الأوردرات/يوم فعلي", f"=IFERROR(C7/{at(S_BUD, 6)},0)", NUM2),
     ("المستهدف أوردر/يوم", f"={at(S_BUD, 7)}", NUM)],
    # row 12/13 all-time
    [("رأس المال المدفوع (كله)", f"=SUMIFS({R_AMT},{R_CAT},{CAPITAL_CAT})", NUM),
     ("اتصرف على التأسيس (كله)", f"={q(S_START)}!$L$12", NUM),
     ("الكاش دلوقتي", "=" + CASH_NOW, NUM),
     ("الربح التراكمي", f"={PLQ}!${TOTCOL}${PL['صافي الربح التشغيلي']}", NUM),
     ("شهور الأمان", f"=IFERROR(D13/{at(S_BUD, FIXED_ROW)},0)", NUM2),
     ("عدد الحركات", f"=COUNT({R_AMT})", NUM)],
]
# The card grid sits in columns B..G; labels above values.
for k, row_cards in enumerate(cards):
    lr, vr = 6 + 3 * k, 7 + 3 * k
    for j, (lab, f, fmt) in enumerate(row_cards):
        col = L(2 + j)
        put(ws, f"{col}{lr}", lab, F_KPI_LABEL, CARD_FILL, align=CENTER)
        put(ws, f"{col}{vr}", f, F_KPI, CARD_FILL, fmt, CENTER)
        ws[f"{col}{lr}"].border = Border(top=Side(style="medium", color=GREEN), left=thin, right=thin)
        ws[f"{col}{vr}"].border = Border(bottom=thin, left=thin, right=thin)
    ws.row_dimensions[vr].height = 30
# fix references that assumed fixed positions: المتبقي = budget - actual (row 10: B=actual, C=budget)
ws["D10"].value = "=C10-B10"
ws["F10"].value = f"=IFERROR(C7/{at(S_BUD, 6)},0)"
ws["F13"].value = f"=IFERROR(D13/{at(S_BUD, FIXED_ROW)},0)"
for ref in ("F7", "G7", "D10", "D13", "E13"):
    ws.conditional_formatting.add(ref, CellIsRule(operator="lessThan", formula=["0"],
                                                  font=Font(name=FONT, size=16, bold=True, color="C00000")))
put(ws, "B14", "الصف الأول والتاني للشهر المختار · الصف التالت لكل الفترة من أول ما بدأتوا تسجلوا.", F_NOTE)

# category table
TB = 16
head(ws, TB, 2, ["الفئة", "التصنيف", "الميزانية", "الفعلي", "المتبقي", "% المستخدم", "الحالة"])
for i in range(N_CAT):
    r = TB + 1 + i
    rp, ra = BP0 + i, BA0 + i
    put(ws, f"B{r}", f"={BUD}!$B${rp}", border=True)
    put(ws, f"C{r}", f"={BUD}!$C${rp}", F_NOTE, border=True)
    put(ws, f"D{r}", f"={at(S_BUD, rp)}", fmt=NUM, border=True)
    put(ws, f"E{r}", f"={at(S_BUD, ra)}", fmt=NUM, border=True)
    put(ws, f"F{r}", f"=D{r}-E{r}", fmt=NUM, border=True)
    put(ws, f"G{r}", f"=IFERROR(E{r}/D{r},0)", fmt=PCT, border=True)
    put(ws, f"H{r}",
        f'=IF(B{r}="","",IF(OR(C{r}="إيراد",C{r}="تمويل"),IF(D{r}>0,TEXT(G{r},"0%")&" من المستهدف",""),'
        f'IF(C{r}="سحب","",IF(D{r}=0,IF(E{r}>0,"⚠️ مفيش ميزانية",""),'
        f'IF(E{r}>D{r},"🔴 عدّى الميزانية",IF(E{r}>=D{r}*{C_ALERT},"🟠 قرّب يخلص","🟢 تمام"))))))',
        align=CENTER, border=True)
TB_END = TB + N_CAT
ws.conditional_formatting.add(f"G{TB + 1}:G{TB_END}", DataBarRule(start_type="num", start_value=0, end_type="num",
                                                                  end_value=1, color=MUSTARD))
ws.conditional_formatting.add(f"F{TB + 1}:F{TB_END}", FormulaRule(
    formula=[f'AND(OR($C{TB + 1}="تكلفة مباشرة",$C{TB + 1}="تشغيل",$C{TB + 1}="تأسيس"),F{TB + 1}<0)'],
    font=Font(name=FONT, color="C00000", bold=True)))
# checks
CK = TB_END + 2
put(ws, f"B{CK}", "مراجعة البيانات", F_BOLD, SUB_FILL, border=True)
for c in "CDEFGH":
    put(ws, f"{c}{CK}", None, fill=SUB_FILL, border=True)
checks = [
    ("حركات بفئة مش موجودة في الإعدادات",
     f'=SUMPRODUCT(({R_AMT}<>"")*ISNA(MATCH({R_CAT},{R_CATS},0)))'),
    ("حركات من غير تاريخ", f'=SUMPRODUCT(({R_AMT}<>"")*({R_DATE}=""))'),
    ("حركات قبل أول شهر في الخطة", f'=COUNTIFS({R_DATE},"<"&{C_START})'),
    ("مجموع نسب الشركاء", f"={SET}!$D${PART_LAST + 1}"),
]
for i, (lab, f) in enumerate(checks):
    r = CK + 1 + i
    put(ws, f"B{r}", lab, border=True)
    put(ws, f"C{r}", f, F_BOLD, fmt=PCT if "نسب" in lab else NUM, border=True)
    put(ws, f"D{r}", (f'=IF(C{r}=0,"✓","⚠️ راجع")' if i < 2 else
                      (f'=IF(C{r}=0,"✓","بتتحسب في رصيد أول شهر")' if i == 2 else
                       f'=IF(ROUND(C{r},4)=1,"✓","⚠️ لازم 100%")')), align=CENTER, border=True)
ws.freeze_panes = "A6"

# ================================================================ 11. Guide
ws = sheet(S_HELP)
title(ws, "دليل الاستخدام")
widths(ws, {"A": 2, "B": 30, "C": 90})
rows = [
    ("الفكرة", "كل الفلوس بتتسجل في تاب «الحركات» بس. كل التابات التانية معادلات بتقرا منه ومن الإعدادات."),
    ("الألوان", "أصفر + أزرق = خانة بتكتب فيها · أسود = معادلة متلمسهاش · أخضر = رقم جاي من تاب تاني."),
    ("المبلغ", "دايماً موجب. الفئة هي اللي بتحدد الفلوس داخلة (مبيعات، رأس مال) ولا خارجة (مصاريف، سحب)."),
    ("الشخص", "مين عمل الحركة. في «رأس مال» و«سحب شريك» لازم يبقى اسم الشريك زي ما هو مكتوب في الإعدادات بالظبط."),
    ("طريقة الدفع", "«من جيبه» = الشريك دفع من فلوسه الشخصية → بتتحسب مستحقة له في تاب الشركاء."),
    ("المبيعات", "سجّلها يومي أو بالأسبوع: الفئة «مبيعات»، والكمية = عدد الأوردرات (عشان متوسط الأوردر ونقطة التعادل)."),
    ("الميزانية", "عدّل الأوردرات المتوقعة في اليوم في تاب الميزانية، والقيم في الإعدادات. الباقي بيتحسب لوحده."),
    ("فئة جديدة", "اكتبها في صف فاضي في الإعدادات (اسم + تصنيف + طريقة ميزانية + كلمات للبوت)."),
    ("مثال حركة", "2026-11-22 · خامات وبضاعة · فراخ 10 كيلو من المورد · 10 · 2300 · عمر · كاش الشركة · بوت"),
    ("مثال مبيعات", "2026-11-28 · مبيعات · مبيعات يوم الجمعة · 14 · 3360 · أحمد · كاش الشركة · يدوي"),
    ("مثال رأس مال", "2026-11-20 · رأس مال · دفعة أولى · 1 · 75000 · أسامة · تحويل بنكي · يدوي"),
    ("البوت", "ابعت له زي ما بتتكلم: «صرفت 5000 على البضاعة» · «جبت فراخ ب1200 ورز ب300» · «بعنا ب 3500 في 15 اوردر» · "
              "«أسامة حط 50000 رأس مال» · «دفعت النت 300 من جيبي امبارح». هيوريك ملخص وتأكد قبل ما يكتب."),
    ("أوامر البوت", "/month ملخص الشهر · /budget الميزانية · /today حركات النهارده · /last آخر الحركات · /cash الكاش · "
                    "/undo إلغاء آخر حركة · /export نسخة إكسيل · /cats الفئات · /help"),
    ("Google Sheets", "البوت بيكتب في نسخة Google Sheets من الملف ده. تقدر تنزلها إكسيل في أي وقت (أو /export من البوت)."),
]
for i, (k, v) in enumerate(rows):
    r = 5 + i
    put(ws, f"B{r}", k, F_BOLD, SUB_FILL, align=RIGHT, border=True)
    put(ws, f"C{r}", v, border=True)
    ws.row_dimensions[r].height = 32

# defined names the bot can rely on (stable even if rows move)
wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print("saved", OUT)
print("PL rows", PL)
print("CF rows", CF)
print("BUD", BP0, BA0, BV0, FIXED_ROW, BE_ROW)
