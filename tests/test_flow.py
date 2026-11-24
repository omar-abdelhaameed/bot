"""End-to-end tests: parser cases + the real handlers against a copy of the workbook.

Run:  python -m pytest -q tests   (from the bot folder; needs ../out/MAZBOOT_Finance.xlsx or WORKBOOK env)
"""
import asyncio
import datetime as dt
import os
import shutil
import sys
from types import SimpleNamespace

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from mazboot_bot.bot import MazbootBot            # noqa: E402
from mazboot_bot.config import Config             # noqa: E402
from mazboot_bot.ledger import Ledger             # noqa: E402
from mazboot_bot.parser import Parser             # noqa: E402
from mazboot_bot.parser_rules import RuleParser   # noqa: E402
from mazboot_bot.storage import ExcelStorage      # noqa: E402

_HERE = os.path.dirname(os.path.abspath(__file__))
WB = os.environ.get("WORKBOOK") or next(
    (p for p in (os.path.join(_HERE, "..", "MAZBOOT_Finance.xlsx"), os.path.join(_HERE, "..", "..", "out", "MAZBOOT_Finance.xlsx"))
     if os.path.exists(p)), "MAZBOOT_Finance.xlsx")
TODAY = dt.date(2026, 11, 22)


@pytest.fixture()
def wb(tmp_path):
    p = tmp_path / "MAZBOOT_Finance.xlsx"
    shutil.copy(WB, p)
    return str(p)


@pytest.fixture()
def rules(wb):
    return RuleParser(ExcelStorage(wb).read_settings())


CASES = [
    ("صرفت 5000 على البضاعة", [(5000, "خامات وبضاعة")]),
    ("جبت فراخ ب1200 ورز ب300", [(1200, "خامات وبضاعة"), (300, "خامات وبضاعة")]),
    ("صرفت 5000 على البضاعة و 300 نت", [(5000, "خامات وبضاعة"), (300, "إنترنت وموبايل")]),
    ("2 كيلو فراخ ب 460", [(460, "خامات وبضاعة")]),
    ("بعنا ب 3500 في 15 اوردر", [(3500, "مبيعات")]),
    ("أسامة حط 50000 رأس مال", [(50000, "رأس مال")]),
    ("دفعت النت 300 من جيبي امبارح", [(300, "إنترنت وموبايل")]),
    ("اشتريت ثلاجة ب 25 الف", [(25000, "معدات وأدوات")]),
    ("علب تغليف 1500 جنيه فودافون كاش", [(1500, "تغليف")]),
    ("دفعت الإيجار ٤٥٠٠", [(4500, "إيجار")]),
    ("الف ونص كهربا", [(1500, "كهرباء ومياه وغاز")]),
    ("3 الاف و 500 بضاعة", [(3500, "خامات وبضاعة")]),
    ("اعلان ممول 700", [(700, "إعلانات ممولة")]),
    ("مرتب المساعد 5000", [(5000, "مرتبات")]),
    ("سحبت 2000", [(2000, "سحب شريك")]),
    ("2.5k تغليف", [(2500, "تغليف")]),
    ("حاجة غريبة 250", [(250, None)]),
]


@pytest.mark.parametrize("text,expected", CASES)
def test_rule_parser(rules, text, expected):
    got = [(e.amount, e.category) for e in rules.parse(text, "عمر", TODAY)]
    assert got == [(float(a), c) for a, c in expected]


def test_parser_details(rules):
    e = rules.parse("دفعت النت 300 من جيبي امبارح", "عمر", TODAY)[0]
    assert e.method == "من جيبه" and e.date == TODAY - dt.timedelta(days=1) and e.person == "عمر"
    e = rules.parse("أسامة حط 50000 رأس مال", "عمر", TODAY)[0]
    assert e.person == "أسامة"
    e = rules.parse("بعنا ب 3500 في 15 اوردر", "أحمد", TODAY)[0]
    assert e.qty == 15 and e.person == "أحمد"
    e = rules.parse("15/11 بنزين توصيل 120", "عمر", TODAY)[0]
    assert e.date == dt.date(2026, 11, 15) and e.category == "توصيل" and e.amount == 120


# ---------------------------------------------------------------- handler simulation
class Msg:
    def __init__(self, text="", chat_id=1):
        self.text, self.chat_id, self.out = text, chat_id, []

    async def reply_text(self, text, **kw):
        self.out.append((text, kw.get("reply_markup")))

    async def reply_document(self, f, filename=None, caption=None):
        self.out.append(("DOC:" + filename, f.read()))


class Query:
    def __init__(self, data, message):
        self.data, self.message, self.edits, self.answers = data, message, [], []

    async def answer(self, *a, **k):
        self.answers.append(a)

    async def edit_message_text(self, text, **kw):
        self.edits.append((text, kw.get("reply_markup")))

    async def edit_message_reply_markup(self, markup):
        self.edits.append(("MARKUP", markup))


def upd(uid, msg=None, query=None):
    return SimpleNamespace(effective_user=SimpleNamespace(id=uid), effective_message=msg, callback_query=query)


def buttons(markup):
    return [b.callback_data for row in markup.inline_keyboard for b in row]


@pytest.fixture()
def bot(wb):
    cfg = Config(telegram_token="x", users={1: "عمر", 2: "أحمد"}, backend="excel", sheet_id="",
                 google_credentials="", excel_path=wb, anthropic_key="", model="", timezone="Africa/Cairo",
                 notify_chat=None)
    b = MazbootBot(cfg, Ledger(ExcelStorage(wb)), Parser(None))
    b.now = lambda: dt.datetime(2026, 11, 22, 14, 30, tzinfo=b.tz)
    return b


def run(c):
    return asyncio.run(c)


def send(bot, uid, text):
    m = Msg(text)
    run(bot.on_text(upd(uid, m), None))
    return m.out[-1]


def press(bot, uid, data):
    q = Query(data, Msg())
    run(bot.on_button(upd(uid, None, q), SimpleNamespace(bot=None)))
    return q


def test_full_flow(bot, wb):
    # unauthorised user gets their id
    m = Msg("صرفت 100")
    run(bot.on_text(upd(99, m), None))
    assert "99" in m.out[0][0]

    # 1) simple expense -> confirm
    text, kb = send(bot, 1, "صرفت 5000 على البضاعة")
    assert "خامات وبضاعة" in text and "5,000" in text
    ok = [d for d in buttons(kb) if d.startswith("ok|")][0]
    q = press(bot, 1, ok)
    assert "اتسجل" in q.edits[-1][0] and "#1" in q.edits[-1][0]

    # 2) someone else cannot confirm my pending entry
    text, kb = send(bot, 1, "جبت فراخ ب1200 ورز ب300")
    ok = [d for d in buttons(kb) if d.startswith("ok|")][0]
    q = press(bot, 2, ok)
    assert q.answers and "حد تاني" in q.answers[-1][0]
    q = press(bot, 1, ok)
    assert "#2" in q.edits[-1][0] and "#3" in q.edits[-1][0]

    # 3) unknown category -> category buttons -> choose -> confirm
    text, kb = send(bot, 2, "حاجة غريبة 250")
    sc = [d for d in buttons(kb) if d.startswith("sc|")]
    settings = bot.ledger.settings()
    k = [i for i, c in enumerate(settings.categories) if c.name == "نثريات"][0]
    pick = [d for d in sc if d.endswith(f"|{k}")][0]
    q = press(bot, 2, pick)
    assert "نثريات" in q.edits[-1][0]
    ok = [d for d in buttons(q.edits[-1][1]) if d.startswith("ok|")][0]
    press(bot, 2, ok)

    # 4) change payment method before saving
    text, kb = send(bot, 1, "دفعت النت 300 امبارح")
    pm = [d for d in buttons(kb) if d.startswith("pm|")][0]
    q = press(bot, 1, pm)
    sm = [d for d in buttons(q.edits[-1][1]) if d.endswith("|1")][0]     # من جيبه
    q = press(bot, 1, sm)
    assert "من جيبه" in q.edits[-1][0]
    press(bot, 1, [d for d in buttons(q.edits[-1][1]) if d.startswith("ok|")][0])

    # 5) capital + sales + cancel
    for t in ("أسامة حط 75000 رأس مال", "بعنا ب 3360 في 14 اوردر"):
        _, kb = send(bot, 2, t)
        press(bot, 2, [d for d in buttons(kb) if d.startswith("ok|")][0])
    _, kb = send(bot, 1, "صرفت 999 على العلب")
    q = press(bot, 1, [d for d in buttons(kb) if d.startswith("no|")][0])
    assert "اتلغى" in q.edits[-1][0]

    txs = bot.ledger.transactions()
    assert [(t.id, t.category, t.amount, t.person, t.method) for t in txs] == [
        (1, "خامات وبضاعة", 5000, "عمر", "كاش الشركة"),
        (2, "خامات وبضاعة", 1200, "عمر", "كاش الشركة"),
        (3, "خامات وبضاعة", 300, "عمر", "كاش الشركة"),
        (4, "نثريات", 250, "أحمد", "كاش الشركة"),
        (5, "إنترنت وموبايل", 300, "عمر", "من جيبه"),
        (6, "رأس مال", 75000, "أسامة", "كاش الشركة"),
        (7, "مبيعات", 3360, "أحمد", "كاش الشركة"),
    ]
    assert txs[4].date == dt.date(2026, 11, 21) and txs[6].qty == 14 and txs[0].source == "بوت"

    # 6) budget alert: internet budget is 500/month; 300 spent -> 250 more crosses 80%
    _, kb = send(bot, 1, "رصيد موبايل 150")
    q = press(bot, 1, [d for d in buttons(kb) if d.startswith("ok|")][0])
    assert "قرّبت تخلص" in q.edits[-1][0]
    _, kb = send(bot, 1, "باقة نت 200")
    q = press(bot, 1, [d for d in buttons(kb) if d.startswith("ok|")][0])
    assert "عدّت ميزانية" in q.edits[-1][0]

    # 7) summaries
    m = Msg(); run(bot.cmd_month(upd(1, m), SimpleNamespace(args=[])))
    assert "3,360" in m.out[-1][0] and "14 أوردر" in m.out[-1][0]
    cash_expected = 75000 + 3360 - (5000 + 1200 + 300 + 250 + 300 + 150 + 200) + 300   # من جيبه adds back
    m = Msg(); run(bot.cmd_cash(upd(1, m), SimpleNamespace(args=[])))
    assert f"{cash_expected:,}" in m.out[-1][0]
    m = Msg(); run(bot.cmd_budget(upd(1, m), SimpleNamespace(args=[])))
    assert "إنترنت وموبايل" in m.out[-1][0] and "🔴" in m.out[-1][0]
    m = Msg(); run(bot.cmd_today(upd(1, m), SimpleNamespace(args=[])))
    assert "#1" in m.out[-1][0]
    m = Msg(); run(bot.cmd_last(upd(1, m), SimpleNamespace(args=["3"])))
    assert m.out[-1][0].count("#") == 3

    # 8) undo my last bot entry (#9)
    m = Msg(); run(bot.cmd_undo(upd(1, m), SimpleNamespace(args=[])))
    ud = [d for d in buttons(m.out[-1][1]) if d.startswith("ud|")][0]
    assert ud == "ud|9"
    q = press(bot, 1, ud)
    assert "#9" in q.edits[-1][0]
    assert [t.id for t in bot.ledger.transactions()][-1] == 8

    # 9) export
    m = Msg(); run(bot.cmd_export(upd(1, m), SimpleNamespace(args=[])))
    assert m.out[-1][0].startswith("DOC:") and m.out[-1][1][:2] == b"PK"
