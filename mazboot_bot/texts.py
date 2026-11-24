"""Message formatting (Arabic, Telegram HTML)."""
from __future__ import annotations

import datetime as dt
from html import escape

from . import layout
from .ledger import BudgetLine
from .models import Entry, Settings, Tx

AR_MONTHS = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر",
             "نوفمبر", "ديسمبر"]
ICON = {layout.CLS_INCOME: "💰", layout.CLS_FUNDING: "🏦", layout.CLS_WITHDRAW: "↩️",
        layout.CLS_COGS: "🥩", layout.CLS_OPEX: "💸", layout.CLS_CAPEX: "🧰"}


def money(x: float | None) -> str:
    if x is None:
        return "—"
    s = f"{abs(x):,.0f}" if abs(x - round(x)) < 0.005 else f"{abs(x):,.2f}"
    return f"{'-' if x < 0 else ''}{s} ج"


def num(x: float | None) -> str:
    if x is None:
        return ""
    return f"{x:,.0f}" if float(x).is_integer() else f"{x:,.1f}"


def ddmm(d: dt.date | None) -> str:
    return d.strftime("%d/%m") if d else "؟"


def month_name(m: dt.date) -> str:
    return f"{AR_MONTHS[m.month - 1]} {m.year}"


def pct(x: float | None) -> str:
    return "—" if x is None else f"{x * 100:.0f}%"


def entry_line(i: int, e: Entry, settings: Settings) -> str:
    c = settings.category(e.category)
    icon = ICON.get(c.cls, "•") if c else "❓"
    cat = escape(e.category) if e.category else "<i>مش عارف الفئة</i>"
    amt = f"<b>{money(e.amount)}</b>" if e.amount else "<i>مفيش مبلغ</i>"
    meta = [escape(e.details)] if e.details and e.details != e.category else []
    if e.qty is not None:
        meta.append(f"كمية {num(e.qty)}")
    meta += [ddmm(e.date), escape(e.person or ""), escape(e.method or layout.DEFAULT_METHOD)]
    return f"{i}) {icon} {cat} — {amt}\n    <i>{' · '.join(m for m in meta if m)}</i>"


def preview(entries: list[Entry], settings: Settings, how: str) -> str:
    lines = ["📝 <b>هسجّل كده:</b>"]
    lines += [entry_line(i + 1, e, settings) for i, e in enumerate(entries)]
    total = sum(e.amount or 0 for e in entries)
    if len(entries) > 1:
        lines.append(f"\nالإجمالي: <b>{money(total)}</b>")
    if any(not e.amount for e in entries):
        lines.append("\n⚠️ فيه بند من غير مبلغ — ابعت الرسالة تاني بالمبلغ، أو ألغي.")
    elif any(not e.category for e in entries):
        lines.append("\n👇 اختار الفئة من الزراير.")
    else:
        lines.append("\nتمام كده؟")
    if how == "rules":
        lines.append("<i>(فهمتها بالقواعد)</i>")
    return "\n".join(lines)


def saved(entries: list[Entry], ids: list[int], settings: Settings) -> str:
    lines = [f"✅ <b>اتسجل</b> ({'، '.join('#' + str(i) for i in ids)})"]
    lines += [entry_line(i + 1, e, settings) for i, e in enumerate(entries)]
    return "\n".join(lines)


def alerts(items: list[tuple[str, str, BudgetLine]]) -> str:
    out = []
    for level, cat, line in items:
        if level == "over":
            out.append(f"🔴 <b>{escape(cat)}</b> عدّت ميزانية الشهر: {money(line.actual)} من {money(line.planned)} "
                       f"(زيادة {money(-line.remaining)})")
        elif level == "near":
            out.append(f"🟠 <b>{escape(cat)}</b> قرّبت تخلص: {money(line.actual)} من {money(line.planned)} "
                       f"({pct(line.used)}) — باقي {money(line.remaining)}")
        else:
            out.append(f"⚠️ <b>{escape(cat)}</b> مالهاش ميزانية الشهر ده واتصرف عليها {money(line.actual)}")
    return "\n".join(out)


def tx_line(t: Tx) -> str:
    return (f"#{t.id or '?'} · {ddmm(t.date)} · {escape(t.category)} — <b>{money(t.amount)}</b>"
            f"\n    <i>{escape(t.details)} · {escape(t.person)}</i>")


def month_summary(s: dict, cash: float, budget: list[BudgetLine]) -> str:
    planned_exp = sum(b.planned for b in budget if b.cls in layout.EXPENSE_CLASSES)
    actual_exp = sum(b.actual for b in budget if b.cls in layout.EXPENSE_CLASSES)
    lines = [f"📊 <b>{month_name(s['month'])}</b>",
             f"💰 المبيعات: <b>{money(s['income'])}</b>" + (f" · {num(s['orders'])} أوردر" if s['orders'] else ""),
             f"   متوسط الأوردر: {money(s['avg_order'])}" if s["avg_order"] else None,
             f"🥩 تكلفة مباشرة: {money(s['cogs'])}" + (f" ({pct(s['food_pct'])} أكل)" if s["food_pct"] is not None else ""),
             f"💸 مصاريف تشغيل: {money(s['opex'])}",
             f"{'🟢' if s['net'] >= 0 else '🔴'} صافي الربح التشغيلي: <b>{money(s['net'])}</b>",
             f"🧰 تأسيس ومعدات: {money(s['capex'])}" if s["capex"] else None,
             f"🏦 رأس مال دخل: {money(s['funding'])}" if s["funding"] else None,
             f"↩️ سحوبات: {money(s['withdraw'])}" if s["withdraw"] else None,
             "",
             f"📒 المصروف قصاد الميزانية: {money(actual_exp)} من {money(planned_exp)}"
             + (f" ({pct(actual_exp / planned_exp)})" if planned_exp else ""),
             f"💵 الكاش دلوقتي: <b>{money(cash)}</b>",
             f"<i>{s['count']} حركة الشهر ده</i>"]
    if s["unknown"]:
        lines.append("⚠️ فئات مش موجودة في الإعدادات: " + "، ".join(escape(u) for u in s["unknown"]))
    return "\n".join(l for l in lines if l is not None)


def budget_table(month: dt.date, lines: list[BudgetLine], alert_pct: float) -> str:
    out = [f"📒 <b>ميزانية {month_name(month)}</b>"]
    exp = [b for b in lines if b.cls in layout.EXPENSE_CLASSES and (b.planned or b.actual)]
    for b in exp:
        if b.planned <= 0:
            mark = "⚠️"
        elif b.actual > b.planned:
            mark = "🔴"
        elif b.actual >= b.planned * alert_pct:
            mark = "🟠"
        else:
            mark = "🟢"
        rest = f"باقي {money(b.remaining)}" if b.remaining >= 0 else f"زيادة {money(-b.remaining)}"
        out.append(f"{mark} {escape(b.category)}: {money(b.actual)} / {money(b.planned)} ({pct(b.used)}) — {rest}")
    tp = sum(b.planned for b in exp)
    ta = sum(b.actual for b in exp)
    out.append(f"\nالإجمالي: <b>{money(ta)}</b> من {money(tp)} · باقي {money(tp - ta)}")
    sales = next((b for b in lines if b.cls == layout.CLS_INCOME), None)
    if sales and sales.planned:
        out.append(f"💰 المبيعات: {money(sales.actual)} من المستهدف {money(sales.planned)} ({pct(sales.used)})")
    if not tp:
        out.append("<i>مفيش أرقام ميزانية للشهر ده — راجع تاب الميزانية أو أول شهر في الإعدادات.</i>")
    return "\n".join(out)


def categories(settings: Settings) -> str:
    groups: dict[str, list[str]] = {}
    for c in settings.categories:
        groups.setdefault(c.cls or "؟", []).append(c.name)
    out = ["🗂 <b>الفئات</b> (بتتعدل من تاب الإعدادات)"]
    for cls, names in groups.items():
        out.append(f"{ICON.get(cls, '•')} <b>{escape(cls)}</b>: " + "، ".join(escape(n) for n in names))
    return "\n".join(out)


HELP = """👋 <b>بوت حسابات MAZBOOT</b>

ابعتلي أي حاجة صرفتها أو دخلت زي ما بتتكلم:
• <code>صرفت 5000 على البضاعة</code>
• <code>جبت فراخ ب1200 ورز ب300</code>
• <code>2 كيلو فراخ ب 460</code>
• <code>بعنا ب 3500 في 15 اوردر</code>
• <code>أسامة حط 50000 رأس مال</code>
• <code>دفعت النت 300 من جيبي امبارح</code>
• <code>15/11 بنزين توصيل 120 فودافون كاش</code>

هوريك ملخص وتدوس ✅ قبل ما يتكتب في الشيت.

<b>الأوامر</b>
/month — ملخص الشهر (أو <code>/month 2026-12</code>)
/budget — الميزانية قصاد المصروف
/today — حركات النهارده
/last — آخر 10 حركات
/cash — الكاش دلوقتي
/undo — امسح آخر حركة سجلتها
/export — نسخة إكسيل من الملف
/cats — الفئات
/reload — حدّث الفئات بعد ما تعدل الإعدادات"""
