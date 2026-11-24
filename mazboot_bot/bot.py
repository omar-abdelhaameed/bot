"""Telegram bot: text in, confirmed ledger rows out."""
from __future__ import annotations

import asyncio
import datetime as dt
import io
import logging
import re
import secrets
import time
from zoneinfo import ZoneInfo

from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.constants import ParseMode
from telegram.ext import (Application, ApplicationBuilder, CallbackQueryHandler, CommandHandler, ContextTypes,
                          MessageHandler, filters)

from . import layout, texts
from .config import Config
from .ledger import Ledger, month_start
from .models import Entry
from .parser import Parser

log = logging.getLogger(__name__)
PENDING_TTL = 6 * 3600


class MazbootBot:
    def __init__(self, cfg: Config, ledger: Ledger, parser: Parser):
        self.cfg = cfg
        self.ledger = ledger
        self.parser = parser
        self.tz = ZoneInfo(cfg.timezone)
        self.pending: dict[str, dict] = {}

    # ----------------------------------------------------------------- helpers
    def now(self) -> dt.datetime:
        return dt.datetime.now(self.tz)

    def who(self, update: Update) -> str | None:
        u = update.effective_user
        return self.cfg.users.get(u.id) if u else None

    async def guard(self, update: Update) -> str | None:
        name = self.who(update)
        if name is None and update.effective_message:
            uid = update.effective_user.id if update.effective_user else "?"
            await update.effective_message.reply_text(
                f"🔒 البوت ده لشركاء MAZBOOT بس.\nالـ ID بتاعك: <code>{uid}</code>\n"
                f"ابعته لعمر يضيفك في ALLOWED_USERS.", parse_mode=ParseMode.HTML)
        return name

    @staticmethod
    async def run(fn, *a):
        return await asyncio.to_thread(fn, *a)

    def _gc(self):
        cutoff = time.time() - PENDING_TTL
        for k in [k for k, v in self.pending.items() if v["t"] < cutoff]:
            self.pending.pop(k, None)

    def keyboard(self, pid: str, entries: list[Entry]) -> InlineKeyboardMarkup:
        settings = self.ledger.settings()
        missing = next((i for i, e in enumerate(entries) if not e.category), None)
        rows = []
        if missing is not None and entries[missing].amount:
            rows += self.category_buttons(pid, missing, settings)
        elif all(e.complete for e in entries):
            rows.append([InlineKeyboardButton("✅ سجّل", callback_data=f"ok|{pid}"),
                         InlineKeyboardButton("❌ إلغاء", callback_data=f"no|{pid}")])
            edit = [InlineKeyboardButton(f"✏️ فئة {i + 1}" if len(entries) > 1 else "✏️ غيّر الفئة",
                                         callback_data=f"ed|{pid}|{i}") for i in range(len(entries))]
            for k in range(0, len(edit), 4):
                rows.append(edit[k:k + 4])
            rows.append([InlineKeyboardButton("💳 طريقة الدفع", callback_data=f"pm|{pid}")])
            return InlineKeyboardMarkup(rows)
        rows.append([InlineKeyboardButton("❌ إلغاء", callback_data=f"no|{pid}")])
        return InlineKeyboardMarkup(rows)

    @staticmethod
    def category_buttons(pid, i, settings):
        order = {layout.CLS_COGS: 0, layout.CLS_OPEX: 1, layout.CLS_CAPEX: 2, layout.CLS_INCOME: 3,
                 layout.CLS_FUNDING: 4, layout.CLS_WITHDRAW: 5}
        cats = sorted(enumerate(settings.categories), key=lambda kc: (order.get(kc[1].cls, 9), kc[0]))
        btns = [InlineKeyboardButton(f"{texts.ICON.get(c.cls, '•')} {c.name}", callback_data=f"sc|{pid}|{i}|{k}")
                for k, c in cats]
        return [btns[j:j + 2] for j in range(0, len(btns), 2)]

    # ----------------------------------------------------------------- text messages
    async def on_text(self, update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        name = await self.guard(update)
        if not name:
            return
        text = (update.effective_message.text or "").strip()
        if not text:
            return
        self._gc()
        settings = await self.run(self.ledger.settings)
        entries, how = await self.run(self.parser.parse, text, name, self.now().date(), settings)
        if not entries:
            await update.effective_message.reply_text("مش لاقي فلوس في الرسالة دي 🤔 جرّب مثلاً: «صرفت 500 على العلب»")
            return
        if len(entries) > 15:
            entries = entries[:15]
        pid = secrets.token_hex(4)
        self.pending[pid] = {"entries": entries, "uid": update.effective_user.id, "name": name, "t": time.time(),
                             "how": how}
        await update.effective_message.reply_text(
            texts.preview(entries, settings, how), parse_mode=ParseMode.HTML,
            reply_markup=self.keyboard(pid, entries))

    # ----------------------------------------------------------------- buttons
    async def on_button(self, update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        q = update.callback_query
        name = self.who(update)
        if not name:
            await q.answer("🔒 مش مسموح", show_alert=True)
            return
        parts = (q.data or "").split("|")
        action = parts[0]

        if action == "ud":                                   # confirm undo
            await q.answer()
            tx_id = int(parts[1])
            ok = await self.run(self.ledger.delete, tx_id)
            await q.edit_message_text(f"🗑 اتمسحت الحركة #{tx_id}" if ok else "مش لاقي الحركة دي (يمكن اتمسحت).")
            return
        if action == "uc":
            await q.answer()
            await q.edit_message_text("تمام، محدش اتمسح.")
            return

        pid = parts[1] if len(parts) > 1 else ""
        p = self.pending.get(pid)
        if not p:
            await q.answer("الطلب ده قديم — ابعت الرسالة تاني", show_alert=True)
            return
        if p["uid"] != update.effective_user.id:
            await q.answer("ده طلب حد تاني 🙂", show_alert=True)
            return
        settings = self.ledger.settings()
        entries: list[Entry] = p["entries"]

        if action == "no":
            self.pending.pop(pid, None)
            await q.answer()
            await q.edit_message_text("❌ اتلغى، مفيش حاجة اتسجلت.")
        elif action == "ed":
            await q.answer()
            i = int(parts[2])
            await q.edit_message_reply_markup(InlineKeyboardMarkup(
                self.category_buttons(pid, i, settings) + [[InlineKeyboardButton("↩️ رجوع", callback_data=f"bk|{pid}")]]))
        elif action == "bk":
            await q.answer()
            await q.edit_message_reply_markup(self.keyboard(pid, entries))
        elif action == "sc":
            i, k = int(parts[2]), int(parts[3])
            entries[i].category = settings.categories[k].name
            if not entries[i].details:
                entries[i].details = entries[i].category
            await q.answer(f"الفئة: {entries[i].category}")
            await q.edit_message_text(texts.preview(entries, settings, p["how"]), parse_mode=ParseMode.HTML,
                                      reply_markup=self.keyboard(pid, entries))
        elif action == "pm":
            await q.answer()
            btns = [InlineKeyboardButton(m, callback_data=f"sm|{pid}|{j}") for j, m in enumerate(layout.METHODS)]
            await q.edit_message_reply_markup(InlineKeyboardMarkup([btns[j:j + 2] for j in range(0, len(btns), 2)]))
        elif action == "sm":
            method = layout.METHODS[int(parts[2])]
            for e in entries:
                e.method = method
            await q.answer(method)
            await q.edit_message_text(texts.preview(entries, settings, p["how"]), parse_mode=ParseMode.HTML,
                                      reply_markup=self.keyboard(pid, entries))
        elif action == "ok":
            if not all(e.complete for e in entries):
                await q.answer("لسه فيه بند ناقص", show_alert=True)
                return
            self.pending.pop(pid, None)
            await q.answer("بسجّل...")
            try:
                ids = await self.run(self.ledger.add, entries, p["name"], self.now())
            except PermissionError:
                self.pending[pid] = p
                await q.edit_message_text("⚠️ الملف مفتوح في إكسيل ومش قادر أكتب فيه. اقفله ودوس سجّل تاني.",
                                          reply_markup=self.keyboard(pid, entries))
                return
            except Exception as e:
                log.exception("save failed")
                self.pending[pid] = p
                await q.edit_message_text(f"⚠️ حصلت مشكلة وأنا بسجّل: {e}\nجرّب تاني.",
                                          reply_markup=self.keyboard(pid, entries))
                return
            msg = texts.saved(entries, ids, settings)
            try:
                months = {month_start(e.date or self.now().date()) for e in entries}
                warn = []
                for m in sorted(months):
                    cats = {e.category for e in entries if month_start(e.date or self.now().date()) == m}
                    warn += await self.run(self.ledger.alerts_for, cats, m)
                if warn:
                    msg += "\n\n" + texts.alerts(warn)
            except Exception:
                log.exception("budget alert check failed")
            await q.edit_message_text(msg, parse_mode=ParseMode.HTML)
            if self.cfg.notify_chat and self.cfg.notify_chat != q.message.chat_id:
                try:
                    await ctx.bot.send_message(self.cfg.notify_chat, f"👤 {p['name']}\n" + msg, parse_mode=ParseMode.HTML)
                except Exception:
                    log.exception("notify failed")
        else:
            await q.answer()

    # ----------------------------------------------------------------- commands
    def _month_arg(self, ctx) -> dt.date:
        if ctx.args:
            m = re.match(r"(\d{4})[-/](\d{1,2})", ctx.args[0])
            if m:
                return dt.date(int(m.group(1)), int(m.group(2)), 1)
            m = re.match(r"(\d{1,2})[-/](\d{4})", ctx.args[0])
            if m:
                return dt.date(int(m.group(2)), int(m.group(1)), 1)
        return month_start(self.now().date())

    async def cmd_start(self, update: Update, ctx):
        if await self.guard(update):
            await update.effective_message.reply_text(texts.HELP, parse_mode=ParseMode.HTML)

    async def cmd_id(self, update: Update, ctx):
        await update.effective_message.reply_text(f"الـ ID بتاعك: <code>{update.effective_user.id}</code>",
                                                  parse_mode=ParseMode.HTML)

    async def cmd_month(self, update: Update, ctx):
        if not await self.guard(update):
            return
        m = self._month_arg(ctx)

        def work():
            txs = self.ledger.transactions()
            return self.ledger.month_summary(m), self.ledger.cash_now(txs), self.ledger.budget(m, txs)
        s, cash, bud = await self.run(work)
        await update.effective_message.reply_text(texts.month_summary(s, cash, bud), parse_mode=ParseMode.HTML)

    async def cmd_budget(self, update: Update, ctx):
        if not await self.guard(update):
            return
        m = self._month_arg(ctx)
        lines = await self.run(self.ledger.budget, m)
        await update.effective_message.reply_text(
            texts.budget_table(m, lines, self.ledger.settings().alert_pct), parse_mode=ParseMode.HTML)

    async def cmd_today(self, update: Update, ctx):
        if not await self.guard(update):
            return
        today = self.now().date()
        txs = [t for t in await self.run(self.ledger.transactions) if t.date == today]
        if not txs:
            await update.effective_message.reply_text("مفيش حركات النهارده لسه.")
            return
        total_out = sum(t.amount for t in txs if self.ledger.cls_of(t.category) in layout.MONEY_OUT)
        total_in = sum(t.amount for t in txs if self.ledger.cls_of(t.category) in layout.MONEY_IN)
        body = "\n".join(texts.tx_line(t) for t in txs)
        await update.effective_message.reply_text(
            f"📅 <b>النهارده</b>\n{body}\n\nخارج: <b>{texts.money(total_out)}</b> · داخل: <b>{texts.money(total_in)}</b>",
            parse_mode=ParseMode.HTML)

    async def cmd_last(self, update: Update, ctx):
        if not await self.guard(update):
            return
        n = 10
        if ctx.args and ctx.args[0].isdigit():
            n = max(1, min(30, int(ctx.args[0])))
        txs = sorted(await self.run(self.ledger.transactions), key=lambda t: (t.id or 0, t.row))[-n:]
        if not txs:
            await update.effective_message.reply_text("الدفتر فاضي لسه.")
            return
        await update.effective_message.reply_text(f"🧾 <b>آخر {len(txs)} حركات</b>\n" +
                                                  "\n".join(texts.tx_line(t) for t in reversed(txs)),
                                                  parse_mode=ParseMode.HTML)

    async def cmd_cash(self, update: Update, ctx):
        if not await self.guard(update):
            return

        def work():
            txs = self.ledger.transactions()
            return self.ledger.cash_now(txs), self.ledger.month_summary(month_start(self.now().date()))
        cash, s = await self.run(work)
        await update.effective_message.reply_text(
            f"💵 الكاش دلوقتي: <b>{texts.money(cash)}</b>\n"
            f"صافي الشهر ده (تشغيلي): {texts.money(s['net'])}\n"
            f"<i>= الكاش الافتتاحي + كل اللي دخل − كل اللي خرج. المدفوع «من جيبه» مش بيقلل الكاش.</i>",
            parse_mode=ParseMode.HTML)

    async def cmd_undo(self, update: Update, ctx):
        name = await self.guard(update)
        if not name:
            return
        t = await self.run(self.ledger.last_bot_tx, name)
        if not t:
            await update.effective_message.reply_text("مفيش حركات متسجلة من البوت.")
            return
        await update.effective_message.reply_text(
            "🗑 أمسح الحركة دي؟\n" + texts.tx_line(t), parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("🗑 امسح", callback_data=f"ud|{t.id}"),
                                                InlineKeyboardButton("لا", callback_data="uc")]]))

    async def cmd_export(self, update: Update, ctx):
        if not await self.guard(update):
            return
        data = await self.run(self.ledger.storage.export_xlsx)
        fname = f"MAZBOOT_Finance_{self.now():%Y-%m-%d}.xlsx"
        await update.effective_message.reply_document(io.BytesIO(data), filename=fname,
                                                      caption="📎 نسخة إكسيل من الشيت دلوقتي")

    async def cmd_cats(self, update: Update, ctx):
        if await self.guard(update):
            await update.effective_message.reply_text(texts.categories(self.ledger.settings()), parse_mode=ParseMode.HTML)

    async def cmd_reload(self, update: Update, ctx):
        if await self.guard(update):
            s = await self.run(self.ledger.settings, True)
            await update.effective_message.reply_text(f"🔄 اتحدّث: {len(s.categories)} فئة، {len(s.partners)} شركاء.")

    async def on_error(self, update: object, ctx: ContextTypes.DEFAULT_TYPE):
        log.exception("handler error", exc_info=ctx.error)
        if isinstance(update, Update) and update.effective_message:
            try:
                await update.effective_message.reply_text(f"⚠️ حصلت مشكلة: {ctx.error}")
            except Exception:
                pass

    # ----------------------------------------------------------------- wiring
    def build(self) -> Application:
        # Default 5s timeouts are too tight on slow links (getMe took ~8s here) -> TimedOut at startup.
        app = (ApplicationBuilder().token(self.cfg.telegram_token)
               .connect_timeout(20).read_timeout(30).write_timeout(30).pool_timeout(20)
               .get_updates_connect_timeout(20).get_updates_read_timeout(40)
               .build())
        for cmd, fn in [("start", self.cmd_start), ("help", self.cmd_start), ("id", self.cmd_id),
                        ("month", self.cmd_month), ("budget", self.cmd_budget), ("today", self.cmd_today),
                        ("last", self.cmd_last), ("cash", self.cmd_cash), ("undo", self.cmd_undo),
                        ("export", self.cmd_export), ("cats", self.cmd_cats), ("reload", self.cmd_reload)]:
            app.add_handler(CommandHandler(cmd, fn))
        app.add_handler(CallbackQueryHandler(self.on_button))
        app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, self.on_text))
        app.add_error_handler(self.on_error)
        return app
