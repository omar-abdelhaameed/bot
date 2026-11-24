"""Business logic between the Telegram handlers and the storage backend."""
from __future__ import annotations

import datetime as dt
import threading
import time
from collections import defaultdict
from dataclasses import dataclass

from . import layout
from .models import Entry, Settings, Tx
from .storage import Storage


def month_start(d: dt.date) -> dt.date:
    return d.replace(day=1)


@dataclass
class BudgetLine:
    category: str
    cls: str
    planned: float
    actual: float

    @property
    def remaining(self) -> float:
        return self.planned - self.actual

    @property
    def used(self) -> float | None:
        return self.actual / self.planned if self.planned else None


class Ledger:
    SETTINGS_TTL = 300

    def __init__(self, storage: Storage):
        self.storage = storage
        self._settings: Settings | None = None
        self._settings_at = 0.0
        self._lock = threading.Lock()

    # ----------------------------------------------------------------- reads
    def settings(self, force=False) -> Settings:
        if force or self._settings is None or time.time() - self._settings_at > self.SETTINGS_TTL:
            self._settings = self.storage.read_settings()
            self._settings_at = time.time()
        return self._settings

    def transactions(self) -> list[Tx]:
        return self.storage.read_transactions()

    def cls_of(self, category: str) -> str:
        c = self.settings().category(category)
        return c.cls if c else ""

    # ----------------------------------------------------------------- writes
    def add(self, entries: list[Entry], logged_by: str, now: dt.datetime) -> list[int]:
        with self._lock:
            txs = self.transactions()
            next_id = max([t.id for t in txs if t.id is not None] + [0]) + 1
            rows, ids = [], []
            for e in entries:
                if not e.complete:
                    raise ValueError("entry is incomplete")
                ids.append(next_id)
                rows.append([next_id, e.date or now.date(), e.category, e.details or e.category,
                             e.qty if e.qty is not None else "", round(float(e.amount), 2),
                             e.person or logged_by, e.method or layout.DEFAULT_METHOD,
                             layout.SOURCE_BOT, now.replace(tzinfo=None, microsecond=0), e.notes or ""])
                next_id += 1
            self.storage.append_rows(rows)
            return ids

    def last_bot_tx(self, person: str) -> Tx | None:
        mine = [t for t in self.transactions() if t.source == layout.SOURCE_BOT and t.person == person and t.id]
        if not mine:
            mine = [t for t in self.transactions() if t.source == layout.SOURCE_BOT and t.id]
        return max(mine, key=lambda t: t.id) if mine else None

    def delete(self, tx_id: int) -> bool:
        with self._lock:
            return self.storage.delete_tx(tx_id)

    # ----------------------------------------------------------------- numbers
    def actual_by_category(self, month: dt.date, txs: list[Tx] | None = None) -> dict[str, float]:
        txs = self.transactions() if txs is None else txs
        out: dict[str, float] = defaultdict(float)
        for t in txs:
            if t.date and month_start(t.date) == month:
                out[t.category] += t.amount
        return out

    def budget(self, month: dt.date, txs: list[Tx] | None = None) -> list[BudgetLine]:
        planned = self.storage.read_planned().get(month, {})
        actual = self.actual_by_category(month, txs)
        lines = []
        for c in self.settings().categories:
            lines.append(BudgetLine(c.name, c.cls, planned.get(c.name, 0.0), actual.get(c.name, 0.0)))
        return lines

    def month_summary(self, month: dt.date) -> dict:
        txs = self.transactions()
        act = self.actual_by_category(month, txs)
        s = self.settings()
        by_cls: dict[str, float] = defaultdict(float)
        for name, v in act.items():
            by_cls[self.cls_of(name) or "؟"] += v
        sales = next((c.name for c in s.categories if c.cls == layout.CLS_INCOME), "مبيعات")
        food_cat = next((c.name for c in s.categories if c.cls == layout.CLS_COGS), "خامات وبضاعة")
        orders = sum((t.qty or 0) for t in txs if t.date and month_start(t.date) == month and t.category == sales)
        income = by_cls[layout.CLS_INCOME]
        cogs = by_cls[layout.CLS_COGS]
        opex = by_cls[layout.CLS_OPEX]
        food = act.get(food_cat, 0.0)
        sales_amt = act.get(sales, 0.0)
        return {
            "month": month, "income": income, "orders": orders, "cogs": cogs, "opex": opex,
            "capex": by_cls[layout.CLS_CAPEX], "funding": by_cls[layout.CLS_FUNDING],
            "withdraw": by_cls[layout.CLS_WITHDRAW], "gross": income - cogs, "net": income - cogs - opex,
            "food_pct": food / sales_amt if sales_amt else None, "avg_order": sales_amt / orders if orders else None,
            "count": sum(1 for t in txs if t.date and month_start(t.date) == month),
            "unknown": sorted({k for k in act if not self.cls_of(k)}),
        }

    def cash_now(self, txs: list[Tx] | None = None) -> float:
        txs = self.transactions() if txs is None else txs
        s = self.settings()
        cash = s.opening_cash
        for t in txs:
            cls = self.cls_of(t.category)
            if cls in layout.MONEY_IN:
                cash += t.amount
            elif cls in layout.MONEY_OUT:
                cash -= t.amount
            if t.method == layout.OUT_OF_POCKET:
                cash += t.amount       # a partner paid personally: company cash didn't move
        return cash

    def alerts_for(self, categories: set[str], month: dt.date) -> list[tuple[str, str, BudgetLine]]:
        """(level, category, line) for expense categories at/over the alert threshold."""
        s = self.settings()
        out = []
        for line in self.budget(month):
            if line.category not in categories or line.cls not in layout.EXPENSE_CLASSES:
                continue
            if line.planned <= 0:
                if line.actual > 0:
                    out.append(("none", line.category, line))
            elif line.actual > line.planned:
                out.append(("over", line.category, line))
            elif line.actual >= line.planned * s.alert_pct:
                out.append(("near", line.category, line))
        return out
