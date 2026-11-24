from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field


@dataclass
class Category:
    name: str
    cls: str
    keywords: list[str] = field(default_factory=list)


@dataclass
class Settings:
    start_month: dt.date
    opening_cash: float
    alert_pct: float
    partners: list[str]
    categories: list[Category]

    def category(self, name: str | None) -> Category | None:
        if not name:
            return None
        for c in self.categories:
            if c.name == name:
                return c
        return None


@dataclass
class Tx:
    """One row of the ledger tab."""
    row: int                    # sheet row number
    id: int | None
    date: dt.date | None
    category: str
    details: str
    qty: float | None
    amount: float
    person: str
    method: str
    source: str
    logged_at: dt.datetime | None
    notes: str


@dataclass
class Entry:
    """A parsed, not-yet-saved transaction."""
    amount: float | None
    category: str | None
    details: str = ""
    qty: float | None = None
    date: dt.date | None = None
    method: str | None = None
    person: str | None = None
    notes: str = ""

    @property
    def complete(self) -> bool:
        return bool(self.amount and self.amount > 0 and self.category)
