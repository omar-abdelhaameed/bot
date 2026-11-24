"""Optional AI parser (Claude API). Understands messy Egyptian Arabic and multi-item messages.

Used only when ANTHROPIC_API_KEY is set. Any failure falls back to the rule parser.
"""
from __future__ import annotations

import datetime as dt
import json
import logging

from . import layout
from .models import Entry, Settings

log = logging.getLogger(__name__)

SYSTEM = """You turn short Egyptian-Arabic (or English) messages from the owners of MAZBOOT, a healthy-food \
cloud kitchen in Shibin El Kom, into bookkeeping entries for their ledger. Amounts are Egyptian pounds.

Rules:
- One entry per separate amount of money. "جبت فراخ ب1200 ورز ب300" = two entries.
- amount is always positive. "5 آلاف" = 5000, "الف ونص" = 1500, "2.5k" = 2500.
- Numbers followed by a unit (كيلو، علبة، قطعة...) are quantities, not amounts. For sales, qty = number of orders.
- Pick category ONLY from the list given. If nothing fits, use null — never invent a category.
- person: the partner the money belongs to. Default is the sender. For capital (رأس مال) or withdrawals use the \
partner named in the message.
- method: "من جيبه" when the sender paid from his own pocket (من جيبي، من فلوسي); otherwise the payment channel \
mentioned; default "كاش الشركة".
- date: ISO date. "امبارح" = yesterday. Default today.
- details: a short Arabic description of the item (e.g. "فراخ 2 كيلو"), without the amount.
- If the message is not about money at all, return an empty list."""


def _tool(settings: Settings):
    cats = [c.name for c in settings.categories]
    people = settings.partners or ["عمر"]
    return {
        "name": "record_entries",
        "description": "Record the ledger entries found in the message.",
        "input_schema": {
            "type": "object",
            "properties": {
                "entries": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {
                            "amount": {"type": ["number", "null"]},
                            "category": {"type": ["string", "null"], "enum": cats + [None]},
                            "details": {"type": "string"},
                            "qty": {"type": ["number", "null"]},
                            "date": {"type": "string", "description": "YYYY-MM-DD"},
                            "method": {"type": "string", "enum": layout.METHODS},
                            "person": {"type": "string", "enum": people},
                        },
                        "required": ["amount", "category", "details", "date", "method", "person"],
                    },
                }
            },
            "required": ["entries"],
        },
    }


class LLMParser:
    def __init__(self, api_key: str, model: str):
        import anthropic
        self.client = anthropic.Anthropic(api_key=api_key)
        self.model = model

    def parse(self, text: str, sender: str, today: dt.date, settings: Settings) -> list[Entry]:
        cat_lines = "\n".join(f"- {c.name} [{c.cls}]: {', '.join(c.keywords[:25])}" for c in settings.categories)
        user = (f"Today: {today.isoformat()} ({today.strftime('%A')})\nSender: {sender}\n"
                f"Partners: {', '.join(settings.partners)}\n\nCategories [class]: keywords\n{cat_lines}\n\n"
                f"Message:\n{text}")
        resp = self.client.messages.create(
            model=self.model, max_tokens=1024, system=SYSTEM,
            tools=[_tool(settings)], tool_choice={"type": "tool", "name": "record_entries"},
            messages=[{"role": "user", "content": user}],
        )
        block = next(b for b in resp.content if getattr(b, "type", "") == "tool_use")
        data = block.input if isinstance(block.input, dict) else json.loads(block.input)
        names = {c.name for c in settings.categories}
        out = []
        for e in data.get("entries", []):
            try:
                when = dt.date.fromisoformat(e.get("date") or today.isoformat())
            except ValueError:
                when = today
            if when > today + dt.timedelta(days=1):
                when = today
            cat = e.get("category")
            amount = e.get("amount")
            out.append(Entry(
                amount=float(amount) if isinstance(amount, (int, float)) and amount > 0 else None,
                category=cat if cat in names else None,
                details=(e.get("details") or "").strip(),
                qty=e.get("qty") if isinstance(e.get("qty"), (int, float)) else None,
                date=when,
                method=e.get("method") if e.get("method") in layout.METHODS else layout.DEFAULT_METHOD,
                person=e.get("person") if e.get("person") in settings.partners else sender,
            ))
        return out
