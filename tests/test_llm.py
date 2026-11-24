"""AI parser output handling with a mocked Anthropic client (no network)."""
import datetime as dt
import json
import os
import sys
import types

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from mazboot_bot.parser import Parser          # noqa: E402
from mazboot_bot.storage import ExcelStorage   # noqa: E402

_HERE = os.path.dirname(os.path.abspath(__file__))
WB = os.environ.get("WORKBOOK") or next(
    (p for p in (os.path.join(_HERE, "..", "MAZBOOT_Finance.xlsx"), os.path.join(_HERE, "..", "..", "out", "MAZBOOT_Finance.xlsx"))
     if os.path.exists(p)), "MAZBOOT_Finance.xlsx")


def test_llm_parser(monkeypatch):
    captured = {}

    class Msgs:
        def create(self, **kw):
            captured.update(kw)
            data = {"entries": [
                {"amount": 1200, "category": "خامات وبضاعة", "details": "فراخ", "qty": 5, "date": "2026-11-21",
                 "method": "من جيبه", "person": "عمر"},
                {"amount": 300, "category": "حاجة مش موجودة", "details": "رز", "qty": None, "date": "2030-01-01",
                 "method": "بيتكوين", "person": "مجهول"}]}
            return types.SimpleNamespace(content=[types.SimpleNamespace(type="tool_use", input=data)])

    fake = types.SimpleNamespace(Anthropic=lambda api_key: types.SimpleNamespace(messages=Msgs()))
    monkeypatch.setitem(sys.modules, "anthropic", fake)
    from mazboot_bot.parser_llm import LLMParser
    s = ExcelStorage(WB).read_settings()
    entries, how = Parser(LLMParser("k", "claude-haiku-5-5")).parse("جبت فراخ ب1200 ورز ب300", "أحمد",
                                                                    dt.date(2026, 11, 22), s)
    assert how == "ai"
    assert (entries[0].amount, entries[0].category, entries[0].qty, entries[0].method, entries[0].date) == \
        (1200, "خامات وبضاعة", 5, "من جيبه", dt.date(2026, 11, 21))
    # invalid values are cleaned instead of trusted
    assert entries[1].category is None and entries[1].method == "كاش الشركة" and entries[1].person == "أحمد"
    assert entries[1].date == dt.date(2026, 11, 22)
    tool = captured["tools"][0]
    assert "خامات وبضاعة" in tool["input_schema"]["properties"]["entries"]["items"]["properties"]["category"]["enum"]
    json.dumps(tool)                     # schema is serialisable


def test_llm_failure_falls_back(monkeypatch):
    class Boom:
        def parse(self, *a):
            raise RuntimeError("no network")
    s = ExcelStorage(WB).read_settings()
    entries, how = Parser(Boom()).parse("صرفت 5000 على البضاعة", "عمر", dt.date(2026, 11, 22), s)
    assert how == "rules" and entries[0].amount == 5000
