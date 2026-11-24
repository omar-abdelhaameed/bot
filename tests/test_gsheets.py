"""GoogleSheetsStorage against an in-memory fake of the gspread API (no network)."""
import datetime as dt
import os
import re
import sys
import types
import warnings

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from mazboot_bot import storage as S          # noqa: E402

_HERE = os.path.dirname(os.path.abspath(__file__))
WB = os.environ.get("WORKBOOK") or next(
    (p for p in (os.path.join(_HERE, "..", "MAZBOOT_Finance.xlsx"), os.path.join(_HERE, "..", "..", "out", "MAZBOOT_Finance.xlsx"))
     if os.path.exists(p)), "MAZBOOT_Finance.xlsx")


def _ser(v):
    if isinstance(v, dt.datetime):
        return S.serial_dt(v)
    if isinstance(v, dt.date):
        return S.serial(v)
    return v


class FakeWS:
    def __init__(self, ws):
        self.cells = {}
        for row in ws.iter_rows():
            for c in row:
                if c.value is not None:
                    self.cells[(c.row, c.column)] = _ser(c.value)
        self.row_count = 10000
        self.formats = []

    def _rng(self, a1):
        a1 = a1.split("!")[-1]
        m = re.fullmatch(r"([A-Z]+)(\d+)(?::([A-Z]+)(\d+)?)?", a1)
        c1, r1 = column_index_from_string(m.group(1)), int(m.group(2))
        c2 = column_index_from_string(m.group(3)) if m.group(3) else c1
        r2 = int(m.group(4)) if m.group(4) else (self.row_count if m.group(3) else r1)
        return r1, c1, r2, c2

    def get(self, a1, value_render_option=None):
        r1, c1, r2, c2 = self._rng(a1)
        out = []
        for r in range(r1, r2 + 1):
            row = [self.cells.get((r, c), "") for c in range(c1, c2 + 1)]
            while row and row[-1] == "":
                row.pop()
            out.append(row)
        while out and not out[-1]:
            out.pop()
        return out

    def update(self, values, range_name, value_input_option=None):
        r1, c1, _, _ = self._rng(range_name)
        for i, row in enumerate(values):
            for j, v in enumerate(row):
                self.cells[(r1 + i, c1 + j)] = v

    def batch_format(self, fmts):
        self.formats += fmts

    def delete_rows(self, r):
        new = {}
        for (rr, cc), v in self.cells.items():
            if rr < r:
                new[(rr, cc)] = v
            elif rr > r:
                new[(rr - 1, cc)] = v
        self.cells = new

    def add_rows(self, n):
        self.row_count += n


class FakeSH:
    def __init__(self, path):
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            wb = load_workbook(path, data_only=True)
        self.tabs = {ws.title: FakeWS(ws) for ws in wb.worksheets}

    def worksheet(self, name):
        return self.tabs[name]

    def values_batch_get(self, ranges, params=None):
        res = []
        for r in ranges:
            tab, a1 = r.rsplit("!", 1)
            res.append({"values": self.tabs[tab.strip("'")].get(a1)})
        return {"valueRanges": res}

    def export(self, format=None):
        return b"PK-fake"


def make(monkeypatch):
    sh = FakeSH(WB)
    fake = types.SimpleNamespace(service_account=lambda filename: types.SimpleNamespace(open_by_key=lambda k: sh),
                                 utils=types.SimpleNamespace(ExportFormat=types.SimpleNamespace(EXCEL="xlsx")))
    monkeypatch.setitem(sys.modules, "gspread", fake)
    return S.GoogleSheetsStorage("id", "creds.json"), sh


def test_gsheets_roundtrip(monkeypatch):
    st, sh = make(monkeypatch)
    s = st.read_settings()
    assert s.start_month == dt.date(2026, 11, 1) and s.partners == ["عمر", "أسامة", "أحمد"]
    assert s.category("خامات وبضاعة").cls == "تكلفة مباشرة" and "فراخ" in s.category("خامات وبضاعة").keywords
    planned = st.read_planned()
    assert planned[dt.date(2026, 11, 1)]["إيجار"] == 4500
    now = dt.datetime(2026, 11, 22, 10, 0)
    st.append_rows([[1, dt.date(2026, 11, 22), "خامات وبضاعة", "=HYPERLINK(1)", "", 5000, "عمر", "كاش الشركة", "بوت", now, ""],
                    [2, dt.date(2026, 11, 22), "تغليف", "علب", 10, 900, "عمر", "كاش الشركة", "بوت", now, ""]])
    tx = st.read_transactions()
    assert [(t.id, t.row, t.date, t.amount) for t in tx] == [(1, 5, dt.date(2026, 11, 22), 5000), (2, 6, dt.date(2026, 11, 22), 900)]
    assert tx[0].details.startswith("'=")                    # formula injection neutralised
    st.append_rows([[3, dt.date(2026, 11, 23), "إيجار", "إيجار", "", 4500, "عمر", "كاش الشركة", "بوت", now, ""]])
    assert st.read_transactions()[-1].row == 7
    assert st.delete_tx(2) and [t.id for t in st.read_transactions()] == [1, 3]
    assert st.export_xlsx() == b"PK-fake"
