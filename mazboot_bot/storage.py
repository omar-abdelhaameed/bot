"""Storage backends: Google Sheets (production) and a local .xlsx file (offline / testing).

Both expose the same small interface. The workbook does all accounting with formulas;
the bot only appends/deletes ledger rows and reads values back.
"""
from __future__ import annotations

import datetime as dt
import logging
import os
import shutil
import subprocess
import tempfile
import threading
from abc import ABC, abstractmethod

from . import layout
from .models import Category, Settings, Tx

log = logging.getLogger(__name__)
EPOCH = dt.date(1899, 12, 30)


# ----------------------------------------------------------------- value helpers
def to_date(v) -> dt.date | None:
    if v is None or v == "":
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, (int, float)):
        return EPOCH + dt.timedelta(days=int(v))
    s = str(v).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def to_datetime(v) -> dt.datetime | None:
    if v is None or v == "":
        return None
    if isinstance(v, dt.datetime):
        return v
    if isinstance(v, (int, float)):
        return dt.datetime.combine(EPOCH, dt.time()) + dt.timedelta(days=float(v))
    try:
        return dt.datetime.fromisoformat(str(v))
    except ValueError:
        return None


def to_float(v) -> float | None:
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", "").replace("%", "").strip())
    except ValueError:
        return None


def serial(d: dt.date) -> int:
    return (d - EPOCH).days


def serial_dt(d: dt.datetime) -> float:
    return (d - dt.datetime.combine(EPOCH, dt.time())).total_seconds() / 86400


def safe_text(s: str) -> str:
    """Stop a message like '=IMPORTXML(...)' from becoming a formula in the sheet."""
    s = (s or "").strip()
    return "'" + s if s[:1] in ("=", "+", "-", "@") else s


def parse_settings(start, opening, alert, partner_rows, cat_rows) -> Settings:
    cats = []
    for row in cat_rows:
        row = list(row) + [None] * (5 - len(row))
        name, cls, _drv, _val, kw = row[:5]
        if not name or not str(name).strip():
            continue
        keywords = [k.strip() for k in str(kw or "").replace("،", ",").split(",") if k.strip()]
        cats.append(Category(str(name).strip(), str(cls or "").strip(), keywords))
    partners = [str(r[0]).strip() for r in partner_rows if r and r[0] and str(r[0]).strip()]
    a = to_float(alert)
    if a is not None and a > 1:
        a = a / 100
    return Settings(start_month=to_date(start) or dt.date.today().replace(day=1),
                    opening_cash=to_float(opening) or 0.0, alert_pct=a or 0.8,
                    partners=partners, categories=cats)


def row_to_tx(rownum: int, vals: list) -> Tx | None:
    vals = list(vals) + [None] * (len(layout.TX_COLS) - len(vals))
    d = dict(zip(layout.TX_COLS, vals))
    amount = to_float(d["amount"])
    if amount is None:
        return None
    tid = to_float(d["id"])
    return Tx(row=rownum, id=int(tid) if tid is not None else None, date=to_date(d["date"]),
              category=str(d["category"] or "").strip(), details=str(d["details"] or ""),
              qty=to_float(d["qty"]), amount=amount, person=str(d["person"] or "").strip(),
              method=str(d["method"] or "").strip(), source=str(d["source"] or "").strip(),
              logged_at=to_datetime(d["logged_at"]), notes=str(d["notes"] or ""))


def parse_planned(month_row: list, block: list) -> dict[dt.date, dict[str, float]]:
    months = [to_date(v) for v in month_row]
    out: dict[dt.date, dict[str, float]] = {m: {} for m in months if m}
    for row in block:
        if not row or not row[0]:
            continue
        name = str(row[0]).strip()
        values = list(row[2:]) if len(row) > 2 else []
        for m, v in zip(months, values):
            if m:
                out[m][name] = to_float(v) or 0.0
    return out


class Storage(ABC):
    @abstractmethod
    def read_settings(self) -> Settings: ...

    @abstractmethod
    def read_transactions(self) -> list[Tx]: ...

    @abstractmethod
    def append_rows(self, rows: list[list]) -> None:
        """rows: [id, date(date), category, details, qty, amount, person, method, source, logged_at(datetime), notes]"""

    @abstractmethod
    def delete_tx(self, tx_id: int) -> bool: ...

    @abstractmethod
    def read_planned(self) -> dict[dt.date, dict[str, float]]: ...

    @abstractmethod
    def export_xlsx(self) -> bytes: ...


# ================================================================= Google Sheets
class GoogleSheetsStorage(Storage):
    def __init__(self, sheet_id: str, credentials_file: str, credentials_json: str = ""):
        import gspread
        self._gspread = gspread
        if credentials_json:
            import json
            self.gc = gspread.service_account_from_dict(json.loads(credentials_json))
        else:
            self.gc = gspread.service_account(filename=credentials_file)
        self.sh = self.gc.open_by_key(sheet_id)
        self._ws = {}

    def ws(self, name):
        if name not in self._ws:
            self._ws[name] = self.sh.worksheet(name)
        return self._ws[name]

    def _get(self, tab, rng):
        return self.ws(tab).get(rng, value_render_option="UNFORMATTED_VALUE")

    def read_settings(self) -> Settings:
        ranges = [f"'{layout.TAB_SETTINGS}'!{r}" for r in (
            layout.SET_START_MONTH, layout.SET_OPENING_CASH, layout.SET_ALERT_PCT,
            layout.SET_PARTNERS, layout.SET_CATEGORIES)]
        res = self.sh.values_batch_get(ranges, params={"valueRenderOption": "UNFORMATTED_VALUE"})
        vr = [r.get("values", []) for r in res["valueRanges"]]
        one = lambda v: v[0][0] if v and v[0] else None
        return parse_settings(one(vr[0]), one(vr[1]), one(vr[2]), vr[3], vr[4])

    def read_transactions(self) -> list[Tx]:
        rows = self._get(layout.TAB_TX, f"A{layout.TX_FIRST_ROW}:{layout.TX_LAST_COL}")
        out = []
        for i, r in enumerate(rows):
            tx = row_to_tx(layout.TX_FIRST_ROW + i, r)
            if tx:
                out.append(tx)
        return out

    def _last_used_row(self) -> int:
        col = self.ws(layout.TAB_TX).get(f"A{layout.TX_FIRST_ROW}:F", value_render_option="UNFORMATTED_VALUE")
        last = layout.TX_FIRST_ROW - 1
        for i, r in enumerate(col):
            if any(v not in (None, "") for v in r):
                last = layout.TX_FIRST_ROW + i
        return last

    def append_rows(self, rows: list[list]) -> None:
        ws = self.ws(layout.TAB_TX)
        start = self._last_used_row() + 1
        end = start + len(rows) - 1
        if end > ws.row_count:
            ws.add_rows(end - ws.row_count + 100)
        values = []
        for r in rows:
            r = list(r)
            r[1] = serial(r[1])
            r[9] = round(serial_dt(r[9]), 6)
            r[3] = safe_text(r[3])
            r[10] = safe_text(r[10])
            values.append(["" if v is None else v for v in r])
        ws.update(values=values, range_name=f"A{start}:{layout.TX_LAST_COL}{end}", value_input_option="USER_ENTERED")
        ws.batch_format([
            {"range": f"B{start}:B{end}", "format": {"numberFormat": {"type": "DATE", "pattern": "yyyy-mm-dd"}}},
            {"range": f"J{start}:J{end}", "format": {"numberFormat": {"type": "DATE_TIME", "pattern": "yyyy-mm-dd hh:mm"}}},
            {"range": f"F{start}:F{end}", "format": {"numberFormat": {"type": "NUMBER", "pattern": "#,##0"}}},
        ])

    def delete_tx(self, tx_id: int) -> bool:
        for tx in self.read_transactions():
            if tx.id == tx_id:
                self.ws(layout.TAB_TX).delete_rows(tx.row)
                return True
        return False

    def read_planned(self):
        b = layout
        months = self._get(b.TAB_BUDGET, f"{b.BUD_FIRST_MONTH_COL}{b.BUD_MONTH_ROW}:{b.BUD_LAST_MONTH_COL}{b.BUD_MONTH_ROW}")
        block = self._get(b.TAB_BUDGET, f"B{b.BUD_PLANNED_FIRST}:{b.BUD_LAST_MONTH_COL}{b.BUD_PLANNED_LAST}")
        return parse_planned(months[0] if months else [], block)

    def export_xlsx(self) -> bytes:
        return self.sh.export(format=self._gspread.utils.ExportFormat.EXCEL)


# ================================================================= local Excel file
class ExcelStorage(Storage):
    """Writes straight into a local .xlsx (keep it closed in Excel while the bot runs).

    openpyxl drops cached formula results when it saves, so planned budget values are read
    from a LibreOffice-recalculated temporary copy (if LibreOffice is installed) and cached;
    they only depend on the inputs, not on the ledger.
    """

    def __init__(self, path: str, soffice: str | None = None):
        self.path = path
        self.soffice = soffice or shutil.which("soffice") or shutil.which("libreoffice")
        self._lock = threading.Lock()
        self._planned_cache = None
        self._planned_mtime = None
        self._own_mtime = None

    def _load(self, data_only=False):
        from openpyxl import load_workbook
        import warnings
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            return load_workbook(self.path, data_only=data_only)

    def read_settings(self) -> Settings:
        wb = self._load(data_only=True)
        ws = wb[layout.TAB_SETTINGS]
        rows = lambda rng: [[c.value for c in row] for row in ws[rng]]
        return parse_settings(ws[layout.SET_START_MONTH].value, ws[layout.SET_OPENING_CASH].value,
                              ws[layout.SET_ALERT_PCT].value, rows(layout.SET_PARTNERS),
                              rows(layout.SET_CATEGORIES))

    def read_transactions(self) -> list[Tx]:
        wb = self._load(data_only=True)
        ws = wb[layout.TAB_TX]
        out = []
        for i, row in enumerate(ws.iter_rows(min_row=layout.TX_FIRST_ROW, max_col=len(layout.TX_COLS),
                                             values_only=True)):
            tx = row_to_tx(layout.TX_FIRST_ROW + i, row)
            if tx:
                out.append(tx)
        return out

    @staticmethod
    def _last_used_row(ws) -> int:
        last = layout.TX_FIRST_ROW - 1
        for i, row in enumerate(ws.iter_rows(min_row=layout.TX_FIRST_ROW, max_col=6, values_only=True)):
            if any(v not in (None, "") for v in row):
                last = layout.TX_FIRST_ROW + i
        return last

    def _save(self, wb):
        d = os.path.dirname(os.path.abspath(self.path))
        fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=d)
        os.close(fd)
        wb.save(tmp)
        os.replace(tmp, self.path)          # PermissionError if the file is open in Excel (Windows)
        self._own_mtime = os.path.getmtime(self.path)

    def append_rows(self, rows: list[list]) -> None:
        with self._lock:
            self.read_planned()             # cache planned values before cached results are dropped
            wb = self._load()
            ws = wb[layout.TAB_TX]
            r0 = self._last_used_row(ws) + 1
            for k, r in enumerate(rows):
                for j, v in enumerate(r):
                    cell = ws.cell(row=r0 + k, column=j + 1, value=v if v != "" else None)
                    if j == 1:
                        cell.number_format = "yyyy-mm-dd"
                    elif j == 9:
                        cell.number_format = "yyyy-mm-dd hh:mm"
                    elif j == 5:
                        cell.number_format = '#,##0;(#,##0);"-"'
            self._save(wb)

    def delete_tx(self, tx_id: int) -> bool:
        with self._lock:
            self.read_planned()
            wb = self._load()
            ws = wb[layout.TAB_TX]
            for i, row in enumerate(ws.iter_rows(min_row=layout.TX_FIRST_ROW, max_col=1, values_only=True)):
                v = to_float(row[0])
                if v is not None and int(v) == tx_id:
                    r = layout.TX_FIRST_ROW + i
                    last = self._last_used_row(ws)
                    # shift rows up inside the ledger instead of delete_rows, so formula ranges stay intact
                    for rr in range(r, last):
                        for c in range(1, len(layout.TX_COLS) + 1):
                            src, dst = ws.cell(rr + 1, c), ws.cell(rr, c)
                            dst.value = src.value
                            dst.number_format = src.number_format
                    for c in range(1, len(layout.TX_COLS) + 1):
                        ws.cell(last, c).value = None
                    self._save(wb)
                    return True
            return False

    def _recalculated_copy(self) -> str | None:
        if not self.soffice:
            return None
        out = tempfile.mkdtemp(prefix="mazboot_")
        try:
            subprocess.run([self.soffice, "--headless", "--convert-to", "xlsx", "--outdir", out, self.path],
                           capture_output=True, timeout=120, check=False)
            p = os.path.join(out, os.path.basename(self.path))
            return p if os.path.exists(p) else None
        except Exception as e:     # pragma: no cover
            log.warning("LibreOffice recalculation failed: %s", e)
            return None

    def read_planned(self):
        mtime = os.path.getmtime(self.path)
        external_change = self._planned_mtime is not None and mtime != self._planned_mtime and mtime != self._own_mtime
        if self._planned_cache is not None and not external_change:
            return self._planned_cache
        planned = self._read_planned_from(self.path)
        if not any(v for m in planned.values() for v in m.values()):
            copy = self._recalculated_copy()
            if copy:
                planned = self._read_planned_from(copy)
                shutil.rmtree(os.path.dirname(copy), ignore_errors=True)
        self._planned_cache, self._planned_mtime = planned, mtime
        return planned

    @staticmethod
    def _read_planned_from(path):
        from openpyxl import load_workbook
        import warnings
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            wb = load_workbook(path, data_only=True)
        ws = wb[layout.TAB_BUDGET]
        b = layout
        months = [c.value for c in ws[f"{b.BUD_FIRST_MONTH_COL}{b.BUD_MONTH_ROW}:{b.BUD_LAST_MONTH_COL}{b.BUD_MONTH_ROW}"][0]]
        block = [[c.value for c in row] for row in ws[f"B{b.BUD_PLANNED_FIRST}:{b.BUD_LAST_MONTH_COL}{b.BUD_PLANNED_LAST}"]]
        return parse_planned(months, block)

    def export_xlsx(self) -> bytes:
        with open(self.path, "rb") as f:
            return f.read()
