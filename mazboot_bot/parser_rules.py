"""Rule-based parser for messages like «صرفت 5000 على البضاعة» or «جبت فراخ ب1200 ورز ب300».

Works offline with no API key. Categories and their keywords come from the settings tab,
so adding a keyword in the sheet teaches the bot a new word.
"""
from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass

from . import layout
from .models import Entry, Settings
from .textnorm import AR_LETTER, keyword_pattern, norm, norm_same_len, strip_marks

B = rf"(?<![{AR_LETTER}a-z])"          # left word boundary
E = rf"(?![{AR_LETTER}a-z])"           # right word boundary

WORD_NUMBERS = {
    "الف": 1000, "الفين": 2000, "مليون": 1_000_000,
    "ميه": 100, "مائه": 100, "ميت": 100, "ميتين": 200, "تلتميه": 300, "ثلاثمائه": 300, "ربعميه": 400,
    "اربعمائه": 400, "خمسميه": 500, "خمسمائه": 500, "ستميه": 600, "سبعميه": 700, "تمنميه": 800,
    "تسعميه": 900, "خمسين": 50, "ميتين وخمسين": 250,
}
NUM_RE = re.compile(r"(?<![\d.])(\d{1,3}(?:,\d{3})+|\d+(?:\.\d+)?)")
WORD_NUM_RE = re.compile(B + "(" + "|".join(sorted(map(re.escape, WORD_NUMBERS), key=len, reverse=True)) + ")" + E)
MULT_RE = re.compile(r"\s*(k(?![a-z])|الاف|الف|الآف|مليون)")
HALF_RE = re.compile(r"\s*و?\s*نص" + E)
ORDER_UNIT_RE = re.compile(r"\s*(اوردرات|اوردر|طلبات|طلب|وجبات|وجبه|orders?)" + E)
QTY_UNIT_RE = re.compile(
    r"\s*(كيلوجرام|كيلو|كجم|كج|kg|جرام|جم|gm|لتر|علبه|علب|كرتونه|كراتين|قطعه|قطع|حته|حتت|شيكاره|كيس|اكياس|"
    r"عبوه|عبوات|طبق|بيضه|بيض|رغيف|ارغفه|واحد|حبات|حبه)" + E)
CURRENCY_RE = re.compile(r"\s*(جنيهات|جنيه|ج\.م|ج|egp|le|pounds?)" + E)
DATE_RE = re.compile(r"(?<![\d.])(\d{1,2})[/\-](\d{1,2})(?:[/\-](\d{2,4}))?(?![\d.])"
                     r"|(?<![\d.])(\d{1,2})\.(\d{1,2})\.(\d{2,4})(?![\d.])")
YESTERDAY_RE = re.compile(B + r"(امبارح|امس|امبارحه)" + E)
DAY_BEFORE_RE = re.compile(B + r"(اول\s+امبارح|اول\s+امس)" + E)

METHOD_PATTERNS = [
    (re.compile(r"من\s+(جيبي|جيبه|جيبى|فلوسي|فلوسه)|علي\s+حسابي|على\s+حسابي|من\s+عندي"), "من جيبه"),
    (re.compile(B + r"(فودافون|فودا\s*كاش|vodafone|vf\s*cash)"), "فودافون كاش"),
    (re.compile(B + r"(انستاباي|انستا\s*باي|instapay|انستا)" + E), "إنستاباي"),
    (re.compile(B + r"(تحويل|بنكي|البنك|حواله)" + E), "تحويل بنكي"),
    (re.compile(B + r"(كاش|نقدي|نقدا)" + E), "كاش الشركة"),
]

FILLERS = set(norm(w) for w in (
    "صرفت صرفنا صرف دفعت دفعنا دفع جبت جبنا جاب اشتريت اشترينا اشتري اشترى خدت خدنا على عل علي في فى من ب "
    "و ال حق تمن ثمن سعر مبلغ بمبلغ بقيمه قيمه جنيه جنيهات ج النهارده انهارده اليوم امبارح امس كمان برضه "
    "بس يعني عشان علشان حوالي تقريبا كاش للشركه الشركه نقدي حطيت حط دخل بعنا بعت بيع اوردر اوردرات فودافون فودا انستاباي انستا تحويل بنكي البنك جيبي جيبه فلوسي عندي").split())


@dataclass
class _Num:
    start: int
    end: int
    value: float
    kind: str            # amount | qty | orders
    unit: str = ""
    thousands_word: bool = False


def _to_float(s: str) -> float:
    return float(s.replace(",", ""))


def _scan_numbers(n: str) -> list[_Num]:
    nums: list[_Num] = []
    taken = [False] * len(n)

    def mark(a, b):
        for i in range(a, b):
            taken[i] = True

    for m in NUM_RE.finditer(n):
        val = _to_float(m.group(1))
        end = m.end()
        thousands = False
        mm = MULT_RE.match(n, end)
        if mm:
            word = mm.group(1)
            val *= 1000 if word in ("k", "الاف", "الف", "الآف") else 1_000_000
            end = mm.end()
            thousands = True
            h = HALF_RE.match(n, end)
            if h:
                val += 500 if word != "مليون" else 500_000
                end = h.end()
        nums.append(_Num(m.start(), end, val, "amount", thousands_word=thousands))
        mark(m.start(), end)
    for m in WORD_NUM_RE.finditer(n):
        if any(taken[m.start():m.end()]):
            continue
        val = float(WORD_NUMBERS[m.group(1)])
        end = m.end()
        thousands = val >= 1000
        if thousands:
            h = HALF_RE.match(n, end)
            if h:
                val += val / 2 if val == 1000 else 500
                end = h.end()
        nums.append(_Num(m.start(), end, val, "amount", thousands_word=thousands))
        mark(m.start(), end)
    nums.sort(key=lambda x: x.start)

    # «3 آلاف و 500» / «الف و خمسميه» -> one number
    merged: list[_Num] = []
    for x in nums:
        if merged:
            p = merged[-1]
            gap = n[p.end:x.start]
            if p.thousands_word and re.fullmatch(r"\s*و\s*", gap) and x.value < 1000:
                p.value += x.value
                p.end = x.end
                continue
        merged.append(x)

    for x in merged:
        um = ORDER_UNIT_RE.match(n, x.end)
        if um:
            x.kind, x.unit, x.end = "orders", um.group(1), um.end()
            continue
        um = QTY_UNIT_RE.match(n, x.end)
        if um:
            x.kind, x.unit, x.end = "qty", um.group(1), um.end()
            continue
        cm = CURRENCY_RE.match(n, x.end)
        if cm:
            x.end = cm.end()
    return merged


class RuleParser:
    def __init__(self, settings: Settings):
        self.settings = settings
        self._patterns = []
        for order, cat in enumerate(settings.categories):
            for kw in [cat.name] + cat.keywords:
                if kw and kw.strip():
                    self._patterns.append((keyword_pattern(kw), len(norm(kw)), order, cat.name))
        self._partners = [(keyword_pattern(p), p) for p in settings.partners if p]
        self._kw_norm = {norm(k) for c in settings.categories for k in [c.name] + c.keywords if k}

    # ----------------------------------------------------------------- helpers
    def find_category(self, text_norm: str) -> str | None:
        best = None
        for pat, length, order, name in self._patterns:
            if pat.search(text_norm):
                key = (length, -order)
                if best is None or key > best[0]:
                    best = (key, name)
        return best[1] if best else None

    def find_person(self, text_norm: str) -> str | None:
        for pat, name in self._partners:
            if pat.search(text_norm):
                return name
        return None

    @staticmethod
    def find_method(text_norm: str) -> str | None:
        for pat, method in METHOD_PATTERNS:
            if pat.search(text_norm):
                return method
        return None

    @staticmethod
    def find_date(text_norm: str, today: dt.date) -> tuple[dt.date, list[tuple[int, int]]]:
        spans = []
        if DAY_BEFORE_RE.search(text_norm):
            return today - dt.timedelta(days=2), spans
        if YESTERDAY_RE.search(text_norm):
            return today - dt.timedelta(days=1), spans
        m = DATE_RE.search(text_norm)
        if m:
            g = m.groups()
            d, mo, y = (int(g[0]), int(g[1]), g[2]) if g[0] else (int(g[3]), int(g[4]), g[5])
            year = today.year if not y else (int(y) + 2000 if len(y) == 2 else int(y))
            try:
                when = dt.date(year, mo, d)
                if not y and when > today + dt.timedelta(days=31):
                    when = dt.date(year - 1, mo, d)
                spans.append((m.start(), m.end()))
                return when, spans
            except ValueError:
                pass
        return today, spans

    def _clean_details(self, segment_base: str) -> str:
        words = re.split(r"[\s,.:؛;!?()\-]+", segment_base)
        kept = []
        for w in words:
            if not w:
                continue
            wn = norm(w)
            if wn in FILLERS or re.fullmatch(r"[\d.,]+", wn):
                continue
            if wn.startswith("ب") and re.fullmatch(r"ب[\d.,]+", wn):
                continue
            if not kept and len(wn) > 2 and wn.startswith("و") and wn[1:] in self._kw_norm:
                w = w[1:]
            kept.append(w)
        return " ".join(kept).strip()

    # ----------------------------------------------------------------- main
    def parse(self, text: str, sender: str, today: dt.date) -> list[Entry]:
        base = strip_marks(text)
        n = norm_same_len(base)
        when, date_spans = self.find_date(n, today)
        scan = n
        for a, b in date_spans:
            scan = scan[:a] + " " * (b - a) + scan[b:]
        nums = _scan_numbers(scan)
        blank = list(base)
        for a, b in date_spans + [(x.start, x.end) for x in nums]:
            blank[a:b] = " " * (b - a)
        base = "".join(blank)
        method = self.find_method(n) or layout.DEFAULT_METHOD
        person = self.find_person(n) or sender

        amounts = [x for x in nums if x.kind == "amount"]
        orders = [x for x in nums if x.kind == "orders"]

        if not amounts:
            cat = self.find_category(n)
            qty = orders[0].value if orders else None
            return [Entry(amount=None, category=cat, details=self._clean_details(base) or (cat or ""),
                          qty=qty, date=when, method=method, person=person)]

        # orientation: item-then-price («فراخ ب1200») vs price-then-item («5000 على البضاعة»)
        first = amounts[0]
        before_first = n[:first.start]
        after_mode = self.find_category(before_first) is None

        entries: list[Entry] = []
        for i, a in enumerate(amounts):
            prev_end = amounts[i - 1].end if i else 0
            next_start = amounts[i + 1].start if i + 1 < len(amounts) else len(n)
            before = (prev_end, a.start)
            after = (a.end, next_start)
            primary, secondary = (after, before) if after_mode else (before, after)
            seg = primary
            cat = self.find_category(n[seg[0]:seg[1]])
            if cat is None:
                seg = secondary
                cat = self.find_category(n[seg[0]:seg[1]])
            if cat is None and len(amounts) == 1:
                cat = self.find_category(n)
                seg = (0, len(n))
            if cat is None:
                seg = primary
            details = self._clean_details(base if len(amounts) == 1 else base[seg[0]:seg[1]])
            qty = None
            for q in nums:
                if q.kind == "qty" and (len(amounts) == 1 or seg[0] <= q.start < seg[1]):
                    qty = q.value
                    details = f"{details} ({_fmt(q.value)} {q.unit})".strip()
                    break
            entries.append(Entry(amount=a.value, category=cat, details=details or (cat or ""), qty=qty,
                                 date=when, method=method, person=person))

        # sales: attach the order count
        cats = {c.name: c for c in self.settings.categories}
        for e in entries:
            c = cats.get(e.category or "")
            if c and c.cls == layout.CLS_INCOME and orders and e.qty is None:
                e.qty = orders[0].value
        return entries


def _fmt(x: float) -> str:
    return str(int(x)) if float(x).is_integer() else str(x)
