"""Arabic text normalisation helpers (Egyptian colloquial friendly)."""
import re

_TASHKEEL = re.compile(r"[ؐ-ًؚ-ٰٟۖ-ۭـ]")   # harakat + tatweel

_MAP = {}
for i, d in enumerate("٠١٢٣٤٥٦٧٨٩"):
    _MAP[d] = str(i)
for i, d in enumerate("۰۱۲۳۴۵۶۷۸۹"):
    _MAP[d] = str(i)
_MAP.update({"أ": "ا", "إ": "ا", "آ": "ا", "ٱ": "ا", "ى": "ي", "ة": "ه", "٫": ".", "٬": ",", "،": ","})


def strip_marks(text: str) -> str:
    """Remove harakat and tatweel. Changes length — do this first."""
    return _TASHKEEL.sub("", text or "")


def norm_same_len(text: str) -> str:
    """1:1 character normalisation, so spans found in the result apply to the input."""
    return "".join(_MAP.get(ch, ch) for ch in text).lower()


def norm(text: str) -> str:
    return norm_same_len(strip_marks(text)).strip()


AR_LETTER = "ء-ي"
# common attached prefixes in Egyptian Arabic: و ب ل ف ك ال بال وال لل عال فال
PREFIX = r"(?:و|ف)?(?:ب|ل|ك)?(?:ال|ل)?"
SUFFIX = r"(?:ات|ه|ها|ين|ي|ك)?"


def keyword_pattern(keyword: str) -> re.Pattern:
    kw = norm(keyword)
    body = r"\s+".join(re.escape(p) for p in kw.split())
    return re.compile(rf"(?<![{AR_LETTER}a-z0-9]){PREFIX}{body}{SUFFIX}(?![{AR_LETTER}a-z0-9])")
