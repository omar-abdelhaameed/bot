from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass


def _load_dotenv(path=".env"):
    if not os.path.exists(path):
        return
    with open(path, encoding="utf-8-sig") as f:  # tolerate BOM (Windows editors)
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


@dataclass
class Config:
    telegram_token: str
    users: dict[int, str]            # telegram user id -> partner name (as written in the settings tab)
    backend: str                     # "gsheets" | "excel"
    sheet_id: str
    google_credentials: str
    excel_path: str
    anthropic_key: str
    model: str
    timezone: str
    notify_chat: int | None          # optional group chat that gets every new entry
    google_credentials_json: str = ""   # the service-account JSON itself (for Railway/Render style hosting)
    google_credentials_source: str = ""  # where google_credentials_json came from (for the startup log)


_GCP_FIELDS = ("type", "project_id", "private_key_id", "private_key", "client_email", "client_id",
               "auth_uri", "token_uri", "auth_provider_x509_cert_url", "client_x509_cert_url",
               "universe_domain")
_PREFIXES = ("GCP_", "GOOGLE_", "")   # GCP_PRIVATE_KEY, GOOGLE_PRIVATE_KEY, PRIVATE_KEY, private_key


def _clean(v: str) -> str:
    """Undo common copy/paste damage: surrounding quotes, a trailing comma from the JSON line."""
    v = v.strip()
    if v.endswith(","):
        v = v[:-1].rstrip()
    if len(v) >= 2 and v[0] == v[-1] and v[0] in "\"'":
        v = v[1:-1]
    return v.strip()


def _env_field(field: str) -> str:
    for p in _PREFIXES:
        for name in (p + field.upper(), p + field):
            v = os.environ.get(name, "")
            if v.strip():
                return _clean(v)
    return ""


def _normalize_pem(key: str) -> str:
    """Rebuild a PEM key whatever happened to its line breaks (literal \\n, real newlines,
    newlines turned into spaces by a single-line input box, or the key split in parts)."""
    key = key.replace("\\n", "\n")
    m = re.search(r"-----BEGIN ([A-Z ]+)-----(.*?)-----END \1-----", key, re.S)
    if not m:
        return key
    label, body = m.group(1), re.sub(r"[^A-Za-z0-9+/=]", "", m.group(2).replace("\\n", ""))
    lines = [body[i:i + 64] for i in range(0, len(body), 64)]
    return f"-----BEGIN {label}-----\n" + "\n".join(lines) + f"\n-----END {label}-----\n"


def _private_key() -> str:
    key = _env_field("private_key")
    if not key:   # key split over several variables: GCP_PRIVATE_KEY_1, GCP_PRIVATE_KEY_2, ...
        for p in _PREFIXES:
            parts = []
            for i in range(1, 30):
                v = os.environ.get(f"{p}PRIVATE_KEY_{i}", "") or os.environ.get(f"{p}private_key_{i}", "")
                if not v.strip():
                    break
                parts.append(_clean(v))
            if parts:
                key = "\n".join(parts)
                break
    return _normalize_pem(key) if key else ""


def _credentials_json() -> tuple[str, str]:
    """Service-account JSON + where it came from.

    1. GOOGLE_CREDENTIALS_JSON: the whole JSON in one variable.
    2. Separate variables, one per JSON field, for hosts that cap each value's length:
       GCP_PRIVATE_KEY + GCP_CLIENT_EMAIL (+ the other fields, optional). GOOGLE_ or no prefix,
       and lowercase names, work too; the key may also be split into GCP_PRIVATE_KEY_1, _2, ...
    Returns ("", "") when neither is set (then the credentials file is used).
    """
    raw = os.environ.get("GOOGLE_CREDENTIALS_JSON", "").strip()
    if raw:
        raw = _clean(raw) if not raw.startswith("{") else raw
        try:
            info = json.loads(raw, strict=False)   # strict=False: tolerate real newlines inside the key
        except json.JSONDecodeError as e:
            raise SystemExit(f"GOOGLE_CREDENTIALS_JSON is not valid JSON ({e.msg} at char {e.pos}). "
                             "Paste the whole service-account file content, starting with { and ending with }.")
        if info.get("private_key"):
            info["private_key"] = _normalize_pem(info["private_key"])
        return json.dumps(info), "GOOGLE_CREDENTIALS_JSON"

    key = _private_key()
    if not key:
        return "", ""
    info = {f: v for f in _GCP_FIELDS if f != "private_key" and (v := _env_field(f))}
    info["private_key"] = key
    info.setdefault("type", "service_account")
    info.setdefault("token_uri", "https://oauth2.googleapis.com/token")
    if not info.get("client_email"):
        raise SystemExit("Found the private key variable but no client email. Add GCP_CLIENT_EMAIL "
                         "(the client_email from the service-account JSON).")
    return json.dumps(info), "separate variables (" + ", ".join(sorted(info)) + ")"


def credential_env_names() -> list[str]:
    """Names (never values) of env vars that look credential-related, for error messages."""
    hints = ("GCP", "GOOGLE", "PRIVATE", "CLIENT", "PROJECT", "TOKEN_URI", "CERT")
    return sorted(n for n in os.environ if any(h in n.upper() for h in hints))

def load() -> Config:
    _load_dotenv()
    users = {}
    for part in os.environ.get("ALLOWED_USERS", "").split(","):
        if ":" in part:
            uid, name = part.split(":", 1)
            users[int(uid.strip())] = name.strip()
    notify = os.environ.get("NOTIFY_CHAT_ID", "").strip()
    creds_json, creds_source = _credentials_json()
    return Config(
        telegram_token=os.environ.get("TELEGRAM_BOT_TOKEN", ""),
        users=users,
        backend=os.environ.get("STORAGE", "gsheets").strip().lower(),
        sheet_id=os.environ.get("GOOGLE_SHEET_ID", ""),
        google_credentials=os.environ.get("GOOGLE_CREDENTIALS_FILE", "service_account.json"),
        excel_path=os.environ.get("EXCEL_PATH", "MAZBOOT_Finance.xlsx"),
        anthropic_key=os.environ.get("ANTHROPIC_API_KEY", ""),
        model=os.environ.get("CLAUDE_MODEL", "claude-haiku-5-5"),
        timezone=os.environ.get("TIMEZONE", "Africa/Cairo"),
        notify_chat=int(notify) if notify else None,
        google_credentials_json=creds_json,
        google_credentials_source=creds_source,
    )
