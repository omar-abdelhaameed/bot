from __future__ import annotations

import os
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


_GCP_FIELDS = ("type", "project_id", "private_key_id", "private_key", "client_email", "client_id",
               "auth_uri", "token_uri", "auth_provider_x509_cert_url", "client_x509_cert_url",
               "universe_domain")

def _credentials_json() -> str:
    """Service-account JSON: GOOGLE_CREDENTIALS_JSON as-is, else rebuilt from GCP_<FIELD> variables
    (GCP_PRIVATE_KEY, GCP_CLIENT_EMAIL, ...) for hosts that cap each env value at ~2000 chars."""
    raw = os.environ.get("GOOGLE_CREDENTIALS_JSON", "").strip()
    if raw:
        return raw
    key = os.environ.get("GCP_PRIVATE_KEY", "").strip()
    if not key:
        return ""
    import json
    info = {}
    for field in _GCP_FIELDS:
        val = os.environ.get("GCP_" + field.upper(), "").strip()
        if val:
            info[field] = val
    info["private_key"] = key.replace("\\n", "\n")
    info.setdefault("type", "service_account")
    info.setdefault("token_uri", "https://oauth2.googleapis.com/token")
    return json.dumps(info)

def load() -> Config:
    _load_dotenv()
    users = {}
    for part in os.environ.get("ALLOWED_USERS", "").split(","):
        if ":" in part:
            uid, name = part.split(":", 1)
            users[int(uid.strip())] = name.strip()
    notify = os.environ.get("NOTIFY_CHAT_ID", "").strip()
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
        google_credentials_json=_credentials_json(),
    )
