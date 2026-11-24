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
        google_credentials_json=os.environ.get("GOOGLE_CREDENTIALS_JSON", ""),
    )
