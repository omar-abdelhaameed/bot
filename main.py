"""Run the MAZBOOT finance bot:  python main.py"""
import logging
import sys

from mazboot_bot import config
from mazboot_bot.bot import MazbootBot
from mazboot_bot.ledger import Ledger
from mazboot_bot.parser import Parser
from mazboot_bot.storage import ExcelStorage, GoogleSheetsStorage


def main():
    logging.basicConfig(format="%(asctime)s %(levelname)s %(name)s: %(message)s", level=logging.INFO)
    logging.getLogger("httpx").setLevel(logging.WARNING)
    cfg = config.load()
    if not cfg.telegram_token:
        sys.exit("TELEGRAM_BOT_TOKEN is missing (.env)")
    if not cfg.users:
        sys.exit("ALLOWED_USERS is missing (.env) — e.g. ALLOWED_USERS=123456789:عمر,987654321:أحمد")

    if cfg.backend == "excel":
        storage = ExcelStorage(cfg.excel_path)
    else:
        if not cfg.sheet_id:
            sys.exit("GOOGLE_SHEET_ID is missing (.env)")
        storage = GoogleSheetsStorage(cfg.sheet_id, cfg.google_credentials, cfg.google_credentials_json)

    llm = None
    if cfg.anthropic_key:
        from mazboot_bot.parser_llm import LLMParser
        llm = LLMParser(cfg.anthropic_key, cfg.model)

    ledger = Ledger(storage)
    s = ledger.settings()
    unknown = [n for n in cfg.users.values() if n not in s.partners]
    if unknown:
        logging.warning("These ALLOWED_USERS names are not partners in the settings tab: %s", unknown)
    logging.info("Loaded %d categories, partners %s, backend=%s, parser=%s",
                 len(s.categories), s.partners, cfg.backend, "ai+rules" if llm else "rules")

    MazbootBot(cfg, ledger, Parser(llm)).build().run_polling(drop_pending_updates=False)


if __name__ == "__main__":
    main()
