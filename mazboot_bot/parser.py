from __future__ import annotations

import datetime as dt
import logging

from .models import Entry, Settings
from .parser_rules import RuleParser

log = logging.getLogger(__name__)


class Parser:
    """AI parser when an API key is configured, rule parser otherwise (and as fallback)."""

    def __init__(self, llm=None):
        self.llm = llm
        self._rules: RuleParser | None = None
        self._rules_for: int | None = None

    def rules(self, settings: Settings) -> RuleParser:
        if self._rules is None or self._rules_for != id(settings):
            self._rules, self._rules_for = RuleParser(settings), id(settings)
        return self._rules

    def parse(self, text: str, sender: str, today: dt.date, settings: Settings) -> tuple[list[Entry], str]:
        if self.llm is not None:
            try:
                entries = self.llm.parse(text, sender, today, settings)
                if entries:
                    return entries, "ai"
            except Exception as e:      # network, quota, bad output... never block bookkeeping
                log.warning("AI parser failed, using rules: %s", e)
        return self.rules(settings).parse(text, sender, today), "rules"
