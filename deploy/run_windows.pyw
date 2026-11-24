"""Windowless supervisor for the bot on Windows (started by the Task Scheduler task "MAZBOOT Bot").

Runs `python main.py` from the project folder, appends its output to logs/bot.log,
and restarts it 15 seconds after it exits for any reason (crash, network loss, ...).
The token is never written to the log (httpx request logging is disabled in main.py).
"""
import os
import subprocess
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PYTHON = os.path.join(ROOT, ".venv", "Scripts", "python.exe")
LOG = os.path.join(ROOT, "logs", "bot.log")
os.makedirs(os.path.dirname(LOG), exist_ok=True)

env = dict(os.environ, PYTHONUTF8="1", PYTHONIOENCODING="utf-8", PYTHONUNBUFFERED="1")
while True:
    if os.path.exists(LOG) and os.path.getsize(LOG) > 5 * 1024 * 1024:   # keep the log small
        os.replace(LOG, LOG + ".old")
    with open(LOG, "a", encoding="utf-8") as log:
        log.write(f"\n=== supervisor: starting bot {time.strftime('%Y-%m-%d %H:%M:%S')} ===\n")
        log.flush()
        code = subprocess.call([PYTHON, "main.py"], cwd=ROOT, env=env, stdout=log, stderr=subprocess.STDOUT,
                               creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        log.write(f"=== supervisor: bot exited with code {code}, restarting in 15s ===\n")
    time.sleep(15)
