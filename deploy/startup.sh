#!/bin/bash
# VM startup script (runs as root on every boot of the GCE VM).
# Pulls the latest code from the private bucket and the secrets from Secret Manager,
# then (re)starts the bot as a systemd service. No secrets live in this file.
set -euo pipefail

APP=/opt/mazboot
MD=http://metadata.google.internal/computeMetadata/v1/instance/attributes
BUCKET=$(curl -sf -H 'Metadata-Flavor: Google' "$MD/code-bucket")

echo "[mazboot] startup: bucket=$BUCKET"

# 1 GB swap: e2-micro only has 1 GB RAM.
if [ ! -f /swapfile ]; then
  fallocate -l 1G /swapfile && chmod 600 /swapfile && mkswap /swapfile && swapon /swapfile
  echo '/swapfile none swap sw 0 0' >> /etc/fstab
fi

if ! command -v python3 >/dev/null || ! python3 -c 'import venv, ensurepip' 2>/dev/null; then
  apt-get update -y && apt-get install -y python3-venv
fi

id mazboot >/dev/null 2>&1 || useradd --system --home "$APP" --shell /usr/sbin/nologin mazboot
mkdir -p "$APP"

# Code
gcloud storage cp "gs://$BUCKET/mazboot.tar.gz" /tmp/mazboot.tar.gz --quiet
rm -rf "$APP/mazboot_bot"
tar -xzf /tmp/mazboot.tar.gz -C "$APP"

# Secrets -> files readable only by the bot user
umask 077
gcloud secrets versions access latest --secret=mazboot-env > "$APP/.env"
gcloud secrets versions access latest --secret=mazboot-service-account > "$APP/service_account.json"
umask 022

# Python deps (only reinstall when requirements.txt changes)
if [ ! -x "$APP/.venv/bin/python" ]; then python3 -m venv "$APP/.venv"; fi
if ! cmp -s "$APP/requirements.txt" "$APP/.venv/.installed-requirements" 2>/dev/null; then
  "$APP/.venv/bin/pip" install --no-cache-dir -q --upgrade pip
  "$APP/.venv/bin/pip" install --no-cache-dir -q -r "$APP/requirements.txt"
  cp "$APP/requirements.txt" "$APP/.venv/.installed-requirements"
fi

chown -R mazboot:mazboot "$APP"
chmod 600 "$APP/.env" "$APP/service_account.json"

cat > /etc/systemd/system/mazboot-bot.service <<'EOF'
[Unit]
Description=MAZBOOT finance Telegram bot
After=network-online.target
Wants=network-online.target

[Service]
User=mazboot
WorkingDirectory=/opt/mazboot
Environment=PYTHONUNBUFFERED=1 PYTHONUTF8=1
ExecStart=/opt/mazboot/.venv/bin/python main.py
Restart=always
RestartSec=10
# logs go to journald and to the serial console (readable with get-serial-port-output)
StandardOutput=journal+console
StandardError=journal+console

[Install]
WantedBy=multi-user.target
EOF

systemctl daemon-reload
systemctl enable mazboot-bot
systemctl restart mazboot-bot
echo "[mazboot] startup: bot service restarted"
