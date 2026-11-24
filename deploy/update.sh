#!/bin/bash
# Push local code + secrets to Google Cloud and restart the bot VM.
# Run from the project folder in Git Bash:   bash deploy/update.sh
# Secrets are read from the local .env / service_account.json and stored in
# Secret Manager; they are never printed.
set -euo pipefail
cd "$(dirname "$0")/.."

PROJECT=mazboot-515700
ZONE=us-central1-a
VM=mazboot-bot
BUCKET=mazboot-515700-code
GCLOUD="${GCLOUD:-gcloud}"
command -v "$GCLOUD" >/dev/null 2>&1 || GCLOUD="$LOCALAPPDATA/Google/Cloud SDK/google-cloud-sdk/bin/gcloud.cmd"

g() { "$GCLOUD" --project "$PROJECT" --quiet "$@"; }

echo "== packaging code"
tar --exclude='__pycache__' -czf "${TMPDIR:-/tmp}/mazboot.tar.gz" main.py requirements.txt mazboot_bot
g storage cp "${TMPDIR:-/tmp}/mazboot.tar.gz" "gs://$BUCKET/mazboot.tar.gz"

echo "== updating secrets (only if changed)"
for pair in "mazboot-env:.env" "mazboot-service-account:service_account.json"; do
  name=${pair%%:*}; file=${pair#*:}
  current=$(g secrets versions access latest --secret="$name" 2>/dev/null | md5sum | cut -d' ' -f1 || true)
  if [ "$current" != "$(md5sum < "$file" | cut -d' ' -f1)" ]; then
    g secrets versions add "$name" --data-file="$file" >/dev/null && echo "   $name: new version"
  else
    echo "   $name: unchanged"
  fi
done

echo "== restarting VM (startup script pulls the new code and secrets)"
g compute instances reset "$VM" --zone "$ZONE"
echo "done. Logs:  gcloud compute instances get-serial-port-output $VM --zone $ZONE --project $PROJECT | grep -E 'mazboot|python'"
