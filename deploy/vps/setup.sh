#!/usr/bin/env bash
# На VPS: sudo bash /opt/invoice_bot/deploy/vps/setup.sh
set -euo pipefail

APP_DIR="${APP_DIR:-/opt/invoice_bot}"
SERVICE_NAME="invoice-bot"
PYTHON_BIN="${PYTHON_BIN:-python3}"

echo "==> packages"
apt-get update -y
apt-get install -y "$PYTHON_BIN" "${PYTHON_BIN}-venv" git curl

cd "$APP_DIR"
if [[ ! -d .git ]]; then
  echo "ERROR: нет git в $APP_DIR"
  exit 1
fi
if [[ ! -f .env ]]; then
  echo "ERROR: нет $APP_DIR/.env"
  exit 1
fi

echo "==> venv + deps"
$PYTHON_BIN -m venv .venv
# shellcheck disable=SC1091
source .venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt

echo "==> systemd"
cp deploy/vps/invoice-bot.service /etc/systemd/system/${SERVICE_NAME}.service
systemctl daemon-reload
systemctl enable "${SERVICE_NAME}"
systemctl restart "${SERVICE_NAME}"

sleep 2
systemctl --no-pager --full status "${SERVICE_NAME}" || true
curl -fsS "http://127.0.0.1:8081/" || true
echo
echo "OK. health: http://127.0.0.1:8081/"
