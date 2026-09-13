"""Разворачивает worker.mjs в Cloudflare Workers.

Запуск с ПК (секреты только через окружение, в git ничего не попадает):

    set CF_API_TOKEN=...
    set CF_ACCOUNT_ID=...
    set AI_KEY=...              # ключ OpenRouter, живёт только в воркере
    set RELAY_SECRET=...        # если не задан — берётся из relay.local.json или генерируется
    python deploy/relay/deploy.py

После деплоя печатает базовый URL, который нужно положить в .env бота:
    TELEGRAM_API_BASE=<url>
    GROQ_BASE_URL=<url>/ai
"""

import json
import os
import secrets
import urllib.error
import urllib.request
import uuid
from pathlib import Path

API = "https://api.cloudflare.com/client/v4"
SCRIPT_NAME = "invoice-relay"
HERE = Path(__file__).resolve().parent
WORKER = HERE / "worker.mjs"
LOCAL_CFG = HERE / "relay.local.json"  # в .gitignore

TOKEN = os.environ["CF_API_TOKEN"]
ACCOUNT = os.environ["CF_ACCOUNT_ID"]
AI_KEY = os.environ["AI_KEY"]


def cf(path, method="GET", body=None, raw=None, content_type=None):
    data = None
    headers = {"Authorization": "Bearer " + TOKEN}
    if raw is not None:
        data = raw
        headers["Content-Type"] = content_type
    elif body is not None:
        data = json.dumps(body).encode()
        headers["Content-Type"] = "application/json"
    req = urllib.request.Request(API + path, data=data, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req, timeout=90) as r:
            return r.status, json.load(r)
    except urllib.error.HTTPError as e:
        return e.code, json.loads(e.read().decode("utf-8", "replace") or "{}")


def relay_secret() -> str:
    env = (os.environ.get("RELAY_SECRET") or "").strip()
    if env:
        return env
    if LOCAL_CFG.exists():
        saved = json.loads(LOCAL_CFG.read_text(encoding="utf-8")).get("relay_secret")
        if saved:
            return saved
    return secrets.token_urlsafe(24)


def upload(secret: str):
    metadata = {
        "main_module": "worker.mjs",
        "compatibility_date": "2025-09-01",
        "bindings": [
            {"type": "secret_text", "name": "RELAY_SECRET", "text": secret},
            {"type": "secret_text", "name": "AI_KEY", "text": AI_KEY},
            {"type": "plain_text", "name": "AI_BASE",
             "text": os.environ.get("AI_BASE", "https://openrouter.ai/api/v1")},
        ],
    }
    boundary = "----relay" + uuid.uuid4().hex
    parts = []

    def add(name, content, filename=None, ctype="application/json"):
        disp = 'form-data; name="%s"' % name
        if filename:
            disp += '; filename="%s"' % filename
        parts.append(
            ("--%s\r\nContent-Disposition: %s\r\nContent-Type: %s\r\n\r\n" % (boundary, disp, ctype)).encode()
            + content
            + b"\r\n"
        )

    add("metadata", json.dumps(metadata).encode())
    add("worker.mjs", WORKER.read_bytes(), filename="worker.mjs", ctype="application/javascript+module")
    payload = b"".join(parts) + ("--%s--\r\n" % boundary).encode()

    return cf(
        "/accounts/%s/workers/scripts/%s" % (ACCOUNT, SCRIPT_NAME),
        "PUT",
        raw=payload,
        content_type="multipart/form-data; boundary=" + boundary,
    )


def main():
    secret = relay_secret()

    st, d = upload(secret)
    print("upload:", st, "ok" if d.get("success") else d.get("errors"))
    if not d.get("success"):
        raise SystemExit(1)

    # именно POST: PUT на этом эндпоинте отдаёт 10405 для API-токенов
    st, d = cf("/accounts/%s/workers/scripts/%s/subdomain" % (ACCOUNT, SCRIPT_NAME), "POST",
               {"enabled": True, "previews_enabled": False})
    print("workers.dev route:", st, "ok" if d.get("success") else d.get("errors"))

    st, d = cf("/accounts/%s/workers/subdomain" % ACCOUNT)
    sub = (d.get("result") or {}).get("subdomain")
    base = "https://%s.%s.workers.dev/r/%s" % (SCRIPT_NAME, sub, secret)

    LOCAL_CFG.write_text(json.dumps({"relay_secret": secret, "base_url": base,
                                     "script": SCRIPT_NAME, "subdomain": sub}, indent=2), encoding="utf-8")

    print()
    print("TELEGRAM_API_BASE=" + base)
    print("GROQ_BASE_URL=" + base + "/ai")
    print()
    print("сохранено в", LOCAL_CFG)


if __name__ == "__main__":
    main()
