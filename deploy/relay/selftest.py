"""Проверка ретранслятора: сам воркер, Telegram через него, OpenRouter через него.

    set TELEGRAM_TOKEN=...        # опционально, для проверки getMe
    python deploy/relay/selftest.py
"""

import json
import os
import urllib.error
import urllib.request
from pathlib import Path

CFG = Path(__file__).resolve().parent / "relay.local.json"
BASE = json.loads(CFG.read_text(encoding="utf-8"))["base_url"]


def get(url, timeout=40):
    try:
        with urllib.request.urlopen(url, timeout=timeout) as r:
            return r.status, r.read()[:300].decode("utf-8", "replace")
    except urllib.error.HTTPError as e:
        return e.code, e.read()[:300].decode("utf-8", "replace")
    except Exception as e:
        return "ERR", str(e)[:200]


def post(url, body, timeout=120):
    req = urllib.request.Request(url, data=json.dumps(body).encode(),
                                 headers={"Content-Type": "application/json",
                                          "Authorization": "Bearer dummy"})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return r.status, json.load(r)
    except urllib.error.HTTPError as e:
        return e.code, e.read()[:400].decode("utf-8", "replace")
    except Exception as e:
        return "ERR", str(e)[:200]


print("1. worker alive:", get(BASE))

# secret обязателен
bad = BASE.rsplit("/r/", 1)[0] + "/r/wrong-secret"
print("2. wrong secret:", get(bad))

tok = (os.environ.get("TELEGRAM_TOKEN") or "").strip()
if tok:
    st, body = get("%s/bot%s/getMe" % (BASE, tok))
    print("3. telegram getMe:", st, body[:160])
else:
    print("3. telegram getMe: пропущено (нет TELEGRAM_TOKEN)")

st, d = post(BASE + "/ai/openai/v1/chat/completions", {
    "model": "google/gemma-4-31b-it:free",
    "messages": [{"role": "user", "content": "Ответь одним словом: OK"}],
    "max_tokens": 16,
    "reasoning_effort": "low",       # groq-only, воркер должен вырезать
    "include_reasoning": False,
})
if isinstance(d, dict) and "choices" in d:
    print("4. openrouter chat:", st, repr(d["choices"][0]["message"]["content"][:60]))
else:
    print("4. openrouter chat:", st, d)
