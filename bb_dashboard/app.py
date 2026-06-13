"""
bb_dashboard/app.py — дашборд рентабельности ББ Урал Комсомол.
Читает данные из Google Sheets, отдаёт HTML + JSON API.

Env vars:
  GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET, GOOGLE_REFRESH_TOKEN
  SPREADSHEET_ID (опционально, по умолчанию вшит)
"""
import os, re, time
from flask import Flask, jsonify, send_from_directory
from flask_cors import CORS
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

app = Flask(__name__, static_folder="static")
CORS(app)

SPREADSHEET_ID = os.environ.get(
    "SPREADSHEET_ID", "1v7FQF7NfXINPp6yiy9DF8JPiVwywYoIkz3Ej8aWUYtg"
)
CACHE_TTL = 300
_cache = {"data": None, "ts": 0}

MATERIAL_SECTIONS = {
    "Оборудование", "Воздуховоды", "Фасонные части", "Клапаны",
    "Диффузоры и решётки", "Изоляция", "Расходники", "Позиции не из ВДЦ",
}


def get_service():
    creds = Credentials(
        token=None,
        refresh_token=os.environ["GOOGLE_REFRESH_TOKEN"],
        client_id=os.environ["GOOGLE_CLIENT_ID"],
        client_secret=os.environ["GOOGLE_CLIENT_SECRET"],
        token_uri="https://oauth2.googleapis.com/token",
    )
    creds.refresh(Request())
    return build("sheets", "v4", credentials=creds, cache_discovery=False)


def parse_num(s):
    try:
        return float(re.sub(r"[^\d,.-]", "", str(s).replace("\xa0", "")).replace(",", "."))
    except Exception:
        return 0.0


def fetch_data():
    svc = get_service()
    rows = svc.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range="'Рентабельность'",
        valueRenderOption="FORMATTED_VALUE",
    ).execute().get("values", [])

    sections, systems = [], []
    mat_total = {"vdc": 0, "seb": 0, "spent": 0}
    rab_total = {"seb": 0, "spent": 0}
    overall   = {"vdc": 0, "seb": 0, "spent": 0}
    updated   = ""
    mode      = "mat"

    for row in rows:
        label = row[0].strip() if row else ""
        if label == "Работы (монтаж)":
            mode = "work"; continue
        if label == "ИТОГО":
            overall = {k: parse_num(row[i]) if len(row) > i else 0
                       for k, i in [("vdc",1),("seb",2),("spent",3)]}
            mode = "done"; continue
        if label == "По состоянию на:" and len(row) > 1:
            updated = row[1].strip(); continue
        if label == "ИТОГО материалы":
            mat_total = {k: parse_num(row[i]) if len(row) > i else 0
                         for k, i in [("vdc",1),("seb",2),("spent",3)]}
            continue
        if label == "ИТОГО работы":
            rab_total = {k: parse_num(row[i]) if len(row) > i else 0
                         for k, i in [("seb",2),("spent",3)]}
            continue
        if mode == "mat" and label in MATERIAL_SECTIONS:
            sections.append({
                "name":  label,
                "vdc":   parse_num(row[1]) if len(row) > 1 else 0,
                "seb":   parse_num(row[2]) if len(row) > 2 else 0,
                "spent": parse_num(row[3]) if len(row) > 3 else 0,
            })
        elif mode == "work" and label and label not in ("Раздел",):
            systems.append({
                "name":  label,
                "seb":   parse_num(row[2]) if len(row) > 2 else 0,
                "spent": parse_num(row[3]) if len(row) > 3 else 0,
            })

    vdc   = overall.get("vdc") or mat_total.get("vdc", 0)
    total_seb   = overall.get("seb", 0)
    total_spent = overall.get("spent", 0)
    margin      = vdc - total_seb if vdc else 0
    mat_pct = mat_total["spent"] / mat_total["seb"] * 100 if mat_total["seb"] else 0
    rab_pct = rab_total["spent"] / rab_total["seb"] * 100 if rab_total["seb"] else 0

    return {
        "updated": updated,
        "materials": {
            "vdc_plan": mat_total["vdc"], "seb_plan": mat_total["seb"],
            "spent": mat_total["spent"], "pct": round(mat_pct, 1),
            "sections": sections,
        },
        "work": {
            "seb_plan": rab_total["seb"], "spent": rab_total["spent"],
            "pct": round(rab_pct, 1), "systems": systems,
        },
        "total": {
            "spent": total_spent, "seb_plan": total_seb, "vdc_plan": vdc,
            "margin": margin, "margin_pct": round(margin / vdc * 100, 1) if vdc else 0,
        },
    }


@app.route("/api/data")
def api_data():
    now = time.time()
    if not _cache["data"] or now - _cache["ts"] > CACHE_TTL:
        try:
            _cache["data"] = fetch_data()
            _cache["ts"] = now
        except Exception as e:
            if _cache["data"]:
                pass
            else:
                return jsonify({"error": str(e)}), 500
    return jsonify(_cache["data"])


@app.route("/")
def index():
    return send_from_directory("static", "index.html")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5001)))
