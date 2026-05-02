import os
import io
import json
import re
import base64
import time
from datetime import datetime

import requests
import telebot
from groq import Groq
from flask import Flask, request as flask_request
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from apscheduler.schedulers.background import BackgroundScheduler

import pdfplumber
import openpyxl

TELEGRAM_TOKEN = os.environ["TELEGRAM_TOKEN"]
GROQ_API_KEY = os.environ["GROQ_API_KEY"]
GOOGLE_CLIENT_ID = os.environ["GOOGLE_CLIENT_ID"]
GOOGLE_CLIENT_SECRET = os.environ["GOOGLE_CLIENT_SECRET"]
GOOGLE_SHEETS_REFRESH_TOKEN = os.environ["GOOGLE_SHEETS_REFRESH_TOKEN"]
WEBHOOK_URL = os.environ["WEBHOOK_URL"]
PORT = int(os.environ.get("PORT", 8080))
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "1tmuj1f2D2euUZlr-CXHzgkurRavF-MV8UxvjG3SIKDc")

INVOICES_SHEET = "Счета"
REGISTRY_SHEET = "Реестр"

INVOICES_HEADERS = ["№", "Поставщик", "Номер счёта", "Дата счёта", "Позиция в счете",
                    "Наименование", "Артикул/Описание", "Ед.изм.", "Кол-во",
                    "Цена с НДС", "Сумма с НДС", "Дата добавления", "Общая сумма с НДС в счете", "Имя файла"]

REGISTRY_HEADERS = ["Имя файла", "Статус", "Получен", "Обработан", "Ошибка", "file_id", "file_type", "chat_id"]

bot = telebot.TeleBot(TELEGRAM_TOKEN)
client = Groq(api_key=GROQ_API_KEY)
app = Flask(__name__)


# ── Google Sheets ──────────────────────────────────────────────────────────────

def get_sheets_service():
    creds = Credentials(
        token=None,
        refresh_token=GOOGLE_SHEETS_REFRESH_TOKEN,
        client_id=GOOGLE_CLIENT_ID,
        client_secret=GOOGLE_CLIENT_SECRET,
        token_uri="https://oauth2.googleapis.com/token",
    )
    creds.refresh(Request())
    return build("sheets", "v4", credentials=creds)


def ensure_sheets():
    service = get_sheets_service()
    meta = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    existing = {s["properties"]["title"] for s in meta["sheets"]}

    requests_body = []
    if INVOICES_SHEET not in existing:
        requests_body.append({"addSheet": {"properties": {"title": INVOICES_SHEET}}})
    if REGISTRY_SHEET not in existing:
        requests_body.append({"addSheet": {"properties": {"title": REGISTRY_SHEET}}})

    if requests_body:
        service.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"requests": requests_body}
        ).execute()

    # Write headers if sheets are new/empty
    for sheet, headers in [(INVOICES_SHEET, INVOICES_HEADERS), (REGISTRY_SHEET, REGISTRY_HEADERS)]:
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID, range=f"{sheet}!A1"
        ).execute()
        if not result.get("values"):
            service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=f"{sheet}!A1",
                valueInputOption="RAW",
                body={"values": [headers]},
            ).execute()

    print("✅ Таблица готова", flush=True)


def sheet_append(sheet: str, rows: list):
    service = get_sheets_service()
    service.spreadsheets().values().append(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{sheet}!A1",
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body={"values": rows},
    ).execute()


def sheet_get_all(sheet: str) -> list:
    service = get_sheets_service()
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{sheet}!A:Z",
    ).execute()
    return result.get("values", [])


def sheet_update_cell(sheet: str, row: int, col: int, value: str):
    service = get_sheets_service()
    col_letter = chr(ord("A") + col - 1)
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{sheet}!{col_letter}{row}",
        valueInputOption="RAW",
        body={"values": [[value]]},
    ).execute()


# ── Registry ───────────────────────────────────────────────────────────────────

def registry_find_row(filename: str) -> int | None:
    rows = sheet_get_all(REGISTRY_SHEET)
    for i, row in enumerate(rows[1:], start=2):
        if row and row[0] == filename:
            return i
    return None


def registry_is_done(filename: str) -> bool:
    rows = sheet_get_all(REGISTRY_SHEET)
    for row in rows[1:]:
        if row and row[0] == filename and len(row) > 1 and row[1] == "✅":
            return True
    return False


def registry_add(filename: str, file_id: str, file_type: str, chat_id: int):
    now = datetime.now().strftime("%d.%m.%Y %H:%M")
    sheet_append(REGISTRY_SHEET, [[filename, "⏳", now, "", "", file_id, file_type, str(chat_id)]])


def registry_update(row_num: int, status: str, error: str = ""):
    now = datetime.now().strftime("%d.%m.%Y %H:%M") if status == "✅" else ""
    sheet_update_cell(REGISTRY_SHEET, row_num, 2, status)
    sheet_update_cell(REGISTRY_SHEET, row_num, 4, now)
    sheet_update_cell(REGISTRY_SHEET, row_num, 5, error)


def registry_get_failed() -> list[dict]:
    rows = sheet_get_all(REGISTRY_SHEET)
    failed = []
    for i, row in enumerate(rows[1:], start=2):
        if len(row) >= 8 and row[1] == "❌":
            failed.append({
                "row": i,
                "filename": row[0],
                "file_id": row[5],
                "file_type": row[6],
                "chat_id": int(row[7]),
            })
    return failed


# ── File parsing ───────────────────────────────────────────────────────────────

def extract_text_from_pdf(data: bytes) -> str:
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        pages = []
        for page in pdf.pages:
            text = page.extract_text() or ""
            tables = page.extract_tables()
            if tables:
                for table in tables:
                    for row in table:
                        if row:
                            text += "\n" + "\t".join(str(c) if c else "" for c in row)
            pages.append(text)
    return "\n\n---\n\n".join(pages)


def extract_text_from_excel(data: bytes) -> str:
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    lines = []
    for sheet in wb.worksheets:
        lines.append(f"=== {sheet.title} ===")
        for row in sheet.iter_rows(values_only=True):
            if any(v is not None for v in row):
                lines.append("\t".join(str(v) if v is not None else "" for v in row))
    return "\n".join(lines)


EXTRACT_PROMPT = """Ты парсер счетов-фактур и накладных.
Извлеки данные из счёта и верни ТОЛЬКО валидный JSON без пояснений:

{
  "supplier": "Название поставщика",
  "invoice_number": "Номер счёта",
  "invoice_date": "ДД.ММ.ГГГГ или пусто если нет",
  "total_amount": 0.0,
  "items": [
    {"pos": 1, "name": "Наименование", "article": "артикул/тип или пусто", "unit": "шт", "quantity": 1, "price_with_vat": 0.0}
  ]
}

Правила:
- supplier: кто ВЫСТАВИЛ счёт (продавец/исполнитель). НЕ покупатель и НЕ ООО "Термодинамика" — они всегда покупатель
- total_amount: итоговая сумма ВСЕГО счёта с НДС (строка "Итого")
- pos: номер позиции в счёте (1, 2, 3...)
- article: артикул, тип, марка, производитель (если есть)
- price_with_vat: ЦЕНА ЗА ЕДИНИЦУ товара с НДС. Это НЕ сумма строки! Если в счёте колонки "Цена" и "Сумма" — бери значение из колонки "Цена", не из "Сумма". Если цена без НДС — умножь на 1.2
- quantity: количество единиц товара
- Если поле неизвестно — пустая строка или 0"""


def extract_json(text: str) -> dict:
    start = text.find("{")
    if start == -1:
        raise ValueError("JSON не найден")
    depth = 0
    for i, ch in enumerate(text[start:], start):
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                return json.loads(text[start:i + 1])
    raise ValueError("Незакрытый JSON")


def parse_with_ai(text: str) -> dict:
    for attempt in range(3):
        try:
            resp = client.chat.completions.create(
                model="llama-3.1-8b-instant",
                messages=[
                    {"role": "system", "content": EXTRACT_PROMPT},
                    {"role": "user", "content": f"Счёт:\n\n{text[:2000]}"},
                ],
                max_tokens=2000,
                temperature=0.1,
            )
            raw = resp.choices[0].message.content.strip()
            print(f"🤖 AI: {raw[:300]}", flush=True)
            return extract_json(raw)
        except Exception as e:
            if ("429" in str(e) or "413" in str(e)) and attempt < 2:
                wait = 30 * (attempt + 1)
                print(f"⏳ Rate limit/size, жду {wait}с", flush=True)
                time.sleep(wait)
            else:
                raise


def parse_image_with_ai(image_bytes: bytes) -> dict:
    b64 = base64.b64encode(image_bytes).decode()
    resp = client.chat.completions.create(
        model="meta-llama/llama-4-scout-17b-16e-instruct",
        messages=[{
            "role": "user",
            "content": [
                {"type": "text", "text": EXTRACT_PROMPT + "\n\nИзвлеки из фото счёта:"},
                {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64}"}},
            ],
        }],
        max_tokens=2000,
        temperature=0.1,
    )
    return extract_json(resp.choices[0].message.content.strip())


def invoice_to_rows(data: dict, filename: str) -> list:
    today = datetime.now().strftime("%d.%m.%Y")
    supplier = data.get("supplier", "—")
    inv_num = data.get("invoice_number", "—")
    inv_date = data.get("invoice_date", "—")
    total_amount = data.get("total_amount", 0) or 0
    rows = []
    for i, item in enumerate(data.get("items", []), start=1):
        qty = item.get("quantity", 0) or 0
        price = item.get("price_with_vat", 0) or 0
        try:
            line_total = round(float(qty) * float(price), 2)
        except Exception:
            line_total = 0
        rows.append([
            i, filename, supplier, inv_num, inv_date,
            item.get("pos", i), item.get("name", "—"), item.get("article", ""),
            item.get("unit", "—"), qty, price, line_total,
            today, total_amount,
        ])
    return rows


# ── Core processing ────────────────────────────────────────────────────────────

def download_file(file_id: str) -> bytes:
    file_info = bot.get_file(file_id)
    url = f"https://api.telegram.org/file/bot{TELEGRAM_TOKEN}/{file_info.file_path}"
    return requests.get(url).content


def set_reaction(msg, emoji: str):
    try:
        bot.set_message_reaction(
            msg.chat.id, msg.message_id,
            [telebot.types.ReactionTypeEmoji(emoji)],
        )
    except Exception as e:
        print(f"⚠️ Reaction error: {e}", flush=True)


def process_file(file_id: str, file_type: str, filename: str) -> dict:
    file_bytes = download_file(file_id)
    if file_type == "pdf":
        text = extract_text_from_pdf(file_bytes)
        return parse_with_ai(text)
    elif file_type == "excel":
        text = extract_text_from_excel(file_bytes)
        return parse_with_ai(text)
    elif file_type == "photo":
        return parse_image_with_ai(file_bytes)
    else:
        raise ValueError(f"Неизвестный тип: {file_type}")


def process_invoice(msg, file_id: str, file_type: str, filename: str):
    print(f"📥 {filename} ({file_type})", flush=True)

    # Дубль?
    if registry_is_done(filename):
        bot.reply_to(msg, f"⚠️ «{filename}» уже обработан ранее.")
        return

    set_reaction(msg, "⏳")
    registry_add(filename, file_id, file_type, msg.chat.id)
    row_num = registry_find_row(filename)

    try:
        data = process_file(file_id, file_type, filename)
        rows = invoice_to_rows(data, filename)
        if not rows:
            raise ValueError("Позиции не найдены")
        sheet_append(INVOICES_SHEET, rows)
        registry_update(row_num, "✅")
        set_reaction(msg, "✅")
        supplier = data.get("supplier", "—")
        inv_num = data.get("invoice_number", "—")
        inv_date = data.get("invoice_date", "—")
        bot.reply_to(
            msg,
            f"✅ Добавлено {len(rows)} позиций\n"
            f"🏢 {supplier}\n"
            f"📄 Счёт №{inv_num} от {inv_date}\n"
            f"📊 https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}",
        )
    except Exception as e:
        err = str(e)[:200]
        print(f"❌ {filename}: {err}", flush=True)
        registry_update(row_num, "❌", err)
        set_reaction(msg, "❌")
        bot.reply_to(msg, f"❌ Ошибка: {err}\n⏳ Повтор через 20 мин автоматически.")


# ── Retry scheduler ────────────────────────────────────────────────────────────

def retry_failed_invoices():
    print("🔄 Проверяю неудачные счета...", flush=True)
    failed = registry_get_failed()
    if not failed:
        return

    results = {"ok": [], "fail": []}
    chat_ids = set()

    for item in failed:
        chat_ids.add(item["chat_id"])
        try:
            data = process_file(item["file_id"], item["file_type"], item["filename"])
            rows = invoice_to_rows(data, item["filename"])
            if not rows:
                raise ValueError("Позиции не найдены")
            sheet_append(INVOICES_SHEET, rows)
            registry_update(item["row"], "✅")
            results["ok"].append(item["filename"])
        except Exception as e:
            registry_update(item["row"], "❌", str(e)[:200])
            results["fail"].append(item["filename"])

    for chat_id in chat_ids:
        lines = []
        if results["ok"]:
            lines.append("✅ Успешно обработаны:")
            lines += [f"  • {f}" for f in results["ok"]]
        if results["fail"]:
            lines.append("❌ Снова ошибка:")
            lines += [f"  • {f}" for f in results["fail"]]
            lines.append("⏳ Повтор через 20 мин.")
        if lines:
            try:
                bot.send_message(chat_id, "\n".join(lines))
            except Exception:
                pass


# ── Telegram handlers ──────────────────────────────────────────────────────────

@bot.message_handler(content_types=["document"])
def on_document(msg):
    doc = msg.document
    name = (doc.file_name or "").lower()
    if name.endswith(".pdf"):
        ftype = "pdf"
    elif name.endswith((".xlsx", ".xls")):
        ftype = "excel"
    else:
        return
    process_invoice(msg, doc.file_id, ftype, doc.file_name or "unknown.pdf")


@bot.message_handler(content_types=["photo"])
def on_photo(msg):
    photo = msg.photo[-1]
    filename = f"photo_{datetime.now().strftime('%Y%m%d_%H%M%S')}.jpg"
    process_invoice(msg, photo.file_id, "photo", filename)


# ── Flask ──────────────────────────────────────────────────────────────────────

@app.route(f"/{TELEGRAM_TOKEN}", methods=["POST"])
def webhook():
    json_data = flask_request.get_json()
    msg = json_data.get("message", {}) if json_data else {}
    print(f"📨 keys: {list(msg.keys())}", flush=True)
    update = telebot.types.Update.de_json(json_data)
    bot.process_new_updates([update])
    return "ok", 200


@app.route("/retry")
def manual_retry():
    retry_failed_invoices()
    return "Retry triggered", 200


@app.route("/")
def home():
    return "Invoice bot is running", 200


# ── Startup ────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    try:
        ensure_sheets()
    except Exception as e:
        print(f"⚠️ Sheets init error: {e}", flush=True)

    bot.remove_webhook()
    bot.set_webhook(url=f"{WEBHOOK_URL}/{TELEGRAM_TOKEN}")
    print(f"Webhook: {WEBHOOK_URL}/{TELEGRAM_TOKEN}", flush=True)

    scheduler = BackgroundScheduler()
    scheduler.add_job(retry_failed_invoices, "interval", minutes=20)
    scheduler.start()

    print("Invoice bot запущен", flush=True)
    app.run(host="0.0.0.0", port=PORT)
