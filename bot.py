import os
import io
import json
import re
import base64
import tempfile
import time
from datetime import datetime

import requests
import telebot
from groq import Groq
from flask import Flask, request as flask_request
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

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

SHEET_HEADERS = ["№", "Поставщик", "Номер счёта", "Дата счёта", "Позиция в счете", "Наименование", "Артикул/Описание", "Ед.изм.", "Кол-во", "Цена с НДС", "Сумма с НДС", "Дата добавления", "Общая сумма с НДС в счете"]

bot = telebot.TeleBot(TELEGRAM_TOKEN)
client = Groq(api_key=GROQ_API_KEY)
app = Flask(__name__)

_spreadsheet_id = SPREADSHEET_ID


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


def get_or_create_spreadsheet():
    global _spreadsheet_id
    if _spreadsheet_id:
        return _spreadsheet_id

    service = get_sheets_service()
    spreadsheet = service.spreadsheets().create(body={
        "properties": {"title": "База поставщиков — счета"},
        "sheets": [{"properties": {"title": "Счета"}}],
    }).execute()
    _spreadsheet_id = spreadsheet["spreadsheetId"]

    service.spreadsheets().values().update(
        spreadsheetId=_spreadsheet_id,
        range="Счета!A1",
        valueInputOption="RAW",
        body={"values": [SHEET_HEADERS]},
    ).execute()

    url = f"https://docs.google.com/spreadsheets/d/{_spreadsheet_id}"
    print(f"\n✅ Создана новая таблица!\nSPREADSHEET_ID={_spreadsheet_id}\n{url}\n")
    print("⚠️  Добавь SPREADSHEET_ID в переменные Render чтобы таблица не пересоздавалась!")
    return _spreadsheet_id


def write_rows_to_sheet(rows):
    service = get_sheets_service()
    sid = get_or_create_spreadsheet()
    service.spreadsheets().values().append(
        spreadsheetId=sid,
        range="Счета!A1",
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body={"values": rows},
    ).execute()


def extract_text_from_pdf(data: bytes) -> str:
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        pages = []
        for page in pdf.pages:
            text = page.extract_text() or ""
            # Also try table extraction for tabular invoices
            tables = page.extract_tables()
            if tables:
                for table in tables:
                    for row in table:
                        if row:
                            text += "\n" + "\t".join(str(c) if c else "" for c in row)
            pages.append(text)
    return "\n\n--- страница ---\n\n".join(pages)


def extract_text_from_excel(data: bytes) -> str:
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    lines = []
    for sheet in wb.worksheets:
        lines.append(f"=== Лист: {sheet.title} ===")
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
    {"pos": 1, "name": "Наименование товара/услуги", "article": "артикул/тип/производитель или пусто", "unit": "ед.изм.", "quantity": 1, "price_with_vat": 0.0}
  ]
}

Правила:
- supplier: название компании-поставщика (кто выставил счёт)
- invoice_number: номер счёта/накладной
- invoice_date: дата выставления счёта
- total_amount: итоговая сумма всего счёта с НДС
- items: все позиции товаров/услуг из счёта
- pos: номер позиции в счёте (1, 2, 3...)
- name: наименование товара или услуги
- article: артикул, тип, марка, производитель или любое доп. описание товара (если есть)
- unit: шт, м, м2, м3, кг, л, компл, усл и т.п.
- price_with_vat: цена за единицу С учётом НДС (если в счёте цена без НДС — прибавь НДС 20%)
- quantity: количество (число)
- Если поле неизвестно — пустая строка или 0"""


def extract_json(text: str) -> dict:
    # Find outermost { } block
    start = text.find("{")
    if start == -1:
        raise ValueError("JSON не найден в ответе AI")
    depth = 0
    for i, ch in enumerate(text[start:], start):
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                return json.loads(text[start:i+1])
    raise ValueError("Незакрытый JSON")


def parse_invoice_from_text(text: str) -> dict:
    for attempt in range(3):
        try:
            resp = client.chat.completions.create(
                model="llama-3.1-8b-instant",
                messages=[
                    {"role": "system", "content": EXTRACT_PROMPT},
                    {"role": "user", "content": f"Счёт:\n\n{text[:8000]}"},
                ],
                max_tokens=2000,
                temperature=0.1,
            )
            raw = resp.choices[0].message.content.strip()
            print(f"🤖 AI ответ: {raw[:500]}", flush=True)
            return extract_json(raw)
        except Exception as e:
            if "429" in str(e) and attempt < 2:
                wait = 30 * (attempt + 1)
                print(f"⏳ Rate limit, жду {wait}с (попытка {attempt+1}/3)", flush=True)
                time.sleep(wait)
            else:
                raise


def parse_invoice_from_image(image_bytes: bytes, mime: str = "image/jpeg") -> dict:
    b64 = base64.b64encode(image_bytes).decode()
    resp = client.chat.completions.create(
        model="meta-llama/llama-4-scout-17b-16e-instruct",
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": EXTRACT_PROMPT + "\n\nИзвлеки данные из фотографии счёта:"},
                    {"type": "image_url", "image_url": {"url": f"data:{mime};base64,{b64}"}},
                ],
            }
        ],
        max_tokens=2000,
        temperature=0.1,
    )
    raw = resp.choices[0].message.content.strip()
    return extract_json(raw)


def invoice_to_rows(data: dict, start_num: int = 1) -> list:
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
            start_num + i - 1,
            supplier,
            inv_num,
            inv_date,
            item.get("pos", i),
            item.get("name", "—"),
            item.get("article", ""),
            item.get("unit", "—"),
            qty,
            price,
            line_total,
            today,
            total_amount,
        ])
    return rows


def download_file(file_id: str) -> bytes:
    file_info = bot.get_file(file_id)
    url = f"https://api.telegram.org/file/bot{TELEGRAM_TOKEN}/{file_info.file_path}"
    return requests.get(url).content


def set_reaction(msg, emoji: str):
    try:
        bot.set_message_reaction(
            msg.chat.id,
            msg.message_id,
            [telebot.types.ReactionTypeEmoji(emoji)],
        )
    except Exception as e:
        print(f"⚠️ Reaction error ({emoji}): {e}")


def process_invoice(msg, file_bytes: bytes, file_type: str, filename: str = ""):
    print(f"📥 Получен файл: {file_type} {filename} ({len(file_bytes)} bytes)")
    try:
        set_reaction(msg, "⏳")

        if file_type == "pdf":
            text = extract_text_from_pdf(file_bytes)
            print(f"📄 Текст извлечён: {len(text)} символов")
            data = parse_invoice_from_text(text)
        elif file_type == "excel":
            text = extract_text_from_excel(file_bytes)
            print(f"📊 Excel прочитан: {len(text)} символов")
            data = parse_invoice_from_text(text)
        elif file_type == "photo":
            data = parse_invoice_from_image(file_bytes)
        else:
            set_reaction(msg, "❌")
            return

        print(f"🤖 AI вернул: {data}")
        rows = invoice_to_rows(data)
        print(f"📋 Строк для записи: {len(rows)}", flush=True)
        if not rows:
            set_reaction(msg, "❌")
            bot.reply_to(msg, "⚠️ Позиции не найдены в счёте.")
            return

        write_rows_to_sheet(rows)
        sid = get_or_create_spreadsheet()
        set_reaction(msg, "✅")
        supplier = data.get("supplier", "—")
        inv_num = data.get("invoice_number", "—")
        inv_date = data.get("invoice_date", "—")
        bot.reply_to(
            msg,
            f"✅ Добавлено {len(rows)} позиций\n"
            f"🏢 {supplier}\n"
            f"📄 Счёт №{inv_num} от {inv_date}\n"
            f"📊 https://docs.google.com/spreadsheets/d/{sid}",
        )
        print("✅ Записано в таблицу", flush=True)
    except Exception as e:
        print(f"❌ Ошибка в process_invoice: {e}", flush=True)
        set_reaction(msg, "❌")
        bot.reply_to(msg, f"❌ Ошибка: {e}")


@bot.message_handler(content_types=["document"])
def on_document(msg):
    doc = msg.document
    name = (doc.file_name or "").lower()
    if name.endswith(".pdf"):
        ftype = "pdf"
    elif name.endswith((".xlsx", ".xls")):
        ftype = "excel"
    else:
        return  # Игнорируем другие файлы
    data = download_file(doc.file_id)
    process_invoice(msg, data, ftype, doc.file_name)


@bot.message_handler(content_types=["photo"])
def on_photo(msg):
    # Берём фото наибольшего размера
    photo = msg.photo[-1]
    data = download_file(photo.file_id)
    process_invoice(msg, data, "photo")


@app.route(f"/{TELEGRAM_TOKEN}", methods=["POST"])
def webhook():
    json_data = flask_request.get_json()
    print(f"📨 Update type: {list(json_data.keys()) if json_data else None}", flush=True)
    update = telebot.types.Update.de_json(json_data)
    bot.process_new_updates([update])
    return "ok", 200


@app.route("/")
def home():
    return "Invoice bot is running", 200


if __name__ == "__main__":
    # Инициализируем таблицу при старте если SPREADSHEET_ID не задан
    try:
        get_or_create_spreadsheet()
    except Exception as e:
        print(f"⚠️ Sheets init error: {e}")

    bot.remove_webhook()
    bot.set_webhook(url=f"{WEBHOOK_URL}/{TELEGRAM_TOKEN}")
    print(f"Webhook: {WEBHOOK_URL}/{TELEGRAM_TOKEN}")
    print("Invoice bot запущен")
    app.run(host="0.0.0.0", port=PORT)
