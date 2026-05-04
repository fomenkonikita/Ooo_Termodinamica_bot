import os
import io
import json
import re
import base64
import time
from datetime import datetime
from queue import Queue
from threading import Thread

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
import xlrd

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

_invoice_queue: Queue = Queue()


def _queue_worker():
    while True:
        item = _invoice_queue.get()
        try:
            msg, file_id, file_type, filename = item
            process_invoice(msg, file_id, file_type, filename)
        except Exception as e:
            print(f"⚠️ Queue worker error: {e}", flush=True)
        finally:
            _invoice_queue.task_done()


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
    now = datetime.now()
    for i, row in enumerate(rows[1:], start=2):
        if len(row) < 8:
            continue
        status = row[1]
        if status == "❌":
            failed.append({"row": i, "filename": row[0], "file_id": row[5],
                           "file_type": row[6], "chat_id": int(row[7])})
        elif status == "⏳":
            # Ретраим только если завис более 30 минут (защита от гонки с планировщиком)
            try:
                received = datetime.strptime(row[2], "%d.%m.%Y %H:%M")
                if (now - received).total_seconds() > 1800:
                    failed.append({"row": i, "filename": row[0], "file_id": row[5],
                                   "file_type": row[6], "chat_id": int(row[7])})
            except Exception:
                pass
    return failed


# ── File parsing ───────────────────────────────────────────────────────────────

def extract_text_from_pdf(data: bytes) -> str:
    """Извлекает текст PDF: заголовок счёта + строки таблиц + строка итого."""
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        header_lines = []
        supplier_line = ""
        table_rows = []
        total_lines = []
        for page in pdf.pages:
            full_text = page.extract_text() or ""
            for line in full_text.splitlines():
                if not supplier_line and re.search(r'^поставщик:', line, re.IGNORECASE):
                    supplier_line = line.strip()
                if not header_lines and re.search(r'счет[аё]?\s*(на\s*оплату)?\s*№', line, re.IGNORECASE):
                    header_lines.append(line.strip())
                if re.search(r'итого|всего к оплате|к оплате', line, re.IGNORECASE):
                    total_lines.append(line.strip())
            tables = page.extract_tables()
            if tables:
                for table in tables:
                    for row in table:
                        if row and any(c for c in row if c and str(c).strip()):
                            table_rows.append("\t".join(str(c).strip() if c else "" for c in row))
            elif not header_lines:
                table_rows.append(full_text[:800])
        parts = []
        if supplier_line:
            parts.append(supplier_line)
        if header_lines:
            parts.append("\n".join(header_lines))
        parts.append("\n".join(table_rows)[:2500])
        if total_lines:
            parts.append("ИТОГО: " + total_lines[-1])
        return "\n---\n".join(parts)


def is_scanned_pdf(data: bytes) -> bool:
    """True если PDF не содержит текста (скан)."""
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        total = sum(len(p.extract_text() or "") for p in pdf.pages)
    return total < 50


def is_garbled_text(text: str) -> bool:
    """True если кириллица нечитаема — кодировка шрифта сломана."""
    alpha = [c for c in text if c.isalpha()]
    if len(alpha) < 20:
        return True
    readable = sum(1 for c in alpha if 'Ѐ' <= c <= 'ӿ' or c.isascii())
    return readable / len(alpha) < 0.5


def _excel_rows_to_text(all_rows: list[list[str]]) -> str:
    """Из списка строк вытаскивает заголовок счёта + строки таблицы."""
    header = ""
    table_lines = []
    for vals in all_rows:
        nonempty = [v for v in vals if v]
        # Заголовок счёта — одна ячейка с "счет.*№"
        if not header and len(nonempty) == 1:
            if re.search(r'счет[аё]?\s*(на\s*оплату)?\s*№', nonempty[0], re.IGNORECASE):
                header = nonempty[0]
                continue
        if len(nonempty) >= 2:
            table_lines.append("\t".join(vals))
    prefix = header + "\n---\n" if header else ""
    return (prefix + "\n".join(table_lines))[:3000]


def extract_text_from_excel(data: bytes) -> str:
    """Парсит .xlsx (openpyxl) и .xls (xlrd)."""
    try:
        wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
        all_rows = []
        for sheet in wb.worksheets:
            for row in sheet.iter_rows(values_only=True):
                vals = []
                for v in row:
                    try:
                        vals.append(str(v).strip() if v is not None else "")
                    except Exception:
                        vals.append("")
                all_rows.append(vals)
    except Exception:
        # Fallback: старый формат .xls
        book = xlrd.open_workbook(file_contents=data)
        all_rows = []
        for sheet in book.sheets():
            for row_idx in range(sheet.nrows):
                vals = []
                for v in sheet.row_values(row_idx):
                    try:
                        vals.append(str(v).strip() if v != "" else "")
                    except Exception:
                        vals.append("")
                all_rows.append(vals)
    return _excel_rows_to_text(all_rows)


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
- supplier: кто ВЫСТАВИЛ счёт — продавец/поставщик/исполнитель. Их название в шапке счёта рядом с ИНН/КПП поставщика. НЕ бери из строки "Покупатель:" — там всегда ООО "Термодинамика" или похожий покупатель
- invoice_number: КОРОТКИЙ номер из заголовка "Счёт № X" или "Счёт на оплату № X". Это 1-5 цифр/символов. НЕ бери 20-значный расчётный счёт (р/с) банка. Ищи строку вида "Счёт на оплату № 88 от..." — нужна цифра после "№"
- invoice_date: дата из той же строки заголовка "Счёт № X от ДД.ММ.ГГГГ" или "от 03 февраля 2026 г." — конвертируй в ДД.ММ.ГГГГ. Русские месяцы: января=01, февраля=02, марта=03, апреля=04, мая=05, июня=06, июля=07, августа=08, сентября=09, октября=10, ноября=11, декабря=12
- total_amount: итоговая сумма ВСЕГО счёта с НДС — ищи последнюю строку "Итого:" или "Всего к оплате:" в конце таблицы. Это одно конкретное число, НЕ суммируй сам
- pos: номер позиции в счёте (1, 2, 3...)
- article: артикул, тип, марка, производитель (если есть)
- price_with_vat: ЦЕНА ЗА ЕДИНИЦУ товара с НДС. Это НЕ сумма строки! Если есть колонки "Цена" и "Сумма" — бери из "Цена". Если цена без НДС — умножь на 1.2
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


PROMPT_PLACEHOLDERS = {"название поставщика", "номер счёта", "наименование", "артикул/тип или пусто"}

def clean_field(val: str, fallback: str = "—") -> str:
    if not val or val.strip().lower() in PROMPT_PLACEHOLDERS:
        return fallback
    return val.strip()

def invoice_to_rows(data: dict, filename: str) -> list:
    today = datetime.now().strftime("%d.%m.%Y")
    supplier = clean_field(data.get("supplier", ""))
    inv_num = clean_field(data.get("invoice_number", ""))
    inv_date = clean_field(data.get("invoice_date", ""))
    total_amount = data.get("total_amount", 0) or 0
    try:
        total_amount = float(total_amount)
    except Exception:
        total_amount = 0

    rows = []
    items_sum = 0.0
    for i, item in enumerate(data.get("items", []), start=1):
        qty = item.get("quantity", 0) or 0
        price = item.get("price_with_vat", 0) or 0
        try:
            qty = float(qty)
            price = float(price)
            line_total = round(qty * price, 2)
            if qty > 1 and price > 0 and line_total == price:
                price = round(price / qty, 2)
                line_total = round(qty * price, 2)
        except Exception:
            line_total = 0
        items_sum += line_total
        rows.append([
            i, filename, supplier, inv_num, inv_date,
            item.get("pos", i), item.get("name", "—"), item.get("article", ""),
            item.get("unit", "—"), qty, price, line_total,
            today, total_amount,
        ])

    # Если AI вернул 0 или явно заниженную сумму (меньше 60% от суммы позиций) — считаем сами
    if items_sum > 0 and (total_amount == 0 or total_amount < items_sum * 0.6):
        fixed = round(items_sum, 2)
        print(f"⚠️ total_amount {total_amount} → пересчитан из позиций: {fixed}", flush=True)
        total_amount = fixed
        for row in rows:
            row[13] = total_amount

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
        if is_scanned_pdf(file_bytes):
            print(f"🖼 Скан — использую Vision", flush=True)
            return parse_image_with_ai(file_bytes)
        text = extract_text_from_pdf(file_bytes)
        print(f"📄 Текст PDF: {len(text)} символов", flush=True)
        if is_garbled_text(text):
            print(f"🖼 Кириллица нечитаема — использую Vision", flush=True)
            return parse_image_with_ai(file_bytes)
        data = parse_with_ai(text)
        supplier = (data.get("supplier") or "").strip().lower()
        if not supplier or supplier in PROMPT_PLACEHOLDERS:
            print(f"🖼 Поставщик не распознан — пробую Vision", flush=True)
            return parse_image_with_ai(file_bytes)
        return data
    elif file_type == "excel":
        text = extract_text_from_excel(file_bytes)
        print(f"📊 Текст Excel: {len(text)} символов", flush=True)
        return parse_with_ai(text)
    elif file_type == "photo":
        return parse_image_with_ai(file_bytes)
    else:
        raise ValueError(f"Неизвестный тип: {file_type}")


def process_invoice(msg, file_id: str, file_type: str, filename: str):
    print(f"📥 {filename} ({file_type})", flush=True)
    row_num = registry_find_row(filename)
    if row_num is None:
        return  # не зарегистрирован — пропускаем

    try:
        data = process_file(file_id, file_type, filename)
        rows = invoice_to_rows(data, filename)
        if not rows:
            raise ValueError("Позиции не найдены")
        sheet_append(INVOICES_SHEET, rows)
        registry_update(row_num, "✅")
        set_reaction(msg, "🏆")
    except Exception as e:
        err = str(e)
        is_limit = "429" in err or "413" in err
        print(f"❌ {filename}: {err[:200]}", flush=True)
        registry_update(row_num, "❌", err[:200])
        if is_limit:
            set_reaction(msg, "😴")
            bot.reply_to(msg, f"😴 Лимит запросов — «{filename}» повторю через 20 мин автоматически.")
        else:
            set_reaction(msg, "🤬")
            bot.reply_to(msg, f"🤬 Не удалось обработать «{filename}»\nОшибка: {err[:150]}")


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

@bot.message_handler(content_types=["new_chat_members"])
def on_new_member(msg):
    for member in msg.new_chat_members:
        if member.id == bot.get_me().id:
            bot.send_message(
                msg.chat.id,
                "Привет! Я бот для учёта счетов от поставщиков 👋\n\n"
                "Что умею:\n"
                "📄 Принимаю счета в виде PDF, Excel (.xlsx/.xls) или фото\n"
                "🤖 Извлекаю данные с помощью AI: поставщик, номер, дата, позиции, цены\n"
                "📊 Автоматически записываю всё в Google Таблицу\n"
                "🔄 При ошибке — повторяю попытку каждые 20 минут\n\n"
                "Мои реакции:\n"
                "👀 — обрабатываю\n"
                "🏆 — добавлено в таблицу\n"
                "😴 — лимит, повторю через 20 мин\n"
                "🤬 — ошибка, загляни в лог\n\n"
                "Просто скиньте счёт в чат — остальное сделаю сам!\n\n"
                f"📊 Таблица: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}"
            )
            break


@bot.message_handler(commands=["start", "help"])
def on_start(msg):
    bot.reply_to(msg,
        "Привет! Отправь мне счёт (PDF, Excel или фото) — я внесу его в таблицу.\n\n"
        "/retry — повторить обработку счетов с ошибками\n"
        f"📊 Таблица: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}"
    )


@bot.message_handler(commands=["retry"])
def on_retry(msg):
    failed = registry_get_failed()
    if not failed:
        bot.reply_to(msg, "Нет счетов для повтора.")
        return
    bot.reply_to(msg, f"Запускаю повтор для {len(failed)} счет(ов)...")
    retry_failed_invoices()


def enqueue_invoice(msg, file_id: str, file_type: str, filename: str):
    """Сразу регистрирует файл и ставит 👀, потом кладёт в очередь обработки."""
    if registry_is_done(filename):
        bot.reply_to(msg, f"⚠️ «{filename}» уже обработан ранее.")
        return
    set_reaction(msg, "👀")
    registry_add(filename, file_id, file_type, msg.chat.id)
    print(f"📋 В очередь: {filename}", flush=True)
    _invoice_queue.put((msg, file_id, file_type, filename))


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
    enqueue_invoice(msg, doc.file_id, ftype, doc.file_name or "unknown.pdf")


@bot.message_handler(content_types=["photo"])
def on_photo(msg):
    photo = msg.photo[-1]
    filename = f"photo_{datetime.now().strftime('%Y%m%d_%H%M%S')}.jpg"
    enqueue_invoice(msg, photo.file_id, "photo", filename)


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

    bot.set_my_commands([
        telebot.types.BotCommand("/start", "Информация и ссылка на таблицу"),
        telebot.types.BotCommand("/retry", "Повторить обработку счетов с ошибками"),
    ])

    worker = Thread(target=_queue_worker, daemon=True)
    worker.start()
    print("📋 Queue worker запущен", flush=True)

    scheduler = BackgroundScheduler()
    scheduler.add_job(retry_failed_invoices, "interval", minutes=20)
    scheduler.start()

    print("Invoice bot запущен", flush=True)
    app.run(host="0.0.0.0", port=PORT)
