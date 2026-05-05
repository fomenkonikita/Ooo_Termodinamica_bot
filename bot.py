import os
import io
import json
import re
import base64
import gc
import time
import queue as _queue_module
from datetime import datetime
from threading import Thread, Lock

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

INVOICES_HEADERS = ["№", "Имя файла", "Поставщик", "Номер счёта", "Дата счёта", "Позиция в счёте",
                    "Наименование", "Артикул/Описание", "Ед.изм.", "Кол-во",
                    "Цена с НДС", "Сумма с НДС", "Дата добавления", "Общая сумма с НДС в счете",
                    "Примечание", "Имя отправителя", "Оплата"]

REGISTRY_HEADERS = ["#", "Имя файла", "Статус", "Получен", "Обработан", "Ошибка",
                    "file_id", "file_type", "chat_id", "Примечание", "Имя отправителя", "Оплата"]

bot = telebot.TeleBot(TELEGRAM_TOKEN)
client = Groq(api_key=GROQ_API_KEY)
app = Flask(__name__)

_invoice_queue = _queue_module.Queue()
_enqueue_lock = Lock()


def _sender_name(user) -> str:
    if user is None:
        return ""
    parts = [user.first_name or "", user.last_name or ""]
    full = " ".join(p for p in parts if p).strip()
    if user.username:
        return f"{full} (@{user.username})" if full else f"@{user.username}"
    return full


# ── Google Sheets ──────────────────────────────────────────────────────────────

_sheets_service_obj = None
_sheets_lock = Lock()

def get_sheets_service():
    global _sheets_service_obj
    with _sheets_lock:
        if _sheets_service_obj is None:
            creds = Credentials(
                token=None,
                refresh_token=GOOGLE_SHEETS_REFRESH_TOKEN,
                client_id=GOOGLE_CLIENT_ID,
                client_secret=GOOGLE_CLIENT_SECRET,
                token_uri="https://oauth2.googleapis.com/token",
            )
            creds.refresh(Request())
            _sheets_service_obj = build("sheets", "v4", credentials=creds)
        return _sheets_service_obj


_STATUSES = {"⏳", "⚙️", "✅", "❌"}


def _migrate_registry(service):
    """Приводит старые строки реестра (без # в колонке A) к новому формату."""
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=f"{REGISTRY_SHEET}!A:Z"
    ).execute()
    rows = result.get("values", [])
    updates = []
    for i, row in enumerate(rows[1:], start=2):
        if len(row) >= 2 and row[1] in _STATUSES:
            updates.append({
                "range": f"{REGISTRY_SHEET}!A{i}",
                "values": [[""] + row],
            })
    if updates:
        service.spreadsheets().values().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"valueInputOption": "RAW", "data": updates},
        ).execute()
        print(f"🔧 Мигрировано {len(updates)} строк реестра", flush=True)


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

    # Always update headers (handles column order fixes)
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{INVOICES_SHEET}!A1",
        valueInputOption="RAW",
        body={"values": [INVOICES_HEADERS]},
    ).execute()
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{REGISTRY_SHEET}!A1",
        valueInputOption="RAW",
        body={"values": [REGISTRY_HEADERS]},
    ).execute()

    _migrate_registry(service)

    # Подхватываем незавершённые файлы после рестарта
    rows = sheet_get_all(REGISTRY_SHEET)
    recovered = 0
    for i, row in enumerate(rows[1:], start=2):
        row = row + [''] * max(0, 13 - len(row))
        if row[2] not in ("⏳", "⚙️"):
            continue
        item = {
            "filename":    row[1],
            "file_id":     row[6],
            "file_type":   row[7],
            "chat_id":     int(row[8]) if row[8] else 0,
            "message_id":  _row_message_id(row),
            "caption":     row[9],
            "sender_name": row[10],
            "row_num":     i,
        }
        _invoice_queue.put(item)
        recovered += 1
    if recovered:
        print(f"♻️ Восстановлено {recovered} файлов из реестра", flush=True)

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
        if len(row) > 1 and row[1] == filename:
            return i
    return None


def registry_add(filename: str, file_id: str, file_type: str, chat_id: int,
                 message_id: int, caption: str = "", sender_name: str = ""):
    rows = sheet_get_all(REGISTRY_SHEET)
    seq_num = len(rows)   # header — строка 1; первая запись → #1, вторая → #2 ...
    now = datetime.now().strftime("%d.%m.%Y %H:%M")
    sheet_append(REGISTRY_SHEET, [[
        seq_num, filename, "⏳", now, "", "", file_id, file_type,
        str(chat_id), caption, sender_name, "", str(message_id)
    ]])


def registry_update(row_num: int, status: str, error: str = ""):
    now = datetime.now().strftime("%d.%m.%Y %H:%M") if status == "✅" else ""
    sheet_update_cell(REGISTRY_SHEET, row_num, 3, status)   # C = Статус
    sheet_update_cell(REGISTRY_SHEET, row_num, 5, now)      # E = Обработан
    sheet_update_cell(REGISTRY_SHEET, row_num, 6, error)    # F = Ошибка


# ── File parsing ───────────────────────────────────────────────────────────────

def extract_text_from_pdf(data: bytes) -> tuple[str, bool]:
    """Один проход по PDF. Возвращает (text, is_scanned).
    is_scanned=True если текста меньше 50 символов."""
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        header_lines = []
        supplier_line = ""
        table_rows = []
        total_lines = []
        total_chars = 0
        for page in pdf.pages:
            full_text = page.extract_text() or ""
            total_chars += len(full_text)
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
    if total_chars < 50:
        return "", True
    parts = []
    if supplier_line:
        parts.append(supplier_line)
    if header_lines:
        parts.append("\n".join(header_lines))
    parts.append("\n".join(table_rows)[:2500])
    if total_lines:
        parts.append("ИТОГО: " + total_lines[-1])
    return "\n---\n".join(parts), False


def compress_for_vision(image_bytes: bytes, max_px: int = 1024) -> bytes:
    """Сжимает изображение перед Vision API. PDF возвращает как есть."""
    try:
        from PIL import Image
        img = Image.open(io.BytesIO(image_bytes))
        if img.mode in ('RGBA', 'P', 'LA'):
            img = img.convert('RGB')
        if max(img.size) > max_px:
            img.thumbnail((max_px, max_px), Image.LANCZOS)
        buf = io.BytesIO()
        img.save(buf, format='JPEG', quality=85, optimize=True)
        compressed = buf.getvalue()
        print(f"🗜 {len(image_bytes)//1024}KB → {len(compressed)//1024}KB", flush=True)
        return compressed
    except Exception:
        return image_bytes  # PDF или нечитаемый формат — без изменений


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

def invoice_to_rows(data: dict, filename: str, caption: str = "", sender_name: str = "") -> list:
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
            today, total_amount, caption, sender_name, "",
        ])

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


def set_reaction(chat_id: int, message_id: int, emoji: str):
    try:
        bot.set_message_reaction(
            chat_id, message_id,
            [telebot.types.ReactionTypeEmoji(emoji)],
        )
    except Exception as e:
        print(f"⚠️ Reaction error: {e}", flush=True)


def process_file(file_id: str, file_type: str, filename: str) -> dict:
    file_bytes = download_file(file_id)
    try:
        if file_type == "pdf":
            text, is_scanned = extract_text_from_pdf(file_bytes)
            if is_scanned:
                print(f"🖼 Скан — использую Vision", flush=True)
                compressed = compress_for_vision(file_bytes)
                del file_bytes; gc.collect()
                return parse_image_with_ai(compressed)
            print(f"📄 Текст PDF: {len(text)} символов", flush=True)
            if is_garbled_text(text):
                print(f"🖼 Кириллица нечитаема — использую Vision", flush=True)
                del text
                compressed = compress_for_vision(file_bytes)
                del file_bytes; gc.collect()
                return parse_image_with_ai(compressed)
            data = parse_with_ai(text)
            del text
            supplier = (data.get("supplier") or "").strip().lower()
            if not supplier or supplier in PROMPT_PLACEHOLDERS:
                print(f"🖼 Поставщик не распознан — пробую Vision", flush=True)
                compressed = compress_for_vision(file_bytes)
                del file_bytes; gc.collect()
                return parse_image_with_ai(compressed)
            return data
        elif file_type == "excel":
            text = extract_text_from_excel(file_bytes)
            del file_bytes
            print(f"📊 Текст Excel: {len(text)} символов", flush=True)
            return parse_with_ai(text)
        elif file_type == "photo":
            compressed = compress_for_vision(file_bytes)
            del file_bytes; gc.collect()
            return parse_image_with_ai(compressed)
        else:
            raise ValueError(f"Неизвестный тип: {file_type}")
    finally:
        gc.collect()


# ── Queue worker ───────────────────────────────────────────────────────────────

def _process_item(item: dict):
    filename   = item["filename"]
    file_id    = item["file_id"]
    file_type  = item["file_type"]
    chat_id    = item["chat_id"]
    message_id = item["message_id"]
    caption     = item["caption"]
    sender_name = item.get("sender_name", "")
    row_num     = item["row_num"]

    def _react(emoji):
        if message_id:
            set_reaction(chat_id, message_id, emoji)

    print(f"📥 {filename} ({file_type}) msg_id={message_id}", flush=True)
    try:
        data = process_file(file_id, file_type, filename)
        rows = invoice_to_rows(data, filename, caption, sender_name)
        if not rows:
            raise ValueError("Позиции не найдены")
        sheet_append(INVOICES_SHEET, rows)
        if row_num:
            registry_update(row_num, "✅")
        _react("🏆")
        try:
            supplier = (data.get("supplier") or "—").strip()
            total = data.get("total_amount", 0) or 0
            try:
                total_str = f"{float(total):,.0f}".replace(",", " ")
            except Exception:
                total_str = str(total)
            text = (f"🏆 «{filename}» добавлен в таблицу\n"
                    f"Поставщик: {supplier}\n"
                    f"Сумма: {total_str} ₽")
            markup = telebot.types.InlineKeyboardMarkup()
            markup.add(telebot.types.InlineKeyboardButton(
                "💰 Отметить оплаченным", callback_data=f"pay:{message_id}"
            ))
            bot.send_message(chat_id, text, reply_markup=markup)
        except Exception as ex:
            print(f"⚠️ send pay button error: {ex}", flush=True)
    except Exception as e:
        err = str(e)
        is_limit = "429" in err or "413" in err
        print(f"❌ {filename}: {err[:200]}", flush=True)
        if row_num:
            registry_update(row_num, "❌", err[:200])
        if is_limit:
            _react("😴")
            try:
                bot.send_message(chat_id, f"😴 Лимит запросов — «{filename}» повторю через 20 мин.")
            except Exception:
                pass
        else:
            _react("🤬")
            try:
                bot.send_message(chat_id, f"🤬 Не удалось обработать «{filename}»\nОшибка: {err[:150]}")
            except Exception:
                pass


def _queue_worker():
    while True:
        item = _invoice_queue.get()
        try:
            _process_item(item)
        finally:
            _invoice_queue.task_done()
            gc.collect()
            time.sleep(3)


# ── Retry ──────────────────────────────────────────────────────────────────────

def retry_failed_invoices():
    """Сбрасывает ❌ → ⏳ и кладёт файлы обратно в очередь."""
    rows = sheet_get_all(REGISTRY_SHEET)
    count = 0
    for i, row in enumerate(rows[1:], start=2):
        row = row + [''] * max(0, 13 - len(row))
        if row[2] != "❌":
            continue
        sheet_update_cell(REGISTRY_SHEET, i, 3, "⏳")
        item = {
            "filename":    row[1],
            "file_id":     row[6],
            "file_type":   row[7],
            "chat_id":     int(row[8]) if row[8] else 0,
            "message_id":  _row_message_id(row),
            "caption":     row[9],
            "sender_name": row[10],
            "row_num":     i,
        }
        _invoice_queue.put(item)
        count += 1
    if count:
        print(f"🔄 Сброшено {count} файлов → очередь", flush=True)
    return count


# ── Paid status ────────────────────────────────────────────────────────────────

def invoices_find_rows(filename: str) -> list:
    rows = sheet_get_all(INVOICES_SHEET)
    return [i for i, row in enumerate(rows[1:], start=2)
            if len(row) > 1 and row[1] == filename]


def _row_message_id(row: list):
    """Возвращает message_id из строки реестра.
    Новый формат: индекс 12 (M). Старый формат: индекс 10 (K)."""
    row = row + [''] * max(0, 13 - len(row))
    for idx in (12, 10):
        val = row[idx]
        if val:
            try:
                return int(val)
            except (ValueError, TypeError):
                pass
    return None


def mark_paid(message_id: int, chat_id: int, paid: bool):
    try:
        value = "✅" if paid else ""
        rows = sheet_get_all(REGISTRY_SHEET)
        for i, row in enumerate(rows[1:], start=2):
            if _row_message_id(row) != message_id:
                continue
            filename = row[1]
            sheet_update_cell(REGISTRY_SHEET, i, 12, value)   # L = Оплата
            for j in invoices_find_rows(filename):
                sheet_update_cell(INVOICES_SHEET, j, 17, value)   # Q = Оплата
            print(f"💰 {filename} {'оплачен ✅' if paid else 'снята отметка оплаты'}", flush=True)
            return
        print(f"⚠️ mark_paid: message_id={message_id} не найден в реестре", flush=True)
    except Exception as e:
        print(f"❌ mark_paid error: {e}", flush=True)


@bot.callback_query_handler(func=lambda c: c.data.startswith(("pay:", "unpay:")))
def on_pay_callback(call):
    try:
        action, msg_id_str = call.data.split(":", 1)
        msg_id = int(msg_id_str)
        paid = action == "pay"
        Thread(target=mark_paid, args=(msg_id, call.message.chat.id, paid), daemon=True).start()
        new_action = "unpay" if paid else "pay"
        new_label  = "✅ Оплачено" if paid else "💰 Отметить оплаченным"
        markup = telebot.types.InlineKeyboardMarkup()
        markup.add(telebot.types.InlineKeyboardButton(new_label, callback_data=f"{new_action}:{msg_id}"))
        bot.edit_message_reply_markup(call.message.chat.id, call.message.id, reply_markup=markup)
        bot.answer_callback_query(call.id)
    except Exception as e:
        print(f"❌ pay callback error: {e}", flush=True)
        try:
            bot.answer_callback_query(call.id, "Ошибка")
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
                "Как присылать счета:\n"
                "📎 Отправляйте по одному файлу\n"
                "✏️ Добавляйте подпись к каждому файлу (например: «срочно», «Сандуны», «АБК») — она попадёт в столбец Примечание и поможет найти счёт в таблице\n\n"
                "Реакции бота:\n"
                "👀 — зарегистрировал, скоро обработаю\n"
                "🏆 — добавлено в таблицу\n"
                "😴 — лимит запросов, повторю через 20 мин\n"
                "🤬 — ошибка, загляни в лог\n\n"
                f"📊 Таблица: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}"
            )
            break


@bot.message_handler(commands=["start", "help"])
def on_start(msg):
    bot.reply_to(msg,
        "Отправляй счета по одному (PDF, Excel или фото).\n"
        "Добавляй подпись к файлу — она сохранится в столбец Примечание.\n\n"
        "Реакции:\n"
        "👀 — зарегистрировал\n"
        "🏆 — добавлено в таблицу\n"
        "😴 — лимит, повторю через 20 мин\n"
        "🤬 — ошибка, загляни в лог\n\n"
        "/retry — повторить обработку счетов с ошибками\n"
        f"📊 Таблица: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}"
    )


@bot.message_handler(commands=["retry"])
def on_retry(msg):
    rows = sheet_get_all(REGISTRY_SHEET)
    count = sum(1 for r in rows[1:] if len(r) > 2 and r[2] == "❌")
    if count == 0:
        bot.reply_to(msg, "Нет счетов с ошибками.")
        return
    retry_failed_invoices()
    bot.reply_to(msg, f"↩️ Запускаю повтор для {count} счет(ов)... Результат в течение минуты.")


def enqueue_invoice(msg, file_id: str, file_type: str, filename: str):
    """Регистрирует файл в Sheets и кладёт в очередь на обработку."""
    try:
        with _enqueue_lock:
            rows = sheet_get_all(REGISTRY_SHEET)
            for row in rows[1:]:
                row_filename = row[1] if len(row) > 1 else row[0]
                row_status   = row[2] if len(row) > 2 else (row[1] if len(row) > 1 else "")
                if row_filename == filename and row_status in ("✅", "⏳", "⚙️"):
                    if row_status == "✅":
                        bot.reply_to(msg, f"⚠️ «{filename}» уже обработан ранее.")
                    else:
                        bot.reply_to(msg, f"⚠️ «{filename}» уже в очереди на обработку.")
                    return
            caption     = msg.caption or ""
            sender_name = _sender_name(msg.from_user)
            print(f"📎 {filename} caption={caption!r} from={sender_name!r}", flush=True)
            registry_add(filename, file_id, file_type, msg.chat.id, msg.message_id, caption, sender_name)
            row_num = registry_find_row(filename)
            item = {
                "filename":    filename,
                "file_id":     file_id,
                "file_type":   file_type,
                "chat_id":     msg.chat.id,
                "message_id":  msg.message_id,
                "caption":     caption,
                "sender_name": sender_name,
                "row_num":     row_num,
            }
            _invoice_queue.put(item)
        set_reaction(msg.chat.id, msg.message_id, "👀")
        print(f"📋 Зарегистрирован: {filename} row={row_num}", flush=True)
    except Exception as e:
        print(f"❌ enqueue {filename}: {e}", flush=True)
        try:
            set_reaction(msg.chat.id, msg.message_id, "🤬")
            bot.send_message(msg.chat.id, f"🤬 Не удалось принять «{filename}». Отправь ещё раз.")
        except Exception:
            pass


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
    update_keys = list(json_data.keys()) if json_data else []
    print(f"📨 update: {update_keys}", flush=True)

    update = telebot.types.Update.de_json(json_data)
    Thread(target=bot.process_new_updates, args=([update],), daemon=True).start()
    return "ok", 200


@app.route("/retry")
def manual_retry():
    count = retry_failed_invoices()
    return f"Retry triggered: {count} файлов → ⏳", 200


@app.route("/debug/webhook")
def debug_webhook():
    info = bot.get_webhook_info()
    return {
        "url": info.url[-30:] if info.url else None,
        "allowed_updates": info.allowed_updates,
        "pending_update_count": info.pending_update_count,
        "last_error_message": info.last_error_message,
        "last_error_date": info.last_error_date,
    }, 200


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
    bot.set_webhook(
        url=f"{WEBHOOK_URL}/{TELEGRAM_TOKEN}",
        allowed_updates=["message", "callback_query"],
    )
    print(f"Webhook: {WEBHOOK_URL}/{TELEGRAM_TOKEN}", flush=True)

    bot.set_my_commands([
        telebot.types.BotCommand("/start", "Информация и ссылка на таблицу"),
        telebot.types.BotCommand("/retry", "Повторить обработку счетов с ошибками"),
    ])

    Thread(target=_queue_worker, daemon=True).start()
    print("📋 Queue worker запущен", flush=True)

    scheduler = BackgroundScheduler()
    scheduler.add_job(retry_failed_invoices, "interval", minutes=20)
    scheduler.start()

    print("Invoice bot запущен", flush=True)
    app.run(host="0.0.0.0", port=PORT)
