import os
import io
import json
import re
import base64
import gc
import time
import queue as _queue_module
from datetime import datetime, timedelta

def _now():
    return datetime.utcnow() + timedelta(hours=5)
from threading import Thread, Lock

import requests
import telebot
from groq import Groq
from flask import Flask, request as flask_request
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from apscheduler.schedulers.background import BackgroundScheduler

import pdfplumber
import openpyxl
import xlrd

TELEGRAM_TOKEN = os.environ["TELEGRAM_TOKEN"]
GROQ_API_KEY = os.environ["GROQ_API_KEY"]
GOOGLE_SERVICE_ACCOUNT_JSON = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]
WEBHOOK_URL = os.environ["WEBHOOK_URL"]
PORT = int(os.environ.get("PORT", 8080))
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "1tmuj1f2D2euUZlr-CXHzgkurRavF-MV8UxvjG3SIKDc")

INVOICES_SHEET = "Счета"
REGISTRY_SHEET = "Реестр"

INVOICES_HEADERS = ["№", "Имя файла", "Поставщик", "Номер счёта", "Дата счёта", "Позиция в счёте",
                    "Наименование", "Артикул/Описание", "Ед.изм.", "Кол-во",
                    "Цена с НДС", "Сумма с НДС", "Дата добавления", "Общая сумма с НДС в счете",
                    "Примечание", "Имя отправителя", "Оплата", "Файл"]

REGISTRY_HEADERS = ["#", "Поставщик", "Сумма счета", "Имя файла", "Статус", "Получен", "Обработан", "Ошибка",
                    "file_id", "file_type", "chat_id", "Примечание", "Имя отправителя", "Оплата", "message_id", "Ссылка"]

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
            info = json.loads(GOOGLE_SERVICE_ACCOUNT_JSON)
            creds = service_account.Credentials.from_service_account_info(
                info,
                scopes=["https://www.googleapis.com/auth/spreadsheets"],
            )
            _sheets_service_obj = build("sheets", "v4", credentials=creds)
        return _sheets_service_obj


def _reset_sheets_service():
    global _sheets_service_obj
    with _sheets_lock:
        _sheets_service_obj = None


# ── Google Drive ────────────────────────────────────────────────────────────────

_drive_service_obj = None
_drive_lock = Lock()
_drive_folder_id = None

def get_drive_service():
    global _drive_service_obj
    with _drive_lock:
        if _drive_service_obj is None:
            info = json.loads(GOOGLE_SERVICE_ACCOUNT_JSON)
            creds = service_account.Credentials.from_service_account_info(
                info,
                scopes=["https://www.googleapis.com/auth/drive"],
            )
            _drive_service_obj = build("drive", "v3", credentials=creds)
        return _drive_service_obj


def _get_drive_folder() -> str:
    global _drive_folder_id
    if _drive_folder_id:
        return _drive_folder_id
    svc = get_drive_service()
    res = svc.files().list(
        q="mimeType='application/vnd.google-apps.folder' and name='Счета_Термодинамика' and trashed=false",
        fields="files(id)",
    ).execute()
    files = res.get("files", [])
    if files:
        _drive_folder_id = files[0]["id"]
    else:
        folder = svc.files().create(
            body={"name": "Счета_Термодинамика", "mimeType": "application/vnd.google-apps.folder"},
            fields="id",
        ).execute()
        _drive_folder_id = folder["id"]
        svc.permissions().create(
            fileId=_drive_folder_id,
            body={"type": "anyone", "role": "reader"},
        ).execute()
        print(f"📁 Drive folder created: {_drive_folder_id}", flush=True)
    return _drive_folder_id


def upload_to_drive(filename: str, file_bytes: bytes, file_type: str) -> str:
    mime_map = {
        "pdf":   "application/pdf",
        "excel": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "photo": "image/jpeg",
    }
    mime = mime_map.get(file_type, "application/octet-stream")
    folder_id = _get_drive_folder()
    svc = get_drive_service()
    media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mime, resumable=False)
    uploaded = svc.files().create(
        body={"name": filename, "parents": [folder_id]},
        media_body=media,
        fields="id,webViewLink",
    ).execute()
    link = uploaded.get("webViewLink", "")
    print(f"☁️ Drive: {filename} → {link}", flush=True)
    return link


def _sheets_call(fn):
    """Выполняет fn(service). При SSL/сетевой ошибке сбрасывает соединение и повторяет."""
    try:
        return fn(get_sheets_service())
    except Exception as e:
        err = str(e).lower()
        if any(x in err for x in ("ssl", "wrong version", "connection", "timeout", "timed out")):
            print(f"⚠️ Sheets connection error, reconnecting: {e}", flush=True)
            _reset_sheets_service()
            return fn(get_sheets_service())
        raise


_STATUSES = {"⏳", "⚙️", "✅", "❌"}


def _migrate_registry(service):
    """Приводит старые строки реестра (без # в колонке A) к новому формату."""
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=f"{REGISTRY_SHEET}!A:Z"
    ).execute()
    rows = result.get("values", [])
    updates = []
    for i, row in enumerate(rows[1:], start=2):
        if not row:
            continue
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
        if not row:
            continue
        row = row + [''] * max(0, 15 - len(row))
        if row[4] not in ("⏳", "⚙️"):
            continue
        item = {
            "filename":    row[3],
            "file_id":     row[8],
            "file_type":   row[9],
            "chat_id":     int(row[10]) if row[10] else 0,
            "message_id":  _row_message_id(row),
            "caption":     row[11],
            "sender_name": row[12],
            "row_num":     i,
        }
        _invoice_queue.put(item)
        recovered += 1
    if recovered:
        print(f"♻️ Восстановлено {recovered} файлов из реестра", flush=True)

    print("✅ Таблица готова", flush=True)


def sheet_append(sheet: str, rows: list):
    _sheets_call(lambda s: s.spreadsheets().values().append(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{sheet}!A1",
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body={"values": rows},
    ).execute())


def sheet_get_all(sheet: str) -> list:
    result = _sheets_call(lambda s: s.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{sheet}!A:Z",
    ).execute())
    return result.get("values", [])


def sheet_update_cell(sheet: str, row: int, col: int, value: str):
    col_letter = chr(ord("A") + col - 1)
    _sheets_call(lambda s: s.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{sheet}!{col_letter}{row}",
        valueInputOption="RAW",
        body={"values": [[value]]},
    ).execute())


def sheet_batch_update_cells(sheet: str, updates: list[tuple[int, int, str]]):
    """updates: list of (row, col, value). One API call for all."""
    if not updates:
        return
    data = [{"range": f"{sheet}!{chr(ord('A') + col - 1)}{row}", "values": [[value]]}
            for row, col, value in updates]
    _sheets_call(lambda s: s.spreadsheets().values().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"valueInputOption": "RAW", "data": data},
    ).execute())


# ── Registry ───────────────────────────────────────────────────────────────────

def registry_find_row(filename: str) -> int | None:
    rows = sheet_get_all(REGISTRY_SHEET)
    for i, row in enumerate(rows[1:], start=2):
        if not row:
            continue
        if len(row) > 3 and row[3] == filename:
            return i
    return None


def registry_add(filename: str, file_id: str, file_type: str, chat_id: int,
                 message_id: int, caption: str = "", sender_name: str = ""):
    rows = sheet_get_all(REGISTRY_SHEET)
    seq_num = len(rows)
    now = _now().strftime("%d.%m.%Y %H:%M")
    # A=# B=Поставщик C=Сумма D=Имя E=Статус F=Получен G=Обработан H=Ошибка
    # I=file_id J=file_type K=chat_id L=Примечание M=Имя отправителя N=Оплата O=message_id P=Ссылка
    tg_link = _tg_link(chat_id, message_id)
    sheet_append(REGISTRY_SHEET, [[
        seq_num, "", "", filename, "⏳", now, "", "", file_id, file_type,
        str(chat_id), caption, sender_name, "", str(message_id), tg_link
    ]])


def registry_update(row_num: int, status: str, error: str = "",
                    supplier: str = "", total: str = ""):
    now = _now().strftime("%d.%m.%Y %H:%M") if status == "✅" else ""
    sheet_update_cell(REGISTRY_SHEET, row_num, 5, status)   # E = Статус
    sheet_update_cell(REGISTRY_SHEET, row_num, 7, now)      # G = Обработан
    sheet_update_cell(REGISTRY_SHEET, row_num, 8, error)    # H = Ошибка
    if status == "✅":
        if supplier:
            sheet_update_cell(REGISTRY_SHEET, row_num, 2, supplier)  # B = Поставщик
        if total:
            sheet_update_cell(REGISTRY_SHEET, row_num, 3, total)     # C = Сумма счета


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
    if total_lines:
        # Total goes before items so it's never cut off by token limits
        parts.append("ВСЕГО К ОПЛАТЕ: " + total_lines[-1])
    parts.append("\n".join(table_rows)[:5000])
    return "\n---\n".join(parts), False


def compress_for_vision(image_bytes: bytes, max_px: int = 1400) -> bytes:
    """Конвертирует в JPEG для Vision API. Поддерживает изображения и PDF."""
    from PIL import Image
    img = None
    try:
        img = Image.open(io.BytesIO(image_bytes))
    except Exception:
        # PDF — рендерим первую страницу через pdfplumber
        with pdfplumber.open(io.BytesIO(image_bytes)) as pdf:
            img = pdf.pages[0].to_image(resolution=200).original
    if img.mode in ('RGBA', 'P', 'LA'):
        img = img.convert('RGB')
    if max(img.size) > max_px:
        img.thumbnail((max_px, max_px), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format='JPEG', quality=85, optimize=True)
    result = buf.getvalue()
    print(f"🗜 {len(image_bytes)//1024}KB → {len(result)//1024}KB JPEG", flush=True)
    return result


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
- supplier: кто ВЫСТАВИЛ счёт — продавец/поставщик/исполнитель. Ищи явную строку "Поставщик (Исполнитель):" или "Поставщик:" — берёт название оттуда. Если такой строки нет — ищи компанию рядом с ИНН/КПП поставщика. НЕ бери из строки "Покупатель:" — там всегда ООО "Термодинамика" или похожий покупатель. НЕ бери название банка (слова БАНК, ПАО Сбербанк, БИК, Сч.№, "Банк получателя" — это платёжные реквизиты, не поставщик)
- invoice_number: КОРОТКИЙ номер из заголовка "Счёт № X" или "Счёт на оплату № X". Это 1-5 цифр/символов. НЕ бери 20-значный расчётный счёт (р/с) банка. Ищи строку вида "Счёт на оплату № 88 от..." — нужна цифра после "№"
- invoice_date: дата из той же строки заголовка "Счёт № X от ДД.ММ.ГГГГ" или "от 03 февраля 2026 г." — конвертируй в ДД.ММ.ГГГГ. Русские месяцы: января=01, февраля=02, марта=03, апреля=04, мая=05, июня=06, июля=07, августа=08, сентября=09, октября=10, ноября=11, декабря=12
- total_amount: итоговая сумма ВСЕГО счёта с НДС — ищи последнюю строку "Итого:" или "Всего к оплате:" в конце таблицы. Это одно конкретное число, НЕ суммируй сам
- pos: номер позиции в счёте (1, 2, 3...)
- article: артикул, тип, марка, производитель (если есть)
- price_with_vat: ЦЕНА ЗА ЕДИНИЦУ товара с НДС. Это НЕ сумма строки! Если есть колонки "Цена" и "Сумма" — бери из "Цена". Если цена без НДС — умножь на 1.2
- quantity: количество единиц товара
- Если поле неизвестно — пустая строка или 0"""


def _repair_json(s: str) -> str:
    """Итеративно исправляет невалидные escape-символы используя позицию из исключения."""
    for _ in range(30):
        try:
            json.loads(s)
            return s
        except json.JSONDecodeError as e:
            if 'escape' not in str(e).lower() or e.pos is None:
                raise
            pos = e.pos  # позиция обратного слеша
            if pos < len(s) and s[pos] == '\\':
                s = s[:pos] + '\\\\' + s[pos + 1:]
            else:
                raise
    raise ValueError("Не удалось исправить JSON escapes")

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
                raw = text[start:i + 1]
                try:
                    return json.loads(raw)
                except json.JSONDecodeError as e:
                    if 'escape' in str(e).lower():
                        return json.loads(_repair_json(raw))
                    raise
    raise ValueError("Незакрытый JSON")


def parse_with_ai(text: str) -> dict:
    for attempt in range(3):
        try:
            resp = client.chat.completions.create(
                model="llama-3.1-8b-instant",
                messages=[
                    {"role": "system", "content": EXTRACT_PROMPT},
                    {"role": "user", "content": f"Счёт:\n\n{text[:4000]}"},
                ],
                max_tokens=4000,
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
        max_tokens=4000,
        temperature=0.1,
    )
    return extract_json(resp.choices[0].message.content.strip())


PROMPT_PLACEHOLDERS = {"название поставщика", "номер счёта", "наименование", "артикул/тип или пусто"}

def clean_field(val: str, fallback: str = "—") -> str:
    if not val or val.strip().lower() in PROMPT_PLACEHOLDERS:
        return fallback
    return val.strip()

def invoice_to_rows(data: dict, filename: str, caption: str = "", sender_name: str = "", drive_link: str = "") -> list:
    today = _now().strftime("%d.%m.%Y")
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
            today, total_amount, caption, sender_name, "", drive_link,
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


_BANK_KEYWORDS    = ("банк", "сбербанк", "втб", "тинькофф", "альфа-банк", "открытие", "бик", "р/с")
_BUYER_KEYWORDS   = ("термодинамика",)
_BAD_INV_RE       = re.compile(r'^\d{10,}$')   # р/с вместо номера счёта

def _bad_extraction_reason(data: dict) -> str:
    """Возвращает причину если AI вернул мусор, иначе пустую строку."""
    supplier = (data.get("supplier") or "").strip().lower()
    inv_num  = str(data.get("invoice_number") or "").strip()

    if not supplier or supplier in PROMPT_PLACEHOLDERS:
        return "Поставщик не распознан"
    if any(k in supplier for k in _BANK_KEYWORDS):
        return f"Поставщик — банк ({supplier[:40]})"
    if any(k in supplier for k in _BUYER_KEYWORDS):
        return f"Поставщик — покупатель ({supplier[:40]})"
    if _BAD_INV_RE.match(inv_num):
        return f"Номер счёта похож на р/с ({inv_num[:20]})"
    return ""


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
            reason = _bad_extraction_reason(data)
            if reason:
                print(f"🖼 {reason} — пробую Vision", flush=True)
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
        # Upload to Drive before AI processing (bytes still fresh from Telegram)
        drive_link = ""
        try:
            file_bytes = download_file(file_id)
            drive_link = upload_to_drive(filename, file_bytes, file_type)
            del file_bytes
            if row_num and drive_link:
                sheet_update_cell(REGISTRY_SHEET, row_num, 16, drive_link)  # P = Ссылка
        except Exception as de:
            print(f"⚠️ Drive upload skipped: {de}", flush=True)

        data = process_file(file_id, file_type, filename)
        rows = invoice_to_rows(data, filename, caption, sender_name, drive_link)
        if not rows:
            raise ValueError("Позиции не найдены")
        sheet_append(INVOICES_SHEET, rows)
        supplier = (data.get("supplier") or "—").strip()
        total = data.get("total_amount", 0) or 0
        try:
            total_str = f"{float(total):,.0f}".replace(",", " ")
        except Exception:
            total_str = str(total)
        if row_num:
            registry_update(row_num, "✅", supplier=supplier, total=total_str)
        _react("🏆")
        try:
            text = (f"«{filename}»\n"
                    f"Поставщик: {supplier}\n"
                    f"Сумма: {total_str} ₽")
            markup = telebot.types.InlineKeyboardMarkup(row_width=2)
            markup.row(
                telebot.types.InlineKeyboardButton("💰 Отметить оплаченным", callback_data=f"pay:{message_id}"),
                telebot.types.InlineKeyboardButton("⏳ Оплатить позже", callback_data=f"later:{message_id}"),
            )
            bot.send_message(chat_id, text, reply_markup=markup, reply_to_message_id=message_id)
        except Exception as ex:
            print(f"⚠️ send pay button error: {ex}", flush=True)
    except Exception as e:
        err = str(e)
        is_limit = "429" in err or "413" in err or "timed out" in err.lower() or "timeout" in err.lower() or "connection" in err.lower()
        print(f"{'⏳' if is_limit else '⛔'} {filename}: {err[:200]}", flush=True)
        if is_limit:
            if row_num:
                registry_update(row_num, "❌", err[:200])   # ❌ = будет ретраиться
            _react("😴")
        else:
            if row_num:
                registry_update(row_num, "⛔", err[:200])   # ⛔ = не ретраить
            _react("🤬")
            try:
                markup = telebot.types.InlineKeyboardMarkup()
                markup.add(telebot.types.InlineKeyboardButton(
                    "🗑 Игнорировать этот файл", callback_data=f"ignore:{message_id}"
                ))
                bot.send_message(chat_id,
                    f"Не могу распознать счёт «{filename}»",
                    reply_markup=markup)
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
        if not row:
            continue
        row = row + [''] * max(0, 15 - len(row))
        if row[4] != "❌":
            continue
        sheet_update_cell(REGISTRY_SHEET, i, 5, "⏳")
        item = {
            "filename":    row[3],
            "file_id":     row[8],
            "file_type":   row[9],
            "chat_id":     int(row[10]) if row[10] else 0,
            "message_id":  _row_message_id(row),
            "caption":     row[11],
            "sender_name": row[12],
            "row_num":     i,
            "is_retry":    True,
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
            if row and len(row) > 1 and row[1] == filename]


def _row_message_id(row: list):
    """Возвращает message_id. Новый: idx 14 (O). Промежуточный: idx 12. Старый: idx 10."""
    row = row + [''] * max(0, 15 - len(row))
    for idx in (14, 12, 10):
        val = row[idx]
        if val:
            try:
                n = int(val)
                if n > 0:  # message_id всегда положительный
                    return n
            except (ValueError, TypeError):
                pass
    return None


def _row_chat_id(row: list) -> int:
    """Возвращает chat_id. Новый: idx 10, старый: idx 8. chat_id групп всегда отрицательный."""
    row = row + [''] * max(0, 15 - len(row))
    for idx in (10, 8):
        val = row[idx]
        if val:
            try:
                n = int(val)
                if n < 0:
                    return n
            except (ValueError, TypeError):
                pass
    return 0


def _set_payment_status(message_id: int, value: str, label: str):
    """Универсальная запись статуса оплаты в реестр и счета. Один batch-запрос."""
    try:
        rows = sheet_get_all(REGISTRY_SHEET)
        for i, row in enumerate(rows[1:], start=2):
            if not row or _row_message_id(row) != message_id:
                continue
            filename = row[3] if len(row) > 3 else row[1]
            invoice_rows = invoices_find_rows(filename)
            sheet_batch_update_cells(REGISTRY_SHEET, [(i, 14, value)])
            if invoice_rows:
                sheet_batch_update_cells(INVOICES_SHEET, [(j, 17, value) for j in invoice_rows])
            print(f"💰 {filename} → {label} ({len(invoice_rows)} строк)", flush=True)
            return
        print(f"⚠️ payment status: message_id={message_id} не найден", flush=True)
    except Exception as e:
        print(f"❌ payment status error: {e}", flush=True)


def mark_paid(message_id: int, chat_id: int, paid: bool):
    _set_payment_status(message_id, "✅" if paid else "", "оплачен ✅" if paid else "снята отметка")


def mark_later(message_id: int, link: str = ""):
    _set_payment_status(message_id, "⏳", "отложен ⏳")
    if link:
        try:
            rows = sheet_get_all(REGISTRY_SHEET)
            for i, row in enumerate(rows[1:], start=2):
                if row and _row_message_id(row) == message_id:
                    sheet_update_cell(REGISTRY_SHEET, i, 16, link)  # P = Ссылка
                    return
        except Exception as e:
            print(f"⚠️ mark_later link save error: {e}", flush=True)


def _tg_link(chat_id: int, message_id: int) -> str:
    cid = str(chat_id)
    if cid.startswith("-100"):
        return f"https://t.me/c/{cid[4:]}/{message_id}"
    return ""


def mark_ignored(message_id: int):
    try:
        rows = sheet_get_all(REGISTRY_SHEET)
        for i, row in enumerate(rows[1:], start=2):
            if not row or _row_message_id(row) != message_id:
                continue
            filename = row[3] if len(row) > 3 else row[1]
            sheet_update_cell(REGISTRY_SHEET, i, 5, "🚫")
            print(f"🚫 {filename} проигнорирован", flush=True)
            return
        print(f"⚠️ mark_ignored: message_id={message_id} не найден", flush=True)
    except Exception as e:
        print(f"❌ mark_ignored error: {e}", flush=True)


@bot.callback_query_handler(func=lambda c: c.data.startswith("ignore:"))
def on_ignore_callback(call):
    try:
        msg_id = int(call.data.split(":", 1)[1])
        Thread(target=mark_ignored, args=(msg_id,), daemon=True).start()
        bot.answer_callback_query(call.id)
        bot.edit_message_text(
            f"🚫 Файл проигнорирован",
            call.message.chat.id, call.message.id
        )
    except Exception as e:
        print(f"❌ ignore callback error: {e}", flush=True)
        try:
            bot.answer_callback_query(call.id, "Ошибка")
        except Exception:
            pass


@bot.callback_query_handler(func=lambda c: c.data.startswith("later:"))
def on_later_callback(call):
    try:
        msg_id = int(call.data.split(":", 1)[1])
        # Берём ссылку на оригинальный счёт из reply_to (бот отвечал на файл)
        reply_to = getattr(call.message, "reply_to_message", None)
        if reply_to:
            link = _tg_link(call.message.chat.id, reply_to.message_id)
        else:
            link = _tg_link(call.message.chat.id, call.message.id)
        Thread(target=mark_later, args=(msg_id, link), daemon=True).start()
        bot.answer_callback_query(call.id)
        markup = telebot.types.InlineKeyboardMarkup(row_width=2)
        markup.row(
            telebot.types.InlineKeyboardButton("💰 Отметить оплаченным", callback_data=f"pay:{msg_id}"),
            telebot.types.InlineKeyboardButton("✓ Отложено", callback_data=f"later:{msg_id}"),
        )
        bot.edit_message_reply_markup(call.message.chat.id, call.message.id, reply_markup=markup)
    except Exception as e:
        print(f"❌ later callback error: {e}", flush=True)
        try:
            bot.answer_callback_query(call.id, "Ошибка")
        except Exception:
            pass


@bot.callback_query_handler(func=lambda c: c.data.startswith(("pay:", "unpay:")))
def on_pay_callback(call):
    try:
        action, msg_id_str = call.data.split(":", 1)
        msg_id = int(msg_id_str)
        paid = action == "pay"
        Thread(target=mark_paid, args=(msg_id, call.message.chat.id, paid), daemon=True).start()
        bot.answer_callback_query(call.id)
        new_action = "unpay" if paid else "pay"
        new_label  = "✅ Оплачено" if paid else "💰 Отметить оплаченным"
        markup = telebot.types.InlineKeyboardMarkup()
        markup.add(telebot.types.InlineKeyboardButton(new_label, callback_data=f"{new_action}:{msg_id}"))
        bot.edit_message_reply_markup(call.message.chat.id, call.message.id, reply_markup=markup)
    except Exception as e:
        print(f"❌ pay callback error: {e}", flush=True)
        try:
            bot.answer_callback_query(call.id, "Ошибка")
        except Exception:
            pass


# ── Telegram handlers ──────────────────────────────────────────────────────────


@bot.message_handler(commands=["pending"])
def on_pending(msg):
    rows = sheet_get_all(REGISTRY_SHEET)
    pending = []
    total_sum = 0.0
    for row in rows[1:]:
        if not row:
            continue
        row = row + [''] * max(0, 16 - len(row))
        if row[13] != "⏳":
            continue
        supplier = row[1] or "—"
        amount_str = row[2] or "0"
        try:
            amount = float(re.sub(r'[^\d.]', '', amount_str.replace(',', '.')))
        except Exception:
            amount = 0.0
        file_id   = row[8]
        file_type = row[9]
        msg_id    = _row_message_id(row) or 0
        pending.append((supplier, amount, file_id, file_type, msg_id))
        total_sum += amount

    if not pending:
        bot.reply_to(msg, "Нет счетов, отложенных на потом.")
        return

    lines = ["📋 Счета к оплате:\n"]
    markup = telebot.types.InlineKeyboardMarkup(row_width=2)
    has_buttons = False
    for n, (supplier, amount, file_id, file_type, msg_id) in enumerate(pending, 1):
        amount_fmt = f"{amount:,.0f}".replace(",", " ")
        lines.append(f"{n}. {supplier} — {amount_fmt} ₽")
        if msg_id:
            has_buttons = True
            markup.row(
                telebot.types.InlineKeyboardButton(f"📎 #{n} Счёт", callback_data=f"show:{msg_id}"),
                telebot.types.InlineKeyboardButton(f"✅ #{n} Оплатить", callback_data=f"pay:{msg_id}"),
            )

    total_fmt = f"{total_sum:,.0f}".replace(",", " ")
    lines.append(f"\nИтого: {total_fmt} ₽")
    bot.reply_to(msg, "\n".join(lines), reply_markup=markup if has_buttons else None)


@bot.callback_query_handler(func=lambda c: c.data.startswith("show:"))
def on_show_callback(call):
    try:
        msg_id = int(call.data.split(":", 1)[1])
        bot.answer_callback_query(call.id)
        rows = sheet_get_all(REGISTRY_SHEET)
        for row in rows[1:]:
            if not row or _row_message_id(row) != msg_id:
                continue
            row = row + [''] * max(0, 10 - len(row))
            file_id   = row[8]
            file_type = row[9]
            filename  = row[3] if len(row) > 3 else row[1]
            if not file_id:
                bot.answer_callback_query(call.id, "Файл не найден")
                return
            if file_type == "photo":
                bot.send_photo(call.message.chat.id, file_id, caption=filename)
            else:
                bot.send_document(call.message.chat.id, file_id, caption=filename)
            return
        bot.answer_callback_query(call.id, "Не найдено в реестре")
    except Exception as e:
        print(f"❌ show callback error: {e}", flush=True)
        try:
            bot.answer_callback_query(call.id, "Ошибка")
        except Exception:
            pass


@bot.message_handler(commands=["chatid"])
def on_chatid(msg):
    cid = msg.chat.id
    chat_type = msg.chat.type
    link = _tg_link(cid, msg.message_id)
    bot.reply_to(msg, f"chat_id: {cid}\ntype: {chat_type}\nlink: {link or 'пусто'}")


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
    count = sum(1 for r in rows[1:] if len(r) > 4 and r[4] == "❌")
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
                if not row:
                    continue
                row_filename = row[3] if len(row) > 3 else (row[1] if len(row) > 1 else row[0])
                row_status   = row[4] if len(row) > 4 else (row[2] if len(row) > 2 else "")
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
        import traceback
        print(f"❌ enqueue {filename}: {e}\n{traceback.format_exc()}", flush=True)
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
    filename = f"photo_{_now().strftime('%Y%m%d_%H%M%S')}.jpg"
    enqueue_invoice(msg, photo.file_id, "photo", filename)


# ── Flask ──────────────────────────────────────────────────────────────────────

@app.route(f"/{TELEGRAM_TOKEN}", methods=["POST"])
def webhook():
    json_data = flask_request.get_json()
    update_keys = list(json_data.keys()) if json_data else []
    print(f"📨 update: {update_keys}", flush=True)

    update = telebot.types.Update.de_json(json_data)
    # Non-daemon for file messages: process survives SIGTERM until registry write completes
    msg_data = json_data.get("message", {}) if json_data else {}
    is_file = "document" in msg_data or "photo" in msg_data
    Thread(target=bot.process_new_updates, args=([update],), daemon=not is_file).start()
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
        telebot.types.BotCommand("/pending", "Список счетов отложенных на оплату"),
    ])

    Thread(target=_queue_worker, daemon=True).start()
    print("📋 Queue worker запущен", flush=True)

    def _keepalive():
        try:
            requests.get(WEBHOOK_URL, timeout=10)
        except Exception:
            pass

    scheduler = BackgroundScheduler()
    scheduler.add_job(retry_failed_invoices, "interval", minutes=20)
    scheduler.add_job(_keepalive, "interval", minutes=10)
    scheduler.start()

    print("Invoice bot запущен", flush=True)
    app.run(host="0.0.0.0", port=PORT)
