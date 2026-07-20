import hashlib
import hmac
import io
import json
import logging
import os
import re
import threading
import base64
import time
from datetime import datetime, timezone
from pathlib import Path

import psycopg
import requests
from flask import Flask, jsonify, request
from openpyxl import load_workbook


logging.basicConfig(level=os.getenv("LOG_LEVEL", "INFO"))
log = logging.getLogger("porsline-reporter")

app = Flask(__name__)
RUN_LOCK = threading.Lock()
BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "assets" / "report-template.xlsx.b64"

PORSLINE_BASE_URL = os.getenv("PORSLINE_BASE_URL", "https://survey.porsline.ir").rstrip("/")
PORSLINE_API_KEY = os.getenv("PORSLINE_API_KEY", "")
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID", "")
DATABASE_URL = os.getenv("DATABASE_URL", "")
INJECTION_SURVEY_CODE = os.getenv("INJECTION_SURVEY_CODE", "6Hf5AK7g")
TECHNICIAN_SURVEY_CODE = os.getenv("TECHNICIAN_SURVEY_CODE", "jiUT4eKo")
TAHER_INJECTION_SURVEY_CODE = os.getenv("TAHER_INJECTION_SURVEY_CODE", "mobh0bQS")
TAHER_SUTURE_SURVEY_CODE = os.getenv("TAHER_SUTURE_SURVEY_CODE", "sNUa7F2D")
ZOHIRI_SURVEY_CODE = os.getenv("ZOHIRI_SURVEY_CODE", "ox2HIlC4")
REPORT_PREFIX = os.getenv("REPORT_PREFIX", "تیر-تکنسین")
APP_SECRET = os.getenv("APP_SECRET", "")
TELEGRAM_POLLING_ENABLED = os.getenv("TELEGRAM_POLLING_ENABLED", "true").lower() == "true"
WEBHOOK_BASE_URL = os.getenv("WEBHOOK_BASE_URL", "").rstrip("/")
WEBHOOK_SECRET = os.getenv("WEBHOOK_SECRET", "")
BOT_ACCESS_CODE = os.getenv("BOT_ACCESS_CODE", "")

FIVE_REPORTS = [
    (INJECTION_SURVEY_CODE, "تزریقات"),
    (TECHNICIAN_SURVEY_CODE, "تکنسین داروخانه"),
    (TAHER_INJECTION_SURVEY_CODE, "تزریقات خانم طاهرخانی"),
    (TAHER_SUTURE_SURVEY_CODE, "بخیه خانم طاهرخانی"),
    (ZOHIRI_SURVEY_CODE, "مدارک خانم ظهیری"),
]

SINGLE_REPORT_COMMANDS = {
    "/zanyar_t": (INJECTION_SURVEY_CODE, "تزریقات"),
    "/zanyar_tek": (TECHNICIAN_SURVEY_CODE, "تکنسین داروخانه"),
    "/taher_t": (TAHER_INJECTION_SURVEY_CODE, "تزریقات خانم طاهرخانی"),
    "/taher_b": (TAHER_SUTURE_SURVEY_CODE, "بخیه خانم طاهرخانی"),
}

UTILITY_COMMANDS = {
    "/help", "/status", "/report", "/report_all", "/confirm_report_all",
    "/cancel_report", "/report_all_new", "/add_form", "/remove_form", "/forms",
}


def require_settings():
    missing = []
    for key, value in {
        "PORSLINE_API_KEY": PORSLINE_API_KEY,
        "TELEGRAM_BOT_TOKEN": TELEGRAM_BOT_TOKEN,
        "DATABASE_URL": DATABASE_URL,
    }.items():
        if not value:
            missing.append(key)
    if missing:
        raise RuntimeError("Missing settings: " + ", ".join(missing))


def db():
    return psycopg.connect(DATABASE_URL)


def init_db():
    require_settings()
    with db() as conn, conn.cursor() as cur:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS processed_responses (
                survey_code TEXT NOT NULL,
                response_key TEXT NOT NULL,
                processed_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
                PRIMARY KEY (survey_code, response_key)
            )
            """
        )
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS app_state (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            )
            """
        )
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS custom_forms (
                bot_command TEXT PRIMARY KEY,
                survey_code TEXT NOT NULL UNIQUE,
                report_name TEXT NOT NULL,
                created_by TEXT NOT NULL,
                active BOOLEAN NOT NULL DEFAULT TRUE,
                created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
            )
            """
        )
        conn.commit()


def get_state(key, default=None):
    with db() as conn, conn.cursor() as cur:
        cur.execute("SELECT value FROM app_state WHERE key=%s", (key,))
        row = cur.fetchone()
    return row[0] if row else default


def set_state(key, value):
    with db() as conn, conn.cursor() as cur:
        cur.execute(
            """
            INSERT INTO app_state(key, value) VALUES(%s, %s)
            ON CONFLICT(key) DO UPDATE SET value=EXCLUDED.value
            """,
            (key, str(value)),
        )
        conn.commit()


def get_custom_forms(active_only=True):
    query = "SELECT bot_command, survey_code, report_name FROM custom_forms"
    if active_only:
        query += " WHERE active=TRUE"
    query += " ORDER BY created_at"
    with db() as conn, conn.cursor() as cur:
        cur.execute(query)
        return cur.fetchall()


def get_report_for_command(command):
    if command in SINGLE_REPORT_COMMANDS:
        return SINGLE_REPORT_COMMANDS[command]
    with db() as conn, conn.cursor() as cur:
        cur.execute(
            "SELECT survey_code, report_name FROM custom_forms WHERE bot_command=%s AND active=TRUE",
            (command,),
        )
        row = cur.fetchone()
    return tuple(row) if row else None


def all_report_definitions():
    reports = list(SINGLE_REPORT_COMMANDS.values())
    reports.extend((survey_code, report_name) for _command, survey_code, report_name in get_custom_forms())
    return reports


def save_custom_form(command, survey_code, report_name, user_id):
    with db() as conn, conn.cursor() as cur:
        cur.execute(
            "SELECT bot_command, survey_code, active FROM custom_forms WHERE survey_code=%s OR bot_command=%s",
            (survey_code, command),
        )
        existing = cur.fetchone()
        if existing:
            old_command, old_code, active = existing
            if not active and old_command == command and old_code == survey_code:
                cur.execute(
                    """
                    UPDATE custom_forms
                    SET active=TRUE, report_name=%s, created_by=%s
                    WHERE bot_command=%s
                    """,
                    (report_name, user_id, command),
                )
                conn.commit()
                return
            raise ValueError("این لینک یا دستور قبلاً ثبت شده است.")
        cur.execute(
            """
            INSERT INTO custom_forms(bot_command, survey_code, report_name, created_by)
            VALUES(%s, %s, %s, %s)
            """,
            (command, survey_code, report_name, user_id),
        )
        conn.commit()


def deactivate_custom_form(command):
    with db() as conn, conn.cursor() as cur:
        cur.execute(
            "UPDATE custom_forms SET active=FALSE WHERE bot_command=%s AND active=TRUE",
            (command,),
        )
        changed = cur.rowcount > 0
        conn.commit()
    return changed


def porsline_get(path, params=None):
    response = requests.get(
        f"{PORSLINE_BASE_URL}{path}",
        params=params,
        headers={
            "Authorization": f"API-Key {PORSLINE_API_KEY}",
            "Content-Type": "application/json",
        },
        timeout=90,
    )
    response.raise_for_status()
    return response.json()


def resolve_surveys(survey_codes=None):
    survey_codes = set(survey_codes or {
        INJECTION_SURVEY_CODE,
        TECHNICIAN_SURVEY_CODE,
        TAHER_INJECTION_SURVEY_CODE,
        TAHER_SUTURE_SURVEY_CODE,
    })
    cached = {code: get_state(f"survey_id:{code}") for code in survey_codes}
    if all(cached.values()):
        return {code: int(survey_id) for code, survey_id in cached.items()}

    folders = porsline_get("/api/folders/")
    found = {}
    for folder in folders:
        for survey in folder.get("surveys", []):
            candidates = {
                str(survey.get("preview_code") or ""),
                str(survey.get("url_slug") or ""),
                str(survey.get("report_code") or ""),
            }
            for code in survey_codes:
                if code in candidates:
                    found[code] = int(survey["id"])

    missing = survey_codes - set(found)
    if missing:
        raise RuntimeError("Could not resolve survey code(s): " + ", ".join(sorted(missing)))
    for code, survey_id in found.items():
        set_state(f"survey_id:{code}", survey_id)
    return found


def fetch_results(survey_id):
    first = porsline_get(
        f"/api/v2/surveys/{survey_id}/responses/results-table/",
        params={"page": 1, "page_size": 1000},
    )
    headers = first.get("header", [])
    rows = list(first.get("body", []))
    total = int(first.get("responders_count", len(rows)))

    page = 2
    while len(rows) < total:
        batch = porsline_get(
            f"/api/v2/surveys/{survey_id}/responses/results-table/",
            params={"page": page, "page_size": 1000},
        ).get("body", [])
        if not batch:
            break
        rows.extend(batch)
        page += 1
    return headers, rows, total


def clean_header(value):
    return re.sub(r"[\s‌:\-_()（）]+", "", str(value or "")).lower()


def header_label(header, index):
    if isinstance(header, str):
        return header
    if isinstance(header, dict):
        for key in (
            "title", "name", "text", "label", "question_title", "question_text",
            "alt_name", "display_name", "header",
        ):
            value = header.get(key)
            if isinstance(value, str) and value.strip():
                return value
        for value in header.values():
            if isinstance(value, dict):
                nested = header_label(value, index)
                if nested != f"column_{index}":
                    return nested
    return f"column_{index}"


def scalar_value(value):
    if isinstance(value, list):
        if len(value) == 1:
            return scalar_value(value[0])
        return "، ".join(str(scalar_value(item)) for item in value if item not in (None, ""))
    if not isinstance(value, dict):
        return value
    for key in ("value", "answer", "text", "name", "display_value", "response", "result"):
        candidate = value.get(key)
        if candidate not in (None, ""):
            return scalar_value(candidate)
    return value


def row_to_mapping(headers, row):
    labels = [header_label(header, i) for i, header in enumerate(headers)]
    if isinstance(row, list):
        return {labels[i]: scalar_value(value) for i, value in enumerate(row) if i < len(labels)}
    if isinstance(row, dict):
        result = {str(key): scalar_value(value) for key, value in row.items()}
        nested_values = (
            row.get("values") or row.get("answers") or row.get("cells")
            or row.get("data") or row.get("row") or row.get("responses")
        )
        if isinstance(nested_values, list):
            result.update({labels[i]: scalar_value(value) for i, value in enumerate(nested_values) if i < len(labels)})
            for cell in nested_values:
                if not isinstance(cell, dict):
                    continue
                cell_label = header_label(cell, -1)
                cell_value = scalar_value(cell)
                if cell_label != "column_-1" and cell_value is not cell:
                    result[cell_label] = cell_value
        elif isinstance(nested_values, dict):
            nested_items = list(nested_values.items())
            result.update({labels[i]: scalar_value(item[1]) for i, item in enumerate(nested_items) if i < len(labels)})
            result.update({str(key): scalar_value(value) for key, value in nested_items})

        # Some API versions wrap the positional row in an undocumented list field.
        if not any(clean_header(label) in {clean_header(key) for key in result} for label in labels):
            for value in row.values():
                if isinstance(value, list) and len(value) >= len(labels):
                    result.update({labels[i]: scalar_value(value[i]) for i in range(len(labels))})
                    break
        for i, header in enumerate(headers):
            if not isinstance(header, dict):
                continue
            candidates = []
            for key in ("id", "key", "column_id", "object_id", "question_id"):
                if header.get(key) is not None:
                    candidates.extend([header[key], str(header[key])])
            for candidate in candidates:
                if candidate in row:
                    result[labels[i]] = scalar_value(row[candidate])
                    break
        return result
    raise ValueError(f"Unsupported Porsline row type: {type(row).__name__}")


def find_value(mapping, candidates):
    normalized = {clean_header(k): v for k, v in mapping.items()}
    for candidate in candidates:
        needle = clean_header(candidate)
        if needle in normalized and normalized[needle] not in (None, ""):
            return str(normalized[needle]).strip()
    for candidate in candidates:
        needle = clean_header(candidate)
        for key, value in normalized.items():
            if needle and needle in key and value not in (None, ""):
                return str(value).strip()
    return ""


def response_key(mapping):
    identifier = find_value(
        mapping,
        ["شناسه پاسخ دهنده", "شناسه پاسخ‌دهنده", "responder id", "response id", "id"],
    )
    if identifier:
        return identifier
    raw = json.dumps(mapping, ensure_ascii=False, sort_keys=True, default=str)
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def extract_person(mapping):
    persian_name = find_value(
        mapping,
        ["نام و نام‌خانوادگی(فارسی)", "نام و نام خانوادگی فارسی", "نام فارسی"],
    )
    english_name = find_value(
        mapping,
        ["نام و نام خانوادگی(انگلیسی)", "نام و نام خانوادگی انگلیسی", "نام انگلیسی"],
    )
    national_id = find_value(mapping, ["کد ملی", "کدملی", "national id", "national code"])
    national_id = re.sub(r"\D", "", national_id)
    return {
        "persian_name": persian_name,
        "english_name": english_name,
        "national_id": national_id,
    }


def is_processed(survey_code, key):
    with db() as conn, conn.cursor() as cur:
        cur.execute(
            "SELECT 1 FROM processed_responses WHERE survey_code=%s AND response_key=%s",
            (survey_code, key),
        )
        return cur.fetchone() is not None


def mark_processed(items):
    with db() as conn, conn.cursor() as cur:
        cur.executemany(
            """
            INSERT INTO processed_responses(survey_code, response_key)
            VALUES(%s, %s) ON CONFLICT DO NOTHING
            """,
            items,
        )
        conn.commit()


def national_id_formula(row_number):
    b = f'SUBSTITUTE($B{row_number},"/","")'
    return (
        f'=IFERROR(IF(AND(LEN({b})=10,AND(LEFT({b},10)<>REPT(ROW($1:$9),10)),'
        f'OR(AND(MOD(SUM(MID({b},ROW($1:$9),1)*(11-ROW($1:$9))),11)<2,'
        f'--RIGHT({b})=MOD(SUM(MID({b},ROW($1:$9),1)*(11-ROW($1:$9))),11)),'
        f'--RIGHT({b})=(11-MOD(SUM(MID({b},ROW($1:$9),1)*(11-ROW($1:$9))),11)))),'
        "TRUE,FALSE),FALSE)"
    )


def display_national_id(national_id):
    value = str(national_id or "").strip()
    return f"/{value}" if value.startswith("0") else value


def build_report(rows, total_count, report_name):
    template_bytes = base64.b64decode(TEMPLATE_PATH.read_text(encoding="ascii"))
    workbook = load_workbook(io.BytesIO(template_bytes))
    sheet = workbook.active

    if sheet.max_column > 5:
        sheet.delete_cols(6, sheet.max_column - 5)

    if sheet.max_row > 1:
        sheet.delete_rows(2, sheet.max_row - 1)

    sheet.sheet_view.rightToLeft = True
    sheet.cell(1, 1, "نام فارسی")
    sheet.cell(1, 2, "کد ملی فارسی")
    sheet.cell(1, 3, "TRUE")
    sheet.cell(1, 4, "نام انگلیسی")
    sheet.cell(1, 5, "کد ملی انگلیسی")

    for row_number, person in enumerate(rows, start=2):
        national_id = display_national_id(person["national_id"])
        sheet.cell(row_number, 1, person["persian_name"])
        sheet.cell(row_number, 2, national_id)
        sheet.cell(row_number, 3, national_id_formula(row_number))
        sheet.cell(row_number, 4, person["english_name"])
        sheet.cell(row_number, 5, national_id)
        for col in range(1, 6):
            cell = sheet.cell(row_number, col)
            if col in (2, 5):
                cell.number_format = "@"

    filename = f"{report_name} {total_count}.xlsx"
    stream = io.BytesIO()
    workbook.save(stream)
    stream.seek(0)
    return filename, stream


def owner_chat_id():
    chat_id = get_state("bot_owner_chat_id")
    if not chat_id:
        raise RuntimeError("Bot owner has not been registered")
    return str(chat_id)


def send_document(filename, stream, caption, chat_id=None):
    response = requests.post(
        f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendDocument",
        data={"chat_id": chat_id or owner_chat_id(), "caption": caption},
        files={"document": (filename, stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
        timeout=120,
    )
    response.raise_for_status()
    payload = response.json()
    if not payload.get("ok"):
        raise RuntimeError(f"Telegram rejected document: {payload}")


def send_message(text, chat_id=None, reply_markup=None):
    payload = {"chat_id": chat_id or owner_chat_id(), "text": text}
    if reply_markup is not None:
        payload["reply_markup"] = reply_markup
    response = requests.post(
        f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage",
        json=payload,
        timeout=45,
    )
    response.raise_for_status()


def is_group_admin(user_id):
    response = requests.get(
        f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/getChatMember",
        params={"chat_id": TELEGRAM_CHAT_ID, "user_id": user_id},
        timeout=45,
    )
    response.raise_for_status()
    payload = response.json()
    if not payload.get("ok"):
        return False
    return (payload.get("result") or {}).get("status") in {"creator", "administrator"}


def parse_add_form(text):
    match = re.match(
        r"^/add_form(?:@\w+)?\s+(/[a-z][a-z0-9_]{1,30})\s+"
        r"(https?://survey\.porsline\.ir/s/([A-Za-z0-9]+))\s*\|\s*(.+?)\s*$",
        text,
        flags=re.IGNORECASE,
    )
    if not match:
        raise ValueError(
            "فرمت دستور درست نیست. نمونه:\n"
            "/add_form /my_form https://survey.porsline.ir/s/AbC123 | نام فایل"
        )
    command = match.group(1).lower()
    survey_code = match.group(3)
    report_name = match.group(4).strip()
    if command in UTILITY_COMMANDS or command in SINGLE_REPORT_COMMANDS:
        raise ValueError("این دستور رزرو شده است؛ یک دستور متفاوت انتخاب کنید.")
    if len(report_name) > 80:
        raise ValueError("نام فایل باید حداکثر ۸۰ کاراکتر باشد.")
    return command, survey_code, report_name


def collect_new_rows(code, survey_id, include_processed=False):
    headers, rows, total = fetch_results(survey_id)
    log.info("Survey %s headers: %s", code, [header_label(h, i) for i, h in enumerate(headers)])
    if rows:
        first = rows[0]
        shape = {
            "row_type": type(first).__name__,
            "keys": list(first.keys()) if isinstance(first, dict) else None,
            "value_types": {str(k): type(v).__name__ for k, v in first.items()} if isinstance(first, dict) else None,
            "length": len(first) if isinstance(first, (dict, list)) else None,
        }
        log.info("Survey %s row shape: %s", code, shape)
    new_people = []
    keys = []
    for raw_row in rows:
        mapping = row_to_mapping(headers, raw_row)
        key = response_key(mapping)
        if not include_processed and is_processed(code, key):
            continue
        person = extract_person(mapping)
        required_fields = ("persian_name", "national_id")
        if not all(person[name] for name in required_fields):
            missing = [name for name in required_fields if not person[name]]
            log.warning("Skipped incomplete response %s from survey %s; missing=%s", key, code, missing)
            continue
        new_people.append(person)
        keys.append((code, key))
    return new_people, keys, total


def run_report(include_processed=False, all_forms=False, selected_reports=None):
    if not RUN_LOCK.acquire(blocking=False):
        return {"status": "already-running"}
    try:
        require_settings()
        init_db()
        reports = selected_reports or [
            (INJECTION_SURVEY_CODE, "تزریقات"),
            (TECHNICIAN_SURVEY_CODE, "تکنسین داروخانه"),
        ]
        if all_forms:
            reports = list(FIVE_REPORTS)
        ids = resolve_surveys({code for code, _ in reports})
        sent_files = []
        results = {}
        for code, report_name in reports:
            people, keys, total = collect_new_rows(code, ids[code], include_processed)
            results[code] = {"report": report_name, "rows": len(people), "total": total}
            if not people:
                continue
            filename, stream = build_report(people, total, report_name)
            row_label = "ردیف" if include_processed else "ردیف جدید"
            send_document(
                filename,
                stream,
                f"گزارش {report_name}: {len(people)} {row_label} از مجموع {total} پاسخ",
            )
            mark_processed(keys)
            sent_files.append(filename)

        if not sent_files:
            if include_processed:
                send_message("در فرم‌ها پاسخی برای ارسال وجود نداشت.")
            else:
                send_message("در این دوره پاسخ جدیدی برای ارسال وجود نداشت.")
        return {
            "status": "sent" if sent_files else "no-new-rows",
            "files": sent_files,
            "reports": results,
        }
    finally:
        RUN_LOCK.release()


def run_single_report(command):
    if not RUN_LOCK.acquire(blocking=False):
        return {"status": "already-running"}
    try:
        require_settings()
        init_db()
        report = get_report_for_command(command)
        if not report:
            return {"status": "unknown-command"}
        code, report_name = report
        ids = resolve_surveys({code})
        people, keys, total = collect_new_rows(code, ids[code])
        if not people:
            send_message(f"برای فرم {report_name} پاسخ جدیدی برای ارسال وجود نداشت.")
            return {"status": "no-new-rows", "report": report_name, "total": total}

        filename, stream = build_report(people, total, report_name)
        send_document(
            filename,
            stream,
            f"گزارش {report_name}: {len(people)} ردیف جدید از مجموع {total} پاسخ",
        )
        mark_processed(keys)
        return {
            "status": "sent",
            "report": report_name,
            "file": filename,
            "new_rows": len(people),
            "total": total,
        }
    finally:
        RUN_LOCK.release()


def run_status():
    if not RUN_LOCK.acquire(blocking=False):
        return None
    try:
        require_settings()
        init_db()
        reports = list(FIVE_REPORTS)
        ids = resolve_surveys({code for code, _ in reports})
        status_rows = []
        for code, report_name in reports:
            people, _keys, total = collect_new_rows(code, ids[code])
            status_rows.append((report_name, len(people), total))
        return status_rows
    finally:
        RUN_LOCK.release()


def help_text():
    return (
        "راهنمای دستورات ربات:\n\n"
        "/zanyar_t — پاسخ‌های جدید تزریقات\n"
        "/zanyar_tek — پاسخ‌های جدید تکنسین داروخانه\n"
        "/taher_t — پاسخ‌های جدید تزریقات خانم طاهرخانی\n"
        "/taher_b — پاسخ‌های جدید بخیه خانم طاهرخانی\n"
        "/report_all_new — پاسخ‌های جدید همه فرم‌ها\n"
        "/report_all — همه پاسخ‌های همه فرم‌ها (نیازمند تأیید)\n"
        "/status — تعداد کل و جدید همه فرم‌ها\n"
        "/forms — فهرست فرم‌ها و دستورهایشان\n"
        "/add_form — افزودن فرم جدید (فقط مدیر گروه)\n"
        "/remove_form — غیرفعال‌کردن فرم افزوده‌شده (فقط مدیر گروه)\n"
        "/help — نمایش همین راهنما"
    )


def friendly_error_message(exc):
    if isinstance(exc, (requests.Timeout, requests.ConnectionError)):
        return "ارتباط اینترنتی با پرسلاین یا تلگرام برقرار نشد. چند دقیقه بعد دوباره تلاش کنید."
    if isinstance(exc, requests.HTTPError):
        response = getattr(exc, "response", None)
        status_code = getattr(response, "status_code", None)
        if status_code in (401, 403):
            return "دسترسی به یکی از سرویس‌ها رد شد. لطفاً کلید API و توکن ربات را بررسی کنید."
        return "یکی از سرویس‌های پرسلاین یا تلگرام موقتاً پاسخ نداد. کمی بعد دوباره تلاش کنید."
    if isinstance(exc, psycopg.Error):
        return "ارتباط با پایگاه‌داده برقرار نشد. کمی بعد دوباره تلاش کنید."
    return "هنگام آماده‌سازی گزارش خطایی رخ داد. لطفاً چند دقیقه بعد دوباره تلاش کنید."


MAIN_MENU = {
    "keyboard": [
        [{"text": "دوره‌های من"}, {"text": "خانم طاهرخانی"}],
        [{"text": "ظهیری"}, {"text": "وضعیت"}],
        [{"text": "گزارش همه فرم‌ها"}],
    ],
    "resize_keyboard": True,
    "is_persistent": True,
}

MY_MENU = {
    "keyboard": [
        [{"text": "ثبتی‌های جدید تزریقات"}],
        [{"text": "ثبتی‌های جدید تکنسین"}],
        [{"text": "گزارش جدید دو دوره من"}],
        [{"text": "بازگشت به منوی اصلی"}],
    ],
    "resize_keyboard": True,
    "is_persistent": True,
}

TAHER_MENU = {
    "keyboard": [
        [{"text": "ثبتی‌های جدید تزریقات طاهرخانی"}],
        [{"text": "ثبتی‌های جدید بخیه طاهرخانی"}],
        [{"text": "گزارش جدید دو دوره طاهرخانی"}],
        [{"text": "بازگشت به منوی اصلی"}],
    ],
    "resize_keyboard": True,
    "is_persistent": True,
}

EXPORT_ACTIONS = {
    "mine_injection": ("ثبتی‌های جدید تزریقات", [FIVE_REPORTS[0]], False),
    "mine_technician": ("ثبتی‌های جدید تکنسین", [FIVE_REPORTS[1]], False),
    "mine_both": ("گزارش جدید دو دوره من", FIVE_REPORTS[:2], False),
    "taher_injection": ("ثبتی‌های جدید تزریقات طاهرخانی", [FIVE_REPORTS[2]], False),
    "taher_suture": ("ثبتی‌های جدید بخیه طاهرخانی", [FIVE_REPORTS[3]], False),
    "taher_both": ("گزارش جدید دو دوره طاهرخانی", FIVE_REPORTS[2:4], False),
    "zohiri": ("ثبتی‌های جدید مدارک خانم ظهیری", [FIVE_REPORTS[4]], False),
    "all_full": ("گزارش کامل همه پنج فرم", FIVE_REPORTS, True),
}

BUTTON_ACTIONS = {
    "ثبتی‌های جدید تزریقات": "mine_injection",
    "ثبتی‌های جدید تکنسین": "mine_technician",
    "گزارش جدید دو دوره من": "mine_both",
    "ثبتی‌های جدید تزریقات طاهرخانی": "taher_injection",
    "ثبتی‌های جدید بخیه طاهرخانی": "taher_suture",
    "گزارش جدید دو دوره طاهرخانی": "taher_both",
    "ظهیری": "zohiri",
    "گزارش همه فرم‌ها": "all_full",
}


def request_export_confirmation(action, user_id, chat_id):
    title = EXPORT_ACTIONS[action][0]
    set_state(
        f"pending_button_action:{user_id}",
        json.dumps({"action": action, "created_at": int(time.time())}),
    )
    markup = {
        "inline_keyboard": [[
            {"text": "تأیید و ارسال", "callback_data": f"confirm:{action}"},
            {"text": "لغو", "callback_data": "cancel"},
        ]]
    }
    send_message(
        f"هشدار: آیا از ساخت و ارسال «{title}» مطمئن هستید؟",
        chat_id=chat_id,
        reply_markup=markup,
    )


def execute_export_action(action):
    title, reports, include_processed = EXPORT_ACTIONS[action]
    send_message(f"تأیید شد؛ در حال آماده‌سازی «{title}»…")
    result = run_report(
        include_processed=include_processed,
        selected_reports=list(reports),
    )
    log.info("Button action %s result: %s", action, result)


def send_status():
    send_message("در حال بررسی وضعیت همه فرم‌ها…")
    rows = run_status()
    if rows is None:
        send_message("یک گزارش دیگر در حال آماده‌سازی است. کمی بعد دوباره تلاش کنید.")
        return
    lines = ["وضعیت فرم‌ها:"]
    for report_name, new_count, total in rows:
        lines.append(f"• {report_name}: {new_count} پاسخ جدید از مجموع {total}")
    send_message("\n".join(lines))


def process_private_message(text, user_id, chat_id):
    try:
        if text == "دوره‌های من":
            send_message("یکی از گزینه‌های دوره‌های من را انتخاب کنید:", chat_id, MY_MENU)
        elif text == "خانم طاهرخانی":
            send_message("یکی از گزینه‌های خانم طاهرخانی را انتخاب کنید:", chat_id, TAHER_MENU)
        elif text == "بازگشت به منوی اصلی":
            send_message("منوی اصلی:", chat_id, MAIN_MENU)
        elif text == "وضعیت":
            send_status()
        elif text in BUTTON_ACTIONS:
            request_export_confirmation(BUTTON_ACTIONS[text], user_id, chat_id)
    except Exception as exc:
        log.exception("Private menu action failed")
        try:
            send_message(friendly_error_message(exc), chat_id=chat_id)
        except Exception:
            log.exception("Could not send private menu error")


def process_callback(callback_id, data, user_id, chat_id):
    try:
        requests.post(
            f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/answerCallbackQuery",
            json={"callback_query_id": callback_id},
            timeout=20,
        ).raise_for_status()
        state_key = f"pending_button_action:{user_id}"
        pending = json.loads(get_state(state_key, "{}") or "{}")
        if data == "cancel":
            set_state(state_key, "{}")
            send_message("ارسال فایل لغو شد.", chat_id=chat_id)
            return
        if not data.startswith("confirm:"):
            return
        action = data.split(":", 1)[1]
        if (
            action not in EXPORT_ACTIONS
            or pending.get("action") != action
            or time.time() - float(pending.get("created_at", 0)) > 120
        ):
            send_message("مهلت تأیید تمام شده است؛ دوباره گزینه موردنظر را انتخاب کنید.", chat_id=chat_id)
            return
        set_state(state_key, "{}")
        execute_export_action(action)
    except Exception as exc:
        log.exception("Callback action failed")
        try:
            send_message(friendly_error_message(exc), chat_id=chat_id)
        except Exception:
            log.exception("Could not send callback error")


def telegram_get_updates(offset):
    response = requests.get(
        f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/getUpdates",
        params={"offset": offset, "timeout": 25, "allowed_updates": json.dumps(["message"])},
        timeout=35,
    )
    response.raise_for_status()
    payload = response.json()
    if not payload.get("ok"):
        raise RuntimeError(f"Telegram getUpdates failed: {payload}")
    return payload.get("result", [])


def telegram_poll_loop():
    while True:
        try:
            require_settings()
            init_db()
            offset = int(get_state("telegram_update_offset", "0"))
            for update in telegram_get_updates(offset):
                update_id = int(update["update_id"])
                offset = update_id + 1
                set_state("telegram_update_offset", offset)
                message = update.get("message") or {}
                chat_id = str((message.get("chat") or {}).get("id", ""))
                text = str(message.get("text") or "").strip()
                command = text.split()[0].split("@")[0].lower() if text else ""
                if chat_id != str(TELEGRAM_CHAT_ID) or not command.startswith("/"):
                    continue
                user_id = str((message.get("from") or {}).get("id", "unknown"))
                process_command_safely(command, user_id, text)
        except Exception:
            log.exception("Telegram polling failed")
            time.sleep(10)


def handle_report_command(command, user_id="unknown", full_text=""):
    if command == "/help":
        send_message(help_text())
        return
    if command == "/forms":
        init_db()
        lines = ["فرم‌های فعال:"]
        for form_command, (_code, report_name) in SINGLE_REPORT_COMMANDS.items():
            lines.append(f"• {form_command} — {report_name}")
        for form_command, _code, report_name in get_custom_forms():
            lines.append(f"• {form_command} — {report_name}")
        send_message("\n".join(lines))
        return
    if command == "/add_form":
        if not is_group_admin(user_id):
            send_message("فقط مدیران گروه اجازه افزودن فرم جدید را دارند.")
            return
        init_db()
        new_command, survey_code, report_name = parse_add_form(full_text)
        if survey_code in {code for code, _name in SINGLE_REPORT_COMMANDS.values()}:
            send_message("این فرم از قبل در ربات ثبت شده است.")
            return
        resolve_surveys({survey_code})
        save_custom_form(new_command, survey_code, report_name, user_id)
        send_message(
            f"فرم «{report_name}» با موفقیت اضافه شد.\n"
            f"دستور دریافت پاسخ‌های جدید: {new_command}"
        )
        return
    if command == "/remove_form":
        if not is_group_admin(user_id):
            send_message("فقط مدیران گروه اجازه غیرفعال‌کردن فرم را دارند.")
            return
        parts = full_text.split()
        if len(parts) != 2 or not re.fullmatch(r"/[a-z][a-z0-9_]{1,30}", parts[1].lower()):
            send_message("فرمت درست:\n/remove_form /command")
            return
        target = parts[1].lower()
        if target in SINGLE_REPORT_COMMANDS:
            send_message("چهار فرم اصلی از داخل تلگرام قابل حذف نیستند.")
            return
        init_db()
        if deactivate_custom_form(target):
            send_message(f"فرم مربوط به دستور {target} غیرفعال شد.")
        else:
            send_message("فرم فعالی با این دستور پیدا نشد.")
        return
    if command == "/status":
        send_message("در حال بررسی وضعیت همه فرم‌ها…")
        rows = run_status()
        if rows is None:
            send_message("یک گزارش دیگر در حال آماده‌سازی است. کمی بعد دوباره تلاش کنید.")
            return
        lines = ["وضعیت فرم‌ها:"]
        for report_name, new_count, total in rows:
            lines.append(f"• {report_name}: {new_count} پاسخ جدید از مجموع {total}")
        send_message("\n".join(lines))
        return
    init_db()
    selected_report = get_report_for_command(command)
    if selected_report:
        report_name = selected_report[1]
        send_message(f"در حال آماده‌سازی گزارش {report_name}…")
        result = run_single_report(command)
        log.info("Single report command %s result: %s", command, result)
        return
    if command == "/report_all":
        init_db()
        set_state(f"report_all_confirmation:{user_id}", int(time.time()))
        send_message(
            "این دستور تمام پاسخ‌های همه فرم‌های فعال را دوباره ارسال می‌کند. "
            "برای تأیید، حداکثر تا دو دقیقه دستور /confirm_report_all را بفرستید. "
            "برای لغو /cancel_report را بفرستید."
        )
        return
    elif command == "/confirm_report_all":
        init_db()
        state_key = f"report_all_confirmation:{user_id}"
        requested_at = float(get_state(state_key, "0") or 0)
        if time.time() - requested_at > 120:
            send_message("درخواست تأیید وجود ندارد یا مهلت دو دقیقه‌ای آن تمام شده است. دوباره /report_all را بفرستید.")
            return
        set_state(state_key, "0")
        send_message("تأیید شد؛ در حال آماده‌سازی گزارش کامل همه فرم‌ها…")
        result = run_report(include_processed=True, all_forms=True)
    elif command == "/cancel_report":
        init_db()
        set_state(f"report_all_confirmation:{user_id}", "0")
        send_message("ارسال گزارش کامل لغو شد.")
        return
    elif command == "/report_all_new":
        send_message("در حال بررسی پاسخ‌های جدید همه فرم‌ها…")
        result = run_report(include_processed=False, all_forms=True)
    elif command == "/report":
        send_message("در حال آماده‌سازی دو گزارش جدید…")
        result = run_report()
    else:
        return
    log.info("Command report result: %s", result)


def process_command_safely(command, user_id="unknown", full_text=""):
    try:
        handle_report_command(command, user_id, full_text)
    except ValueError as exc:
        send_message(str(exc))
    except Exception as exc:
        log.exception("Telegram command %s failed", command)
        try:
            send_message(friendly_error_message(exc))
        except Exception:
            log.exception("Could not send friendly error message")


def register_telegram_webhook():
    if not WEBHOOK_BASE_URL or not WEBHOOK_SECRET:
        return
    response = requests.post(
        f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/setWebhook",
        json={
            "url": f"{WEBHOOK_BASE_URL}/telegram-webhook",
            "secret_token": WEBHOOK_SECRET,
            "allowed_updates": ["message", "callback_query"],
            "drop_pending_updates": True,
        },
        timeout=45,
    )
    response.raise_for_status()
    payload = response.json()
    if not payload.get("ok"):
        raise RuntimeError(f"Telegram rejected webhook: {payload}")
    log.info("Telegram webhook registered")


@app.get("/")
@app.get("/health")
def health():
    return jsonify({"ok": True, "service": "porsline-telegram-reporter"})


@app.post("/telegram-webhook")
def telegram_webhook():
    if not WEBHOOK_SECRET or request.headers.get("X-Telegram-Bot-Api-Secret-Token") != WEBHOOK_SECRET:
        return jsonify({"ok": False, "error": "unauthorized"}), 401
    update = request.get_json(silent=True) or {}
    callback = update.get("callback_query")
    if callback:
        callback_message = callback.get("message") or {}
        callback_chat = callback_message.get("chat") or {}
        user_id = str((callback.get("from") or {}).get("id", ""))
        chat_id = str(callback_chat.get("id", ""))
        if callback_chat.get("type") != "private":
            return jsonify({"ok": True})
        init_db()
        if user_id != str(get_state("bot_owner_user_id", "")):
            return jsonify({"ok": True})
        threading.Thread(
            target=process_callback,
            args=(str(callback.get("id", "")), str(callback.get("data", "")), user_id, chat_id),
            daemon=True,
        ).start()
        return jsonify({"ok": True})

    message = update.get("message") or {}
    chat = message.get("chat") or {}
    if chat.get("type") != "private":
        return jsonify({"ok": True})
    chat_id = str(chat.get("id", ""))
    user_id = str((message.get("from") or {}).get("id", ""))
    text = str(message.get("text") or "").strip()
    init_db()
    registered_owner = str(get_state("bot_owner_user_id", ""))

    if not registered_owner:
        parts = text.split(maxsplit=1)
        supplied_code = parts[1].strip() if len(parts) == 2 and parts[0].lower() == "/start" else ""
        if not BOT_ACCESS_CODE:
            send_message("رمز ورود ربات هنوز در Render تنظیم نشده است.", chat_id=chat_id)
        elif supplied_code and hmac.compare_digest(supplied_code, BOT_ACCESS_CODE):
            set_state("bot_owner_user_id", user_id)
            set_state("bot_owner_chat_id", chat_id)
            send_message("دسترسی اختصاصی شما فعال شد. از منوی زیر استفاده کنید:", chat_id, MAIN_MENU)
        else:
            send_message("برای فعال‌سازی، دستور /start را همراه رمز اختصاصی وارد کنید.", chat_id=chat_id)
        return jsonify({"ok": True})

    if user_id != registered_owner:
        send_message("شما اجازه استفاده از این ربات را ندارید.", chat_id=chat_id)
        return jsonify({"ok": True})

    set_state("bot_owner_chat_id", chat_id)
    if text and text.split(maxsplit=1)[0].lower() == "/start":
        send_message("منوی اصلی:", chat_id, MAIN_MENU)
    elif text:
        threading.Thread(target=process_private_message, args=(text, user_id, chat_id), daemon=True).start()
    return jsonify({"ok": True})


@app.post("/run-now")
def run_now():
    supplied = request.headers.get("X-App-Secret") or request.args.get("secret")
    if not APP_SECRET or supplied != APP_SECRET:
        return jsonify({"ok": False, "error": "unauthorized"}), 401
    try:
        return jsonify({"ok": True, **run_report()})
    except Exception as exc:
        log.exception("Manual report failed")
        return jsonify({"ok": False, "error": str(exc)}), 500


if TELEGRAM_POLLING_ENABLED and os.getenv("DISABLE_TELEGRAM_POLLING", "false").lower() != "true":
    threading.Thread(target=telegram_poll_loop, daemon=True).start()

if WEBHOOK_BASE_URL and WEBHOOK_SECRET:
    threading.Thread(target=register_telegram_webhook, daemon=True).start()


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", "8080")))
