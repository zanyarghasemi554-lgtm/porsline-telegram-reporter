import hashlib
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
REPORT_PREFIX = os.getenv("REPORT_PREFIX", "تیر-تکنسین")
APP_SECRET = os.getenv("APP_SECRET", "")
TELEGRAM_POLLING_ENABLED = os.getenv("TELEGRAM_POLLING_ENABLED", "true").lower() == "true"


def require_settings():
    missing = []
    for key, value in {
        "PORSLINE_API_KEY": PORSLINE_API_KEY,
        "TELEGRAM_BOT_TOKEN": TELEGRAM_BOT_TOKEN,
        "TELEGRAM_CHAT_ID": TELEGRAM_CHAT_ID,
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


def resolve_surveys():
    cached = {
        INJECTION_SURVEY_CODE: get_state(f"survey_id:{INJECTION_SURVEY_CODE}"),
        TECHNICIAN_SURVEY_CODE: get_state(f"survey_id:{TECHNICIAN_SURVEY_CODE}"),
    }
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
            for code in (INJECTION_SURVEY_CODE, TECHNICIAN_SURVEY_CODE):
                if code in candidates:
                    found[code] = int(survey["id"])

    missing = {INJECTION_SURVEY_CODE, TECHNICIAN_SURVEY_CODE} - set(found)
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
    b = f"$B{row_number}"
    return (
        f'=IFERROR(IF(AND(LEN({b})=10,AND(LEFT({b},10)<>REPT(ROW($1:$9),10)),'
        f'OR(AND(MOD(SUM(MID({b},ROW($1:$9),1)*(11-ROW($1:$9))),11)<2,'
        f'--RIGHT({b})=MOD(SUM(MID({b},ROW($1:$9),1)*(11-ROW($1:$9))),11)),'
        f'--RIGHT({b})=(11-MOD(SUM(MID({b},ROW($1:$9),1)*(11-ROW($1:$9))),11)))),'
        "TRUE,FALSE),FALSE)"
    )


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
        sheet.cell(row_number, 1, person["persian_name"])
        sheet.cell(row_number, 2, person["national_id"])
        sheet.cell(row_number, 3, national_id_formula(row_number))
        sheet.cell(row_number, 4, person["english_name"])
        sheet.cell(row_number, 5, person["national_id"])
        for col in range(1, 6):
            cell = sheet.cell(row_number, col)
            if col in (2, 5):
                cell.number_format = "@"

    filename = f"{report_name} {total_count}.xlsx"
    stream = io.BytesIO()
    workbook.save(stream)
    stream.seek(0)
    return filename, stream


def send_document(filename, stream, caption):
    response = requests.post(
        f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendDocument",
        data={"chat_id": TELEGRAM_CHAT_ID, "caption": caption},
        files={"document": (filename, stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
        timeout=120,
    )
    response.raise_for_status()
    payload = response.json()
    if not payload.get("ok"):
        raise RuntimeError(f"Telegram rejected document: {payload}")


def send_message(text):
    response = requests.post(
        f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage",
        json={"chat_id": TELEGRAM_CHAT_ID, "text": text},
        timeout=45,
    )
    response.raise_for_status()


def collect_new_rows(code, survey_id):
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
        if is_processed(code, key):
            continue
        person = extract_person(mapping)
        if not all(person.values()):
            missing = [name for name, value in person.items() if not value]
            log.warning("Skipped incomplete response %s from survey %s; missing=%s", key, code, missing)
            continue
        new_people.append(person)
        keys.append((code, key))
    return new_people, keys, total


def run_report():
    if not RUN_LOCK.acquire(blocking=False):
        return {"status": "already-running"}
    try:
        require_settings()
        init_db()
        ids = resolve_surveys()
        injection, injection_keys, injection_total = collect_new_rows(
            INJECTION_SURVEY_CODE, ids[INJECTION_SURVEY_CODE]
        )
        technician, technician_keys, technician_total = collect_new_rows(
            TECHNICIAN_SURVEY_CODE, ids[TECHNICIAN_SURVEY_CODE]
        )
        sent_files = []
        if injection:
            filename, stream = build_report(injection, injection_total, "تزریقات")
            send_document(
                filename,
                stream,
                f"گزارش تزریقات: {len(injection)} ردیف جدید از مجموع {injection_total} پاسخ",
            )
            mark_processed(injection_keys)
            sent_files.append(filename)

        if technician:
            filename, stream = build_report(technician, technician_total, "تکنسین داروخانه")
            send_document(
                filename,
                stream,
                f"گزارش تکنسین داروخانه: {len(technician)} ردیف جدید از مجموع {technician_total} پاسخ",
            )
            mark_processed(technician_keys)
            sent_files.append(filename)

        if not sent_files:
            send_message("در این دوره پاسخ جدیدی برای ارسال وجود نداشت.")
        return {
            "status": "sent" if sent_files else "no-new-rows",
            "files": sent_files,
            "new_injection": len(injection),
            "new_technician": len(technician),
            "injection_total": injection_total,
            "technician_total": technician_total,
        }
    finally:
        RUN_LOCK.release()


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
                if chat_id != str(TELEGRAM_CHAT_ID) or command != "/report":
                    continue
                send_message("در حال آماده‌سازی دو گزارش جدید…")
                result = run_report()
                log.info("Command report result: %s", result)
        except Exception:
            log.exception("Telegram polling failed")
            time.sleep(10)


@app.get("/")
@app.get("/health")
def health():
    return jsonify({"ok": True, "service": "porsline-telegram-reporter"})


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


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", "8080")))
