#!/usr/bin/env python3
"""
Outstanding Reminders: weekly/monthly/daily payment reminders to vendors on WhatsApp (or SMS).

    pip install -r requirements.txt
    python app.py                      # opens http://127.0.0.1:5050
    python app.py --host 0.0.0.0       # let others on the office network use it
    python app.py --auto-send          # send the reminders due today and exit (used by Task Scheduler)

Server, Excel reading, validation, message building, per-vendor schedules, SQLite storage
(reminders.db next to the app) and sending through Meta's WhatsApp Cloud API or an Android
phone running RestSMS. The screens are in ui/. Starts in dry run: nothing is sent until you
choose WhatsApp or SMS in Settings.
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import hmac
import json
import io
import math
import os
import re
import secrets
import sqlite3
import subprocess
import sys
import tempfile
import threading
import time
import uuid
import webbrowser
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import date, datetime, timedelta, timezone
from functools import wraps
from pathlib import Path
from string import Formatter
from xml.sax.saxutils import escape as xml_escape

import requests
from flask import Flask, Response, jsonify, redirect, request, send_file, session
from werkzeug.security import check_password_hash, generate_password_hash
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

IST = timezone(timedelta(hours=5, minutes=30))   # India has no daylight saving; no tz database needed
APP_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
DB_PATH = Path(os.environ.get("SMS_APP_DB", APP_DIR / "reminders.db"))
WA_API_BASE = os.environ.get("WA_API_BASE", "https://graph.facebook.com")   # override only for testing
PASSWORD = os.environ.get("SMS_APP_PASSWORD", "change-me")


def now_ist() -> datetime:
    return datetime.now(IST)


# ===========================================================================
# Core logic: numbers, amounts, validation, messages, SMS length
# ===========================================================================
NAME_HINTS = ["vendor name", "party name", "customer name", "name", "party", "ledger", "vendor"]
PHONE_HINTS = ["mobile number", "mobile no", "mobile", "phone", "contact", "whatsapp"]
AMOUNT_HINTS = ["outstanding", "amount due", "balance", "amount", "due", "pending"]


def is_blank(value) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    return str(value).strip().lower() in {"", "nan", "none"}


def show_value(value) -> str:
    if is_blank(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def detect_column(headers: list[str], hints: list[str]) -> int | None:
    lowered = [h.strip().lower() for h in headers]
    for hint in hints:
        if hint in lowered:
            return lowered.index(hint)
    for hint in hints:
        for i, low in enumerate(lowered):
            if hint in low:
                return i
    return None


def normalize_mobile(raw) -> tuple[str | None, str]:
    """Indian mobile -> (10 digits, "") or (None, reason). Accepts +91/91/0091/0 prefixes."""
    if is_blank(raw):
        return None, "Mobile number missing"
    text = show_value(raw)
    if re.fullmatch(r"\d+\.0+", text):
        text = text.split(".")[0]
    digits = re.sub(r"[\s\-().]", "", text)
    if digits.startswith("+"):
        digits = digits[1:]
    if not digits.isdigit():
        return None, f"Has letters or symbols: {text}"
    if len(digits) == 14 and digits.startswith("0091"):
        digits = digits[4:]
    elif len(digits) == 12 and digits.startswith("91"):
        digits = digits[2:]
    elif len(digits) == 11 and digits.startswith("0"):
        digits = digits[1:]
    if len(digits) != 10:
        return None, f"Not 10 digits: {text}"
    if digits[0] not in "6789":
        return None, f"Not a mobile number (landline?): {text}"
    return digits, ""


def parse_amount(raw) -> float | None:
    """25000, '25,000', 'Rs. 25,000.00', '₹18,000', '12,500 Dr', '3,000 Cr' (negative), '(4,000)'."""
    if is_blank(raw):
        return None
    if isinstance(raw, (int, float)):
        return float(raw)
    text = str(raw).strip()
    sign = 1
    low = text.lower()
    if low.endswith("cr"):
        sign, text = -1, text[:-2]
    elif low.endswith("dr"):
        text = text[:-2]
    text = re.sub(r"(?i)rs\.?|inr|₹", "", text).strip()
    if text.startswith("(") and text.endswith(")"):
        sign, text = -sign, text[1:-1]
    cleaned = re.sub(r"[^\d.\-]", "", text)
    if cleaned.startswith("-"):
        sign, cleaned = -sign, cleaned[1:]
    try:
        return sign * float(cleaned)
    except ValueError:
        return None


def format_inr(amount: float) -> str:
    """125000 -> '1,25,000'; paise only when non-zero."""
    negative = amount < 0
    amount = abs(amount)
    rupees = int(amount)
    paise = round((amount - rupees) * 100)
    if paise == 100:
        rupees, paise = rupees + 1, 0
    text = str(rupees)
    if len(text) > 3:
        head, tail = text[:-3], text[-3:]
        text = re.sub(r"(\d)(?=(\d{2})+$)", r"\1,", head) + "," + tail
    if paise:
        text += f".{paise:02d}"
    return f"-{text}" if negative else text


def read_sheet(file_name: str, data: bytes) -> tuple[list[str], list[dict]]:
    """First sheet of an .xlsx (or a .csv) -> (headers, rows). Each row keeps its Excel row number."""
    lower = file_name.lower()
    if lower.endswith(".csv"):
        rows = list(csv.reader(io.StringIO(data.decode("utf-8-sig", errors="replace"))))
    elif lower.endswith((".xlsx", ".xlsm")):
        try:
            wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        except Exception as exc:
            raise ValueError(f"Couldn't open this Excel file ({exc.__class__.__name__}). "
                             "If it's password-protected or .xls, save it as .xlsx and try again.")
        rows = [list(r) for r in wb.worksheets[0].iter_rows(values_only=True)]
        wb.close()
    else:
        raise ValueError("Upload an .xlsx or .csv file.")

    start = next((i for i, r in enumerate(rows) if any(not is_blank(v) for v in r)), None)
    if start is None:
        raise ValueError("The file is empty.")
    headers = [show_value(h) or f"Column {i + 1}" for i, h in enumerate(rows[start])]
    while headers and headers[-1].startswith("Column ") and all(
            len(r) < len(headers) or is_blank(r[len(headers) - 1]) for r in rows[start + 1:]):
        headers.pop()                                  # drop empty trailing columns
    if len(headers) < 3:
        raise ValueError("The first sheet needs at least 3 columns: name, mobile number and amount.")
    data_rows = []
    for excel_row, r in enumerate(rows[start + 1:], start=start + 2):
        values = (list(r) + [None] * len(headers))[:len(headers)]
        if any(not is_blank(v) for v in values):
            data_rows.append({"excel_row": excel_row, "values": values})
    if not data_rows:
        raise ValueError("The file has headings but no data rows.")
    return headers, data_rows


def validate(rows: list[dict], name_i: int, phone_i: int, amount_i: int) -> dict:
    """Check each row. Same number + same name -> merged (amounts added).
    Same number + different names -> blocked. Zero/negative -> skipped."""
    candidates, problems = [], []
    for row in rows:
        v = row["values"]
        show = {"excel_row": row["excel_row"], "name": show_value(v[name_i]),
                "mobile": show_value(v[phone_i]), "amount": show_value(v[amount_i])}
        if not show["name"]:
            problems.append({**show, "problem": "Party name missing", "type": "Invalid"})
            continue
        phone, reason = normalize_mobile(v[phone_i])
        if phone is None:
            problems.append({**show, "problem": reason, "type": "Invalid"})
            continue
        amount = parse_amount(v[amount_i])
        if amount is None:
            problems.append({**show, "problem": "Amount missing or unreadable", "type": "Invalid"})
            continue
        if amount <= 0:
            problems.append({**show, "problem": "Nothing due (zero or negative)", "type": "Skipped"})
            continue
        candidates.append({"name": show["name"], "phone": phone, "amount": amount, "show": show})

    by_phone: dict[str, list[dict]] = {}
    for c in candidates:
        by_phone.setdefault(c["phone"], []).append(c)
    ready, merged = [], 0
    for phone, group in by_phone.items():
        if len({g["name"].casefold() for g in group}) > 1:
            names = ", ".join(sorted({g["name"] for g in group}))
            for g in group:
                problems.append({**g["show"], "problem": f"Same number used for different parties: {names}",
                                 "type": "Invalid"})
            continue
        merged += len(group) - 1
        total = round(sum(g["amount"] for g in group), 2)
        ready.append({"name": group[0]["name"], "phone": phone, "amount": total,
                      "amount_text": format_inr(total),
                      "excel_rows": ", ".join(str(g["show"]["excel_row"]) for g in group)})
    problems.sort(key=lambda p: p["excel_row"])
    return {
        "records": len(rows),
        "ready_count": len(ready),
        "invalid": sum(p["type"] == "Invalid" for p in problems),
        "skipped": sum(p["type"] == "Skipped" for p in problems),
        "merged": merged,
        "ready": ready,
        "problems": problems,
    }


PLACEHOLDERS = ["name", "amount", "date", "company", "contact"]
DEFAULT_TEMPLATE = ("Dear {name}, Rs. {amount} is due to {company} as of {date}. "
                    "Please pay or call {contact}. Ignore if paid.")
WA_MAX_CHARS = 1024
VAR_LABEL = {"name": "Name", "amount": "Amount", "date": "Date"}


def meta_template(template: str, company: str, contact: str) -> dict:
    """App message -> WhatsApp template text. {company}/{contact} become fixed text;
    {name}/{amount}/{date} become {{1}}, {{2}}, {{3}} in order of first appearance."""
    order: list[str] = []

    def swap(m):
        key = m.group(1)
        if key == "company":
            return company
        if key == "contact":
            return contact
        if key not in order:
            order.append(key)
        return "{{" + str(order.index(key) + 1) + "}}"

    text = re.sub(r"\{(name|amount|date|company|contact)\}", swap, template)
    warnings = []
    if re.match(r"^\s*\{\{\d+\}\}", text) or re.search(r"\{\{\d+\}\}\s*[.!?]?\s*$", text):
        warnings.append("WhatsApp rejects templates that start or end with a variable. Add a word before or after it.")
    if len(text) > WA_MAX_CHARS:
        warnings.append(f"Too long for a WhatsApp template ({len(text)} of {WA_MAX_CHARS} characters).")
    return {"text": text, "order": order, "warnings": warnings,
            "variables": [f"{{{{{i + 1}}}}} = {VAR_LABEL[k]}" for i, k in enumerate(order)]}


def normalise_text(t: str) -> str:
    return re.sub(r"\s+", " ", t or "").strip()


def unknown_placeholders(template: str) -> list[str]:
    try:
        fields = [f for _, f, _, _ in Formatter().parse(template) if f is not None]
    except ValueError:
        return ["a { or } without its partner"]
    return sorted({f for f in fields if f not in PLACEHOLDERS})


def render(template: str, *, name: str, amount: float, as_of: date, company: str, contact: str) -> str:
    return template.format(name=name, amount=format_inr(amount), date=as_of.strftime("%d-%b-%Y"),
                           company=company, contact=contact)


GSM7_BASIC = set("@£$¥èéùìòÇ\nØø\rÅåΔ_ΦΓΛΩΠΨΣΘΞÆæßÉ !\"#¤%&'()*+,-./0123456789:;<=>?"
                 "¡ABCDEFGHIJKLMNOPQRSTUVWXYZÄÖÑÜ§¿abcdefghijklmnopqrstuvwxyzäöñüà")
GSM7_EXTENDED = set("^{}\\[~]|€\f")


def sms_stats(text: str) -> dict:
    """GSM-7: 160 per SMS (153 when split). Unicode: 70 (67 when split)."""
    forcing = sorted({c for c in text if c not in GSM7_BASIC and c not in GSM7_EXTENDED})
    if forcing:
        length, single, multi, enc = sum(2 if ord(c) > 0xFFFF else 1 for c in text), 70, 67, "Unicode"
    else:
        length, single, multi, enc = sum(2 if c in GSM7_EXTENDED else 1 for c in text), 160, 153, "GSM-7"
    segments = 0 if length == 0 else 1 if length <= single else math.ceil(length / multi)
    return {"encoding": enc, "chars": length, "sms": segments, "unicode_chars": forcing}


# ===========================================================================
# Storage (SQLite)
# ===========================================================================
DEFAULT_SETTINGS = {
    "company_name": "Your Company Name",
    "contact_line": "Accounts 90000 00000",
    "sender_mode": "dry_run",
    "restsms_url": "http://192.168.1.50:8080/send",
    "restsms_token": "",
    "delay_seconds": "3",
    "max_per_run": "25",
    "wa_phone_number_id": "",
    "wa_waba_id": "",
    "wa_token": "",
    "wa_template_name": "outstanding_reminder",
    "wa_language": "en",
    "wa_api_version": "v25.0",
    "auto_enabled": "0",
    "auto_day": "0",                 # 0 = Monday
    "auto_time": "10:30",
    "auto_file": "",
    "auto_template": "Standard reminder",
    "auto_max_age_days": "3",
    "auto_source": "upload",         # "upload" = list saved from the app; "path" = a file on this computer
    "auto_resend": "1",              # 1 = send the same list again if it hasn't changed
    "auto_list_date": "",            # date the weekly list was uploaded (used for {date})
    "auto_mapping": "",              # column names chosen at upload
    "auto_paused_until": "",         # no automatic sends on or before this date
    "auto_skip_sunday": "1",         # daily reminders skip Sundays
    "auto_task_kind": "",            # "daily" once the Windows task runs every day
    "auto_last_week": "",
    "auto_last_hash": "",
    "auto_last_result": "",
}
SENDER_MODES = ("dry_run", "whatsapp_api", "android_restsms")


@contextmanager
def db():
    conn = sqlite3.connect(DB_PATH, timeout=10)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def init_db() -> None:
    with db() as c:
        c.executescript("""
            CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT NOT NULL);
            CREATE TABLE IF NOT EXISTS templates (
                id INTEGER PRIMARY KEY, name TEXT UNIQUE NOT NULL, body TEXT NOT NULL, updated_at TEXT);
            CREATE TABLE IF NOT EXISTS campaigns (
                id INTEGER PRIMARY KEY, created_at TEXT NOT NULL, file_name TEXT, template_body TEXT,
                sender_mode TEXT NOT NULL, total INTEGER NOT NULL,
                succeeded INTEGER DEFAULT 0, failed INTEGER DEFAULT 0, finished_at TEXT);
            CREATE TABLE IF NOT EXISTS vendor_rules (
                phone TEXT PRIMARY KEY, name TEXT, freq TEXT NOT NULL DEFAULT 'weekly',
                start_date TEXT NOT NULL DEFAULT '', updated_at TEXT);
            CREATE TABLE IF NOT EXISTS vendor_last (phone TEXT PRIMARY KEY, last_date TEXT NOT NULL);
            CREATE TABLE IF NOT EXISTS messages (
                id INTEGER PRIMARY KEY, campaign_id INTEGER NOT NULL REFERENCES campaigns(id),
                party_name TEXT, phone TEXT, amount REAL, body TEXT, sms_count INTEGER,
                status TEXT, detail TEXT, sent_at TEXT);
        """)
        for k, v in {**DEFAULT_SETTINGS, "secret_key": secrets.token_hex(32)}.items():
            c.execute("INSERT OR IGNORE INTO settings (key, value) VALUES (?, ?)", (k, v))
        if c.execute("SELECT COUNT(*) FROM templates").fetchone()[0] == 0:
            c.execute("INSERT INTO templates (name, body, updated_at) VALUES (?, ?, ?)",
                      ("Standard reminder", DEFAULT_TEMPLATE, stamp()))


def stamp() -> str:
    return now_ist().strftime("%Y-%m-%d %H:%M:%S")


def get_settings(include_secret: bool = False) -> dict:
    with db() as c:
        stored = {r["key"]: r["value"] for r in c.execute("SELECT key, value FROM settings")}
    merged = {**DEFAULT_SETTINGS, **stored}
    if not include_secret:
        merged.pop("secret_key", None)
        merged.pop("password_hash", None)
    return merged


def save_settings(values: dict) -> None:
    with db() as c:
        for k, v in values.items():
            c.execute("INSERT INTO settings (key, value) VALUES (?, ?) "
                      "ON CONFLICT(key) DO UPDATE SET value = excluded.value", (k, str(v)))


def list_templates() -> list[dict]:
    with db() as c:
        return [dict(r) for r in c.execute("SELECT name, body FROM templates ORDER BY name")]


def real_sends_today(file_name: str) -> list[str]:
    today = now_ist().strftime("%Y-%m-%d")
    with db() as c:
        rows = c.execute("SELECT created_at FROM campaigns WHERE file_name = ? AND sender_mode != 'dry_run' "
                         "AND created_at LIKE ?", (file_name, f"{today}%")).fetchall()
    return [r["created_at"][11:16] for r in rows]


# ===========================================================================
# Sending
# ===========================================================================
@dataclass
class SendResult:
    ok: bool
    status: str      # dry_run | handed_to_phone | failed
    detail: str


class DryRunSender:
    mode, label, pause, channel = "dry_run", "Dry run (nothing is sent)", False, "none"

    def send(self, phone: str, message: str, values: dict | None = None) -> SendResult:
        return SendResult(True, "dry_run", "Not sent: dry run mode")


class RestSMSSender:
    """POSTs phoneno + message (+ token) to RestSMS on the phone. 'success' means the phone's SMS app
    accepted it, not that it was delivered."""
    mode, label, pause, channel = "android_restsms", "SMS through the Android phone (RestSMS)", True, "sms"

    def __init__(self, url: str, token: str = "", timeout: float = 15):
        self.url, self.token, self.timeout = url.strip(), token.strip(), timeout

    def send(self, phone: str, message: str, values: dict | None = None) -> SendResult:
        data = {"phoneno": f"+91{phone}", "message": message}
        if self.token:
            data["token"] = self.token
        try:
            r = requests.post(self.url, data=data, timeout=self.timeout)
        except requests.RequestException as exc:
            return SendResult(False, "failed", f"Phone not reachable ({exc.__class__.__name__}). "
                                               "Check it's on the same Wi-Fi and RestSMS is running.")
        try:
            payload = r.json()
        except ValueError:
            return SendResult(False, "failed", f"Unexpected reply from phone (HTTP {r.status_code})")
        if payload.get("success"):
            return SendResult(True, "handed_to_phone", "Delivery not confirmed.")
        return SendResult(False, "failed", str(payload.get("message", "Phone rejected the message")))


WA_ERROR_HINTS = {
    190: "The access token is invalid or has expired. Create a permanent (system user) token and paste it in Settings.",
    132000: "The number of variables doesn't match the approved template. Make the message match the approved template.",
    132001: "No approved template with this name and language. Check the template name and language code in Settings.",
    131026: "WhatsApp couldn't deliver to this number (often: the number isn't on WhatsApp).",
    130429: "WhatsApp's sending speed limit was hit. Increase 'Seconds between messages' in Settings.",
    131056: "Too many messages to this same number in a short time.",
    131048: "WhatsApp limited sending because of spam reports. Check the number's quality in WhatsApp Manager.",
    100: "WhatsApp rejected the request (invalid parameter). Check the Phone number ID and template settings.",
}


def wa_error_text(payload: dict, http_status: int) -> str:
    err = (payload or {}).get("error") or {}
    code = err.get("code")
    details = (err.get("error_data") or {}).get("details") or err.get("message") or f"HTTP {http_status}"
    hint = WA_ERROR_HINTS.get(code)
    return f"{hint} ({details}, code {code})" if hint else f"{details} (code {code})"


class WhatsAppSender:
    """Sends an approved template through Meta's WhatsApp Cloud API.
    Success means WhatsApp accepted the message (it returns a message id). Delivery and read
    status are visible in WhatsApp Manager; this app does not receive delivery webhooks."""
    mode, label, pause, channel = "whatsapp_api", "WhatsApp (official Business API)", True, "whatsapp"

    def __init__(self, s: dict, order: list[str], timeout: float = 20):
        self.phone_id = s.get("wa_phone_number_id", "").strip()
        self.token = s.get("wa_token", "").strip()
        self.template = s.get("wa_template_name", "").strip()
        self.language = s.get("wa_language", "en").strip() or "en"
        self.version = s.get("wa_api_version", "v25.0").strip() or "v25.0"
        self.order = order
        self.timeout = timeout

    def payload(self, phone: str, values: dict) -> dict:
        tpl = {"name": self.template, "language": {"code": self.language}}
        if self.order:
            tpl["components"] = [{"type": "body", "parameters": [
                {"type": "text", "text": str(values[k])} for k in self.order]}]
        return {"messaging_product": "whatsapp", "recipient_type": "individual",
                "to": f"91{phone}", "type": "template", "template": tpl}

    def send(self, phone: str, message: str, values: dict | None = None) -> SendResult:
        url = f"{WA_API_BASE}/{self.version}/{self.phone_id}/messages"
        try:
            r = requests.post(url, json=self.payload(phone, values or {}), timeout=self.timeout,
                              headers={"Authorization": f"Bearer {self.token}"})
        except requests.RequestException as exc:
            return SendResult(False, "failed", f"Couldn't reach WhatsApp ({exc.__class__.__name__}). "
                                               "Check this computer's internet connection.")
        try:
            data = r.json()
        except ValueError:
            return SendResult(False, "failed", f"Unexpected reply from WhatsApp (HTTP {r.status_code})")
        if r.ok and data.get("messages"):
            return SendResult(True, "accepted_by_whatsapp", f"Message id {data['messages'][0].get('id', '')}")
        return SendResult(False, "failed", wa_error_text(data, r.status_code))


def wa_missing(s: dict) -> list[str]:
    need = {"wa_phone_number_id": "Phone number ID", "wa_token": "Access token", "wa_template_name": "Template name"}
    return [label for key, label in need.items() if not s.get(key, "").strip()]


def get_sender(settings: dict, order: list[str] | None = None):
    mode = settings.get("sender_mode")
    if mode == RestSMSSender.mode:
        return RestSMSSender(settings.get("restsms_url", ""), settings.get("restsms_token", ""))
    if mode == WhatsAppSender.mode:
        return WhatsAppSender(settings, order or [])
    return DryRunSender()


# Background send jobs (so the page can show progress)
JOBS: dict[str, dict] = {}
JOBS_LOCK = threading.Lock()


def run_job(job_id: str, campaign_id: int, messages: list[dict], sender, delay: float) -> None:
    job = JOBS[job_id]
    for i, m in enumerate(messages):
        result = sender.send(m["phone"], m["message"], m["values"])
        with db() as c:
            c.execute("INSERT INTO messages (campaign_id, party_name, phone, amount, body, sms_count, status, "
                      "detail, sent_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                      (campaign_id, m["name"], m["phone"], m["amount"], m["message"], m["sms"],
                       result.status, result.detail, stamp()))
            col = "succeeded" if result.ok else "failed"
            c.execute(f"UPDATE campaigns SET {col} = {col} + 1 WHERE id = ?", (campaign_id,))
        with JOBS_LOCK:
            job["done"] += 1
            job["ok" if result.ok else "failed"] += 1
        if sender.pause and i < len(messages) - 1:
            time.sleep(delay)
    with db() as c:
        c.execute("UPDATE campaigns SET finished_at = ? WHERE id = ?", (stamp(), campaign_id))
    with JOBS_LOCK:
        job["finished"] = True


def job_running() -> bool:
    with JOBS_LOCK:
        return any(not j["finished"] for j in JOBS.values())


# ===========================================================================
# Web app
# ===========================================================================
app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024
UPLOADS: dict[str, dict] = {}      # upload_id -> {file_name, headers, rows}


def login_required(view):
    @wraps(view)
    def wrapper(*args, **kwargs):
        if not session.get("auth"):
            if request.path.startswith("/api/"):
                return jsonify(error="Logged out"), 401
            return redirect("/login")
        return view(*args, **kwargs)
    return wrapper


def fail(message: str, code: int = 400):
    return jsonify(error=message), code


def password_ok(given: str) -> bool:
    stored = get_settings(include_secret=True).get("password_hash", "")
    if stored:
        return check_password_hash(stored, given)
    return hmac.compare_digest(given, PASSWORD)


def using_default_password() -> bool:
    return not get_settings(include_secret=True).get("password_hash") and PASSWORD == "change-me"


@app.route("/login", methods=["GET", "POST"])
def login():
    error = ""
    if request.method == "POST":
        if password_ok(request.form.get("password", "")):
            session["auth"] = True
            return redirect("/")
        error = "Wrong password."
    note = ("First time? The password is change-me. Change it in Settings after logging in."
            if using_default_password() else "")
    html = LOGIN_HTML.replace("{{error}}", error).replace("{{note}}", note)
    return Response(html, mimetype="text/html")


@app.route("/favicon.svg")
@app.route("/favicon.ico")
def favicon():
    svg = ('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32"><rect width="32" height="32" rx="7" '
           'fill="#0E5E52"/><path d="M8 10h16v10H14l-4 4v-4H8z" fill="#fff"/></svg>')
    return Response(svg, mimetype="image/svg+xml")


@app.route("/logout")
def logout():
    session.clear()
    return redirect("/login")


@app.route("/")
@login_required
def index():
    return Response(PAGE_HTML, mimetype="text/html")


@app.route("/sample.xlsx")
@login_required
def sample():
    wb = Workbook()
    ws = wb.active
    ws.title = "Outstanding"
    rows = [("Vendor Name", "Mobile Number", "Outstanding"),
            ("ABC Traders", 9000000001, 25000), ("XYZ Enterprises", "+91 90000 00002", 42500),
            ("PQR Suppliers", "09000000003", "18,000"), ("LMN Industries", 9000000004, 7250.5),
            ("ABC Traders", 9000000001, 5000), ("Sun Packaging", "033-2222-3333", 11000),
            ("Om Plastics", None, 12000), ("Delta Supplies", 9000000005, 0),
            ("Gupta Stores", 9000000006, 9800)]
    for r in rows:
        ws.append(r)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for col, width in zip("ABC", (20, 18, 14)):
        ws.column_dimensions[col].width = width
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="sample_outstanding.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def mapping_from(payload: dict, upload: dict) -> tuple[int, int, int]:
    m = payload.get("mapping") or {}
    idx = (int(m.get("name", -1)), int(m.get("phone", -1)), int(m.get("amount", -1)))
    if any(i < 0 or i >= len(upload["headers"]) for i in idx):
        raise ValueError("Choose the name, mobile and amount columns.")
    if len(set(idx)) < 3:
        raise ValueError("Choose three different columns.")
    return idx


@app.post("/api/upload")
@login_required
def api_upload():
    f = request.files.get("file")
    if not f or not f.filename:
        return fail("No file received.")
    raw = f.read()
    try:
        headers, rows = read_sheet(f.filename, raw)
    except ValueError as exc:
        return fail(str(exc))
    upload_id = uuid.uuid4().hex
    UPLOADS[upload_id] = {"file_name": f.filename, "headers": headers, "rows": rows, "raw": raw}
    while len(UPLOADS) > 20:
        UPLOADS.pop(next(iter(UPLOADS)))
    guess = [detect_column(headers, h) for h in (NAME_HINTS, PHONE_HINTS, AMOUNT_HINTS)]
    detected = None not in guess and len(set(guess)) == 3
    fallback = iter(i for i in range(len(headers)) if i not in guess)
    guess = [g if g is not None else next(fallback, 0) for g in guess]
    mapping = dict(zip(("name", "phone", "amount"), guess))
    result = validate(rows, *guess) if len(set(guess)) == 3 else None
    return jsonify(upload_id=upload_id, file_name=f.filename, headers=headers, mapping=mapping,
                   validation=result, detected=detected)


@app.post("/api/validate")
@login_required
def api_validate():
    payload = request.get_json(force=True)
    upload = UPLOADS.get(payload.get("upload_id", ""))
    if not upload:
        return fail("The uploaded file has expired. Upload it again.")
    try:
        idx = mapping_from(payload, upload)
    except ValueError as exc:
        return fail(str(exc))
    return jsonify(validation=validate(upload["rows"], *idx))


def build_preview(payload: dict) -> dict:
    """Everything steps 2 and 3 need. Used for display AND recomputed when sending."""
    s = get_settings()
    template = payload.get("template") or ""
    try:
        as_of = date.fromisoformat(payload.get("as_of") or "")
    except ValueError:
        as_of = now_ist().date()
    mode = s["sender_mode"]
    live = mode != DryRunSender.mode
    channel = {"whatsapp_api": "whatsapp", "android_restsms": "sms"}.get(mode, "none")
    out = {"unknown": unknown_placeholders(template), "live": live, "mode": mode, "channel": channel,
           "mode_label": get_sender(s).label, "example": None, "messages": None, "warnings": [], "blockers": [],
           "meta_template": None}
    kw = {"as_of": as_of, "company": s["company_name"], "contact": s["contact_line"]}
    if template.strip() and not out["unknown"]:
        text = render(template, name="Sample Enterprises Pvt Ltd", amount=125000, **kw)
        out["example"] = {"text": text, **sms_stats(text)}
        out["meta_template"] = meta_template(template, s["company_name"], s["contact_line"])

    upload = UPLOADS.get(payload.get("upload_id") or "")
    if not upload or not out["example"]:
        return out
    try:
        idx = mapping_from(payload, upload)
    except ValueError as exc:
        out["blockers"].append(str(exc))
        return out
    ready = validate(upload["rows"], *idx)["ready"]
    msgs = []
    date_text = as_of.strftime("%d-%b-%Y")
    for r in ready:
        text = render(template, name=r["name"], amount=r["amount"], **kw)
        values = {"name": r["name"], "amount": format_inr(r["amount"]), "date": date_text}
        msgs.append({**r, "message": text, "values": values, **sms_stats(text)})
    out["messages"] = msgs
    out["file_name"] = upload["file_name"]
    out["totals"] = {"recipients": len(msgs), "sms": sum(m["sms"] for m in msgs),
                     "amount": format_inr(sum(m["amount"] for m in msgs))}

    max_run = int(s["max_per_run"])
    if not msgs:
        out["blockers"].append("No valid rows to send. Fix the file in step 1.")
    if len(msgs) > max_run:
        out["blockers"].append(f"The list has {len(msgs)} recipients but the limit per run is {max_run}. "
                               "Raise it in Settings once smaller tests work.")
    if s["company_name"] == DEFAULT_SETTINGS["company_name"]:
        out["warnings"].append("Company name is still the placeholder. Set it in Settings.")
    if channel == "sms":
        unicode_count = sum(m["encoding"] == "Unicode" for m in msgs)
        if unicode_count:
            out["warnings"].append(f"{unicode_count} message(s) use Unicode (a ₹ sign or non-English text), "
                                   "so they split into more SMS.")
        long_count = sum(m["sms"] > 1 for m in msgs)
        if long_count:
            out["warnings"].append(f"{long_count} message(s) are longer than one SMS (usually long names).")
    if channel == "whatsapp":
        missing = wa_missing(s)
        if missing:
            out["blockers"].append("WhatsApp isn't set up yet. Fill in " + ", ".join(missing) + " in Settings.")
        out["blockers"].extend(out["meta_template"]["warnings"])
        out["warnings"].append(f"Vendors receive the approved template '{s['wa_template_name']}' with their "
                               "name, amount and date filled in. If you changed the wording here, use "
                               "Settings > Check WhatsApp setup to make sure it still matches.")
    if live:
        earlier = real_sends_today(upload["file_name"])
        if earlier:
            out["warnings"].append(f"This file was already sent today at {', '.join(earlier)}. "
                                   "Sending again means duplicate reminders.")
        if not 9 <= now_ist().hour < 19:
            out["warnings"].append("It's outside business hours (9 AM to 7 PM).")
    return out


@app.post("/api/preview")
@login_required
def api_preview():
    return jsonify(build_preview(request.get_json(force=True)))


@app.post("/api/test-send")
@login_required
def api_test_send():
    payload = request.get_json(force=True)
    phone, reason = normalize_mobile(payload.get("phone"))
    if phone is None:
        return fail(reason)
    preview = build_preview(payload)
    if preview["unknown"] or not preview["example"]:
        return fail("Write a valid message first.")
    if preview["channel"] == "whatsapp" and wa_missing(get_settings()):
        return fail("WhatsApp isn't set up yet. Fill in " + ", ".join(wa_missing(get_settings())) + " in Settings.")
    if preview.get("messages"):
        text, values = preview["messages"][0]["message"], preview["messages"][0]["values"]
    else:
        text = preview["example"]["text"]
        values = {"name": "Sample Enterprises Pvt Ltd", "amount": format_inr(125000),
                  "date": now_ist().strftime("%d-%b-%Y")}
    result = get_sender(get_settings(), preview["meta_template"]["order"]).send(phone, text, values)
    return jsonify(ok=result.ok, status=result.status, detail=result.detail)


@app.post("/api/send")
@login_required
def api_send():
    payload = request.get_json(force=True)
    if not payload.get("confirmed"):
        return fail("Tick 'I've checked the preview' first.")
    if job_running():
        return fail("A send is already running. Wait for it to finish.", 409)
    preview = build_preview(payload)
    if preview["unknown"] or not preview["messages"] or preview["blockers"]:
        return fail(" ".join(preview["blockers"]) or "Nothing to send.")
    s = get_settings()
    sender = get_sender(s, preview["meta_template"]["order"])
    msgs = preview["messages"]
    with db() as c:
        cur = c.execute("INSERT INTO campaigns (created_at, file_name, template_body, sender_mode, total) "
                        "VALUES (?, ?, ?, ?, ?)",
                        (stamp(), preview["file_name"], payload.get("template", ""), sender.mode, len(msgs)))
        campaign_id = cur.lastrowid
    job_id = uuid.uuid4().hex
    with JOBS_LOCK:
        JOBS[job_id] = {"total": len(msgs), "done": 0, "ok": 0, "failed": 0, "finished": False,
                        "campaign_id": campaign_id, "mode": sender.mode}
    threading.Thread(target=run_job, args=(job_id, campaign_id, msgs, sender, float(s["delay_seconds"])),
                     daemon=True).start()
    return jsonify(job_id=job_id)


@app.get("/api/jobs/<job_id>")
@login_required
def api_job(job_id):
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        return jsonify(job) if job else fail("Unknown job", 404)


@app.get("/api/templates")
@login_required
def api_templates():
    return jsonify(templates=list_templates())


@app.post("/api/templates")
@login_required
def api_save_template():
    payload = request.get_json(force=True)
    name, body = (payload.get("name") or "").strip(), payload.get("body") or ""
    if not name or not body.strip():
        return fail("Give the template a name and some text.")
    if unknown_placeholders(body):
        return fail("Fix the placeholders before saving.")
    with db() as c:
        c.execute("INSERT INTO templates (name, body, updated_at) VALUES (?, ?, ?) ON CONFLICT(name) "
                  "DO UPDATE SET body = excluded.body, updated_at = excluded.updated_at", (name, body, stamp()))
    return jsonify(templates=list_templates())


@app.delete("/api/templates/<path:name>")
@login_required
def api_delete_template(name):
    with db() as c:
        c.execute("DELETE FROM templates WHERE name = ?", (name,))
    return jsonify(templates=list_templates())


@app.get("/api/settings")
@login_required
def api_get_settings():
    s = get_settings()
    # Secrets never go back to the browser: only whether they are set.
    s["wa_token_set"], s["restsms_token_set"] = bool(s["wa_token"]), bool(s["restsms_token"])
    s["wa_token"], s["restsms_token"] = "", ""
    return jsonify(settings=s)


@app.post("/api/settings")
@login_required
def api_save_settings():
    current = get_settings()
    old_tokens = {"wa_token": current["wa_token"], "restsms_token": current["restsms_token"]}
    p = request.get_json(force=True)
    try:
        delay = min(max(float(p.get("delay_seconds", 3)), 0.5), 120)
        max_run = min(max(int(p.get("max_per_run", 25)), 1), 2000)
    except (TypeError, ValueError):
        return fail("Delay and limit must be numbers.")
    mode = p.get("sender_mode")
    if mode not in SENDER_MODES:
        return fail("Unknown sending mode.")
    url = (p.get("restsms_url") or "").strip()
    if mode == RestSMSSender.mode and not re.match(r"^https?://", url):
        return fail("Enter the phone's RestSMS address, starting with http://")
    save_settings({"company_name": (p.get("company_name") or "").strip() or DEFAULT_SETTINGS["company_name"],
                   "contact_line": (p.get("contact_line") or "").strip(),
                   "sender_mode": mode, "restsms_url": url, "restsms_token": (p.get("restsms_token") or "").strip() or old_tokens["restsms_token"],
                   "delay_seconds": delay, "max_per_run": max_run,
                   "wa_phone_number_id": re.sub(r"\s", "", p.get("wa_phone_number_id") or ""),
                   "wa_waba_id": re.sub(r"\s", "", p.get("wa_waba_id") or ""),
                   "wa_token": (p.get("wa_token") or "").strip() or old_tokens["wa_token"],
                   "wa_template_name": (p.get("wa_template_name") or "").strip(),
                   "wa_language": (p.get("wa_language") or "en").strip() or "en",
                   "wa_api_version": (p.get("wa_api_version") or "v25.0").strip() or "v25.0"})
    return jsonify(settings=get_settings())


@app.post("/api/password")
@login_required
def api_password():
    p = request.get_json(force=True)
    if not password_ok(p.get("current", "")):
        return fail("The current password is wrong.")
    new = p.get("new", "")
    if len(new) < 6:
        return fail("Use at least 6 characters for the new password.")
    save_settings({"password_hash": generate_password_hash(new)})
    return jsonify(ok=True)


@app.get("/api/status")
@login_required
def api_status():
    return jsonify(default_password=using_default_password())


def wa_check(template_text: str) -> dict:
    """Checks the saved WhatsApp settings against Meta: number, token, and the approved template."""
    s = get_settings()
    missing = wa_missing(s)
    if missing:
        return {"ok": False, "checks": [{"ok": False, "label": "Settings", "detail": "Fill in " + ", ".join(missing) + " and save."}]}
    base = f"{WA_API_BASE}/{s['wa_api_version']}"
    auth = {"Authorization": f"Bearer {s['wa_token']}"}
    checks = []
    try:
        r = requests.get(f"{base}/{s['wa_phone_number_id']}", headers=auth, timeout=15,
                         params={"fields": "display_phone_number,verified_name,quality_rating"})
        d = r.json()
        if r.ok and d.get("display_phone_number"):
            checks.append({"ok": True, "label": "Number and token",
                           "detail": f"{d.get('verified_name', '')} ({d['display_phone_number']}), quality {d.get('quality_rating', 'unknown')}"})
        else:
            checks.append({"ok": False, "label": "Number and token", "detail": wa_error_text(d, r.status_code)})
            return {"ok": False, "checks": checks}
    except (requests.RequestException, ValueError) as exc:
        checks.append({"ok": False, "label": "Number and token", "detail": f"Couldn't reach WhatsApp ({exc.__class__.__name__})."})
        return {"ok": False, "checks": checks}

    if not s["wa_waba_id"]:
        checks.append({"ok": None, "label": "Template",
                       "detail": "Add the WhatsApp Business Account ID to also check the template's approval and wording."})
        return {"ok": True, "checks": checks}
    try:
        r = requests.get(f"{base}/{s['wa_waba_id']}/message_templates", headers=auth, timeout=15,
                         params={"name": s["wa_template_name"], "fields": "name,status,language,category,components"})
        d = r.json()
    except (requests.RequestException, ValueError) as exc:
        checks.append({"ok": False, "label": "Template", "detail": f"Couldn't reach WhatsApp ({exc.__class__.__name__})."})
        return {"ok": False, "checks": checks}
    if not r.ok:
        checks.append({"ok": False, "label": "Template", "detail": wa_error_text(d, r.status_code)})
        return {"ok": False, "checks": checks}
    match = [t for t in d.get("data", []) if t.get("name") == s["wa_template_name"]
             and t.get("language", "").lower() == s["wa_language"].lower()]
    if not match:
        langs = ", ".join(sorted({t.get("language", "") for t in d.get("data", []) if t.get("name") == s["wa_template_name"]}))
        checks.append({"ok": False, "label": "Template",
                       "detail": f"No template '{s['wa_template_name']}' in language '{s['wa_language']}'."
                                 + (f" It exists in: {langs}." if langs else " Create it in WhatsApp Manager.")})
        return {"ok": False, "checks": checks}
    t = match[0]
    status = t.get("status", "UNKNOWN")
    checks.append({"ok": status == "APPROVED", "label": "Template status",
                   "detail": f"{status} (category {t.get('category', '?')})" +
                             ("" if status == "APPROVED" else ". It can only be sent once approved.")})
    body = next((c.get("text", "") for c in t.get("components", []) if str(c.get("type", "")).upper() == "BODY"), "")
    current = meta_template(template_text or "", s["company_name"], s["contact_line"])
    same = normalise_text(body) == normalise_text(current["text"])
    checks.append({"ok": same, "label": "Wording",
                   "detail": "The approved template matches the message in the app." if same else
                             "The approved template is different from the message in the app. Vendors will get the "
                             "approved wording: " + body})
    return {"ok": all(c["ok"] for c in checks if c["ok"] is not None), "checks": checks}


@app.post("/api/wa/check")
@login_required
def api_wa_check():
    payload = request.get_json(silent=True) or {}
    return jsonify(**wa_check(payload.get("template") or ""))


@app.get("/api/history")
@login_required
def api_history():
    with db() as c:
        rows = [dict(r) for r in c.execute("SELECT id, created_at, file_name, sender_mode, total, succeeded, "
                                           "failed, finished_at FROM campaigns ORDER BY id DESC LIMIT 200")]
    return jsonify(runs=rows)


def run_messages(run_id: int) -> list[dict]:
    with db() as c:
        return [dict(r) for r in c.execute(
            "SELECT party_name, phone, amount, sms_count, status, detail, sent_at, body FROM messages "
            "WHERE campaign_id = ? ORDER BY id", (run_id,))]


@app.get("/api/history/<int:run_id>")
@login_required
def api_run(run_id):
    return jsonify(messages=run_messages(run_id))


@app.get("/api/history/<int:run_id>.csv")
@login_required
def api_run_csv(run_id):
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["Party", "Mobile", "Amount", "SMS", "Status", "Detail", "Time", "Message"])
    labels = {"dry_run": "Dry run", "handed_to_phone": "Accepted by phone",
              "accepted_by_whatsapp": "Accepted by WhatsApp", "failed": "Failed"}
    for m in run_messages(run_id):
        w.writerow([m["party_name"], m["phone"], m["amount"], m["sms_count"], labels.get(m["status"], m["status"]),
                    m["detail"], m["sent_at"], m["body"]])
    return Response("\ufeff" + buf.getvalue(), mimetype="text/csv",
                    headers={"Content-Disposition": f"attachment; filename=sms_run_{run_id}.csv"})


# ===========================================================================
# Automatic weekly sending
# ===========================================================================
DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
TASK_NAME = "Outstanding Reminders - weekly send"
LOG_PATH = APP_DIR / "auto-send.log"


class AutoStop(Exception):
    """Stops an automatic run; the message says why."""


def auto_log(line: str) -> None:
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(f"{stamp()}  {line}\n")
    except OSError:
        pass


def clean_path(raw: str) -> str:
    return (raw or "").strip().strip('"').strip("'").strip()



FREQ_LABELS = {"weekly": "Every week", "monthly": "Every month", "daily": "Every day", "off": "Don't send"}
FREQ_HINTS = ["frequency", "how often", "reminder frequency", "remind every", "reminder"]
START_HINTS = ["start date", "reminder date", "remind from", "from date", "start from", "start"]


def parse_date_cell(value) -> date | None:
    """Dates from Excel: real date cells, Excel serial numbers, or text like 05-11-2026 / 5/11/2026 / 2026-11-05 / 05-Nov-2026."""
    if is_blank(value):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and 20000 < float(value) < 80000:
        return date(1899, 12, 30) + timedelta(days=int(value))
    text = str(value).strip()
    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%d.%m.%Y", "%Y-%m-%d", "%d-%b-%Y", "%d %b %Y", "%d-%m-%y", "%d/%m/%y",
                "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    return None


def parse_freq(value) -> str | None:
    text = str(value or "").strip().lower()
    if not text:
        return None
    if any(w in text for w in ("don", "stop", "off", "no ", "never")) or text in ("no", "none"):
        return "off"
    if "day" in text or "daily" in text:
        return "daily"
    if "month" in text:
        return "monthly"
    if "week" in text:
        return "weekly"
    return None


def get_rules() -> dict:
    with db() as c:
        return {r["phone"]: {"freq": r["freq"], "start": r["start_date"]} for r in c.execute("SELECT * FROM vendor_rules")}


def set_rule(phone: str, name: str, freq: str, start: str) -> None:
    with db() as c:
        c.execute("INSERT INTO vendor_rules (phone, name, freq, start_date, updated_at) VALUES (?, ?, ?, ?, ?) "
                  "ON CONFLICT(phone) DO UPDATE SET name = excluded.name, freq = excluded.freq, "
                  "start_date = excluded.start_date, updated_at = excluded.updated_at",
                  (phone, name, freq, start, stamp()))


def get_last() -> dict:
    with db() as c:
        return {r["phone"]: date.fromisoformat(r["last_date"]) for r in c.execute("SELECT * FROM vendor_last")}


def mark_sent(phone: str, day: date) -> None:
    with db() as c:
        c.execute("INSERT INTO vendor_last (phone, last_date) VALUES (?, ?) "
                  "ON CONFLICT(phone) DO UPDATE SET last_date = excluded.last_date", (phone, day.isoformat()))


def _month_day(year: int, month: int, day: int) -> date:
    nxt = date(year + (month == 12), month % 12 + 1, 1)
    return date(year, month, min(day, (nxt - timedelta(days=1)).day))


def due_on(day: date, rule: dict | None, last: date | None, s: dict, paused: date | None, catch_up: bool = True) -> bool:
    """Is a reminder due for this vendor on this day?
    weekly: on the weekly day (catch-up next day); monthly: on the start date's day of month (catch-up 2 days);
    daily: every day (Sundays skipped if set). Nothing before the start date or during a pause."""
    rule = rule or {}
    freq = rule.get("freq") or "weekly"
    start = date.fromisoformat(rule["start"]) if rule.get("start") else None
    if freq == "off" or (start and day < start) or (paused and day <= paused):
        return False
    if freq == "daily":
        if s.get("auto_skip_sunday", "1") == "1" and day.weekday() == 6:
            return False
        return last != day
    if freq == "monthly":
        if s.get("auto_skip_sunday", "1") == "1" and day.weekday() == 6:
            return False                                      # a Sunday date goes out on Monday instead
        slot = _month_day(day.year, day.month, start.day if start else 1)
        allowed = 2 if catch_up else 0
        if s.get("auto_skip_sunday", "1") == "1" and slot.weekday() == 6:
            allowed = max(allowed, 1)
    else:                                                     # weekly
        slot = day - timedelta(days=(day.weekday() - int(s["auto_day"])) % 7)
        allowed = 1 if catch_up else 0
    if day < slot or (day - slot).days > allowed:
        return False
    if (start and slot < start) or (paused and slot <= paused):
        return False
    return not (last and last >= slot)


def first_run_day(s: dict, now: datetime) -> date:
    hh, mm = (int(x) for x in (s["auto_time"] or "10:30").split(":"))
    return now.date() if (now.hour, now.minute) < (hh, mm) else now.date() + timedelta(days=1)


def next_due(rule: dict | None, last: date | None, s: dict, paused: date | None, now: datetime) -> date | None:
    day = first_run_day(s, now)
    for _ in range(400):
        if due_on(day, rule, last, s, paused, catch_up=False):
            return day
        day += timedelta(days=1)
    return None


def load_weekly_list(s: dict) -> dict:
    """Reads the saved weekly list and returns everything needed to send from it. Raises AutoStop on problems."""
    raw_path = clean_path(s["auto_file"])
    if not raw_path:
        raise AutoStop("No vendor list is saved for automatic sending.")
    path = Path(raw_path)
    if not path.is_file():
        raise AutoStop(f"The saved vendor list is missing: {raw_path}")
    try:
        data = path.read_bytes()
    except OSError as exc:
        raise AutoStop(f"Couldn't open the vendor list ({exc.__class__.__name__}). Close it in Excel and try again.")
    try:
        headers, rows = read_sheet(path.name, data)
    except ValueError as exc:
        raise AutoStop(str(exc))
    idx = [detect_column(headers, h) for h in (NAME_HINTS, PHONE_HINTS, AMOUNT_HINTS)]
    try:
        chosen = json.loads(s["auto_mapping"] or "[]")
        if len(chosen) == 3 and all(c in headers for c in chosen):
            idx = [headers.index(c) for c in chosen]
    except ValueError:
        pass
    if None in idx or len(set(idx)) < 3:
        raise AutoStop("Couldn't find the name, mobile and amount columns. Use headings like "
                       "'Vendor Name', 'Mobile Number' and 'Outstanding'.")
    modified = datetime.fromtimestamp(path.stat().st_mtime, IST)
    as_of = s["auto_list_date"] if s["auto_source"] == "upload" and s["auto_list_date"] else modified.date().isoformat()
    list_name = s.get("auto_list_name") if s["auto_source"] == "upload" and s.get("auto_list_name") else path.name
    return {"path": path, "data": data, "headers": headers, "rows": rows, "idx": idx, "modified": modified,
            "as_of": as_of, "list_name": list_name, "digest": hashlib.sha256(data).hexdigest(),
            "validation": validate(rows, *idx)}


def rules_from_excel(headers: list, rows: list, idx: list) -> tuple[int, list]:
    """Optional 'Frequency' and 'Start date' columns in the Excel. Returns (rules saved, problems)."""
    taken = set(idx)

    def find(hints, skip_idx, skip_words=()):
        low = [h.strip().lower() for h in headers]
        ok = [i for i in range(len(low)) if i not in skip_idx and not any(w in low[i] for w in skip_words)]
        for h in hints:
            for i in ok:
                if low[i] == h:
                    return i
        for h in hints:
            for i in ok:
                if h in low[i]:
                    return i
        return None

    scol = find(START_HINTS, taken)
    fcol = find(FREQ_HINTS, taken | ({scol} if scol is not None else set()), ("date", "from", "start"))
    if fcol is None and scol is None:
        return 0, []
    saved, problems, seen = 0, [], set()
    existing = get_rules()
    for row in rows:
        v = row["values"]
        phone, _ = normalize_mobile(v[idx[1]])
        if not phone or phone in seen:
            continue
        seen.add(phone)
        freq = parse_freq(v[fcol]) if fcol is not None else None
        if fcol is not None and not is_blank(v[fcol]) and freq is None:
            problems.append(f"Row {row['excel_row']}: frequency '{show_value(v[fcol])}' not understood (use Weekly, Monthly, Daily or Don't send)")
        start = parse_date_cell(v[scol]) if scol is not None else None
        if scol is not None and not is_blank(v[scol]) and start is None:
            problems.append(f"Row {row['excel_row']}: date '{show_value(v[scol])}' not understood (use e.g. 05-11-2026)")
        if freq is None and start is None:
            continue                                      # blank cells: keep whatever was set in the app
        old = existing.get(phone) or {}
        set_rule(phone, show_value(v[idx[0]]), freq or old.get("freq") or "weekly",
                 start.isoformat() if start else old.get("start", ""))
        saved += 1
    return saved, problems


def run_auto(test: bool = False) -> dict:
    """The daily automatic check: sends to every vendor whose reminder is due today.
    With test=True nothing is sent and every check is reported."""
    s = get_settings()
    report = {"status": "", "reason": "", "steps": [], "would_skip": [], "preview": [], "count": 0, "due": 0}

    def step(ok, label, detail):
        report["steps"].append({"ok": ok, "label": label, "detail": detail})

    def guard(ok: bool, label: str, good: str, bad: str):
        step(ok, label, good if ok else bad)
        if not ok:
            if not test:
                raise AutoStop(bad)
            report["would_skip"].append(bad)

    def finish(status: str, reason: str, remember: bool = True) -> dict:
        report["status"], report["reason"] = status, reason
        if not test:
            if remember:
                save_settings({"auto_last_result": f"{stamp()[:16]}  {status.upper()}: {reason}"})
            auto_log(f"{status.upper()}: {reason}")
        return report

    try:
        now = now_ist()
        today = now.date()
        guard(s["auto_enabled"] == "1", "Turned on", "Automatic sending is on.", "Automatic sending is turned off.")
        paused = date.fromisoformat(s["auto_paused_until"]) if s["auto_paused_until"] else None
        if paused and today <= paused:
            guard(False, "Paused", "", f"Paused. Reminders start again {auto_status(s, now)['next_label']}.")
        if not test:
            guard(9 <= now.hour < 19, "Time", "", "Outside business hours (9 AM to 7 PM).")

        wl = load_weekly_list(s)
        updated = wl["modified"].strftime("%d-%b-%Y %H:%M")
        changed = wl["digest"] != s["auto_last_hash"]
        if s["auto_resend"] == "1":
            step(True, "List", f"{wl['list_name']} (amounts as of {date.fromisoformat(wl['as_of']).strftime('%d-%b-%Y')}).")
        else:
            age_days = (time.time() - wl["path"].stat().st_mtime) / 86400
            max_age = float(s["auto_max_age_days"] or 3)
            guard(age_days <= max_age, "List is recent", f"Last updated {updated}.",
                  f"The list was last updated {updated}, more than {max_age:g} days ago, so the amounts may be old.")
            guard(changed, "List changed", "The list has changed since the last automatic send.",
                  "The list hasn't changed since the last automatic send.")
        v = wl["validation"]
        step(True, "Excel", f"{v['records']} rows: {v['ready_count']} ready, {v['invalid']} invalid, {v['skipped']} with nothing due.")

        tpl = next((t for t in list_templates() if t["name"] == s["auto_template"]), None)
        if not tpl:
            raise AutoStop(f"The saved message '{s['auto_template']}' no longer exists. Pick one in Settings.")
        upload_id = "auto-" + uuid.uuid4().hex
        UPLOADS[upload_id] = {"file_name": wl["path"].name, "headers": wl["headers"], "rows": wl["rows"]}
        preview = build_preview({"upload_id": upload_id, "template": tpl["body"], "as_of": wl["as_of"],
                                 "mapping": dict(zip(("name", "phone", "amount"), wl["idx"]))})
        UPLOADS.pop(upload_id, None)
        if preview["unknown"]:
            raise AutoStop("The saved message has placeholders the app can't fill.")
        msgs = preview["messages"] or []
        if not msgs:
            raise AutoStop("No valid rows in the vendor list.")

        # ----- who is due today
        rules, last = get_rules(), get_last()
        due = [m for m in msgs if due_on(today, rules.get(m["phone"]), last.get(m["phone"]), s, paused)]
        report["count"], report["due"] = len(msgs), len(due)
        report["preview"] = [{"name": m["name"], "phone": m["phone"], "amount": m["amount_text"],
                              "message": m["message"]} for m in due[:3]]
        kinds = {}
        for m in due:
            k = FREQ_LABELS[(rules.get(m["phone"]) or {}).get("freq") or "weekly"].lower()
            kinds[k] = kinds.get(k, 0) + 1
        step(True, "Due today", f"{len(due)} of {len(msgs)} vendors" +
             (" (" + ", ".join(f"{n} {k}" for k, n in kinds.items()) + ")" if kinds else "") + ".")
        if not due:
            nxt = auto_status(s, now)["next_label"]
            return finish("nothing_due", f"No reminders due today. Next: {nxt}.", remember=False)

        cap = int(s["max_per_run"])
        guard(len(due) <= cap, "Limit", f"{len(due)} is within the limit of {cap}.",
              f"{len(due)} reminders are due but the limit per run is {cap}. Raise 'Most messages per run' in Settings.")
        for b in preview["blockers"]:
            if not b.startswith("The list has"):                      # the cap is checked above, on today's count
                guard(False, "Ready to send", "", b)
        if preview["channel"] == "whatsapp":
            chk = wa_check(tpl["body"])
            bad = [f"{c['label']}: {c['detail']}" for c in chk["checks"] if c["ok"] is False]
            guard(not bad, "WhatsApp", "Number, token and approved template are fine.", "WhatsApp check failed. " + " ".join(bad))
        step(True, "Send with", preview["mode_label"])

        if test:
            if report["would_skip"]:
                return finish("would_skip", "A real run now would skip: " + " ".join(report["would_skip"]))
            return finish("ready", f"A real run now would send to {len(due)} vendors due today.")

        # ----- send
        sender = get_sender(s, preview["meta_template"]["order"])
        with db() as c:
            cur = c.execute("INSERT INTO campaigns (created_at, file_name, template_body, sender_mode, total) "
                            "VALUES (?, ?, ?, ?, ?)", (stamp(), f"{wl['list_name']} (automatic)", tpl["body"], sender.mode, len(due)))
            campaign_id = cur.lastrowid
        ok = failed = in_a_row = 0
        delay = float(s["delay_seconds"])
        for i, m in enumerate(due):
            result = sender.send(m["phone"], m["message"], m["values"])
            with db() as c:
                c.execute("INSERT INTO messages (campaign_id, party_name, phone, amount, body, sms_count, status, "
                          "detail, sent_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                          (campaign_id, m["name"], m["phone"], m["amount"], m["message"], m["sms"],
                           result.status, result.detail, stamp()))
                col = "succeeded" if result.ok else "failed"
                c.execute(f"UPDATE campaigns SET {col} = {col} + 1 WHERE id = ?", (campaign_id,))
            if result.ok:
                ok, in_a_row = ok + 1, 0
                mark_sent(m["phone"], today)
            else:
                failed, in_a_row = failed + 1, in_a_row + 1
                if in_a_row >= 5:
                    with db() as c:
                        c.execute("UPDATE campaigns SET finished_at = ? WHERE id = ?", (stamp(), campaign_id))
                    raise AutoStop(f"Stopped after 5 failures in a row ({ok} sent before that). Last error: {result.detail}")
            if sender.pause and i < len(due) - 1:
                time.sleep(delay)
        with db() as c:
            c.execute("UPDATE campaigns SET finished_at = ? WHERE id = ?", (stamp(), campaign_id))
        if ok:
            save_settings({"auto_last_week": now.strftime("%G-W%V"), "auto_last_hash": wl["digest"]})
        return finish("sent", f"{ok} sent, {failed} failed, from {wl['list_name']}. Details in History.")
    except AutoStop as stop:
        return finish("skipped", str(stop))
    except Exception as exc:                      # never crash silently in an unattended run
        return finish("error", f"Unexpected problem: {exc.__class__.__name__}: {exc}")


# ----- Windows Task Scheduler
def _schtasks(*args: str) -> subprocess.CompletedProcess:
    flags = 0x08000000 if os.name == "nt" else 0          # CREATE_NO_WINDOW
    return subprocess.run(["schtasks", *args], capture_output=True, text=True, creationflags=flags)


def task_xml(day: int, hhmm: str) -> str:
    if getattr(sys, "frozen", False):
        command, arguments = sys.executable, "--auto-send"
    else:
        command, arguments = sys.executable, f'"{Path(__file__).resolve()}" --auto-send'
    start = datetime.now().date()
    return f"""<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo><Description>Checks every day and sends the outstanding reminders that are due (Outstanding Reminders app).</Description></RegistrationInfo>
  <Triggers>
    <CalendarTrigger>
      <StartBoundary>{start.isoformat()}T{hhmm}:00</StartBoundary>
      <Enabled>true</Enabled>
      <ScheduleByDay><DaysInterval>1</DaysInterval></ScheduleByDay>
    </CalendarTrigger>
  </Triggers>
  <Principals><Principal id="Author"><LogonType>InteractiveToken</LogonType><RunLevel>LeastPrivilege</RunLevel></Principal></Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>true</RunOnlyIfNetworkAvailable>
    <WakeToRun>true</WakeToRun>
    <ExecutionTimeLimit>PT3H</ExecutionTimeLimit>
    <Enabled>true</Enabled>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>{xml_escape(command)}</Command>
      <Arguments>{xml_escape(arguments)}</Arguments>
      <WorkingDirectory>{xml_escape(str(APP_DIR))}</WorkingDirectory>
    </Exec>
  </Actions>
</Task>
"""


def schedule_task(day: int, hhmm: str) -> tuple[bool, str]:
    if os.name != "nt":
        return False, "The weekly schedule can only be created on Windows."
    with tempfile.NamedTemporaryFile("w", suffix=".xml", delete=False, encoding="utf-16") as f:
        f.write(task_xml(day, hhmm))
        xml_path = f.name
    try:
        r = _schtasks("/Create", "/TN", TASK_NAME, "/XML", xml_path, "/F")
    except OSError as exc:
        return False, f"Couldn't reach Windows Task Scheduler ({exc.__class__.__name__})."
    finally:
        try:
            os.remove(xml_path)
        except OSError:
            pass
    if r.returncode != 0:
        return False, "Windows didn't accept the schedule: " + (r.stderr or r.stdout).strip()
    save_settings({"auto_task_kind": "daily"})
    return True, f"Scheduled: the app checks every day at {hhmm}."


def unschedule_task() -> tuple[bool, str]:
    if os.name != "nt":
        return True, ""
    try:
        r = _schtasks("/Delete", "/TN", TASK_NAME, "/F")
    except OSError:
        return False, "Couldn't reach Windows Task Scheduler."
    return True, "Weekly schedule removed." if r.returncode == 0 else "No weekly schedule was set."


def task_status() -> str:
    if os.name != "nt":
        return "Windows schedule: only available on Windows."
    try:
        r = _schtasks("/Query", "/TN", TASK_NAME, "/FO", "LIST")
    except OSError:
        return "Windows schedule: couldn't check."
    if r.returncode != 0:
        return "Windows schedule: not set."
    info = {}
    for line in r.stdout.splitlines():
        if ":" in line:
            k, _, v = line.partition(":")
            info[k.strip().lower()] = v.strip()
    return f"Windows schedule: set. Next run {info.get('next run time', 'unknown')}."


def auto_status(s: dict | None = None, now: datetime | None = None) -> dict:
    s = s or get_settings()
    now = now or now_ist()
    if s["auto_enabled"] != "1":
        return {"state": "off", "next_label": "", "text": "Automatic sending is OFF."}
    paused = date.fromisoformat(s["auto_paused_until"]) if s["auto_paused_until"] else None
    try:
        wl = load_weekly_list(s)
        phones = [r["phone"] for r in wl["validation"]["ready"]]
    except AutoStop:
        phones = []
    rules, last = get_rules(), get_last()
    days = [d for d in (next_due(rules.get(p), last.get(p), s, paused, now) for p in phones) if d]
    hhmm = s["auto_time"]
    if days:
        first = min(days)
        count = days.count(first)
        label = f"{first.strftime('%a %d-%b-%Y')} at {hhmm} ({count} vendor{'s' if count != 1 else ''})"
    else:
        label = "none scheduled"
    if paused and now.date() <= paused:
        return {"state": "paused", "next_label": label, "paused_until": paused.isoformat(),
                "text": f"Automatic sending is PAUSED. It starts again by itself: next reminders {label}."}
    return {"state": "on", "next_label": label, "text": f"Automatic sending is ON. Next reminders: {label}."}


@app.get("/api/auto/status")
@login_required
def api_auto_status():
    s = get_settings()
    weekly = None
    if s.get("auto_list_name") and s["auto_list_date"] and s["auto_source"] == "upload":
        weekly = f"{s['auto_list_name']}, {s.get('auto_list_count', '')} vendors, amounts as of " \
                 f"{date.fromisoformat(s['auto_list_date']).strftime('%d-%b-%Y')}"
    body = next((t["body"] for t in list_templates() if t["name"] == s["auto_template"]), "")
    last = s["auto_last_result"]
    return jsonify(**auto_status(s), weekly=weekly, configured=bool(s["auto_file"]), last_result=last,
                   template_body=body, live=s["sender_mode"] != "dry_run",
                   list_ok=bool(s["auto_file"]) and Path(s["auto_file"]).is_file())


@app.post("/api/auto/pause")
@login_required
def api_auto_pause():
    p = request.get_json(force=True)
    s = get_settings()
    if s["auto_enabled"] != "1":
        return fail("Automatic sending is off, so there is nothing to pause.")
    now = now_ist()
    if p.get("resume_on"):
        try:
            resume = date.fromisoformat(p["resume_on"])
        except ValueError:
            return fail("Choose a valid date.")
        if resume <= now.date():
            return fail("Choose a date after today.")
        until = resume - timedelta(days=1)
    else:
        try:
            weeks = int(p.get("weeks", p.get("skip", 1)))
            assert 1 <= weeks <= 26
        except (ValueError, AssertionError, TypeError):
            return fail("Choose for how many weeks to pause.")
        until = now.date() + timedelta(days=7 * weeks - 1)
    save_settings({"auto_paused_until": until.isoformat()})
    auto_log(f"PAUSED until {until.isoformat()} (by user)")
    return jsonify(ok=True, **auto_status())


@app.post("/api/auto/resume")
@login_required
def api_auto_resume():
    save_settings({"auto_paused_until": ""})
    auto_log("RESUMED (by user)")
    return jsonify(ok=True, **auto_status())


@app.post("/api/auto/start")
@login_required
def api_auto_start():
    """Turn automatic sending back on with the saved day, time, list and message."""
    s = get_settings()
    path = clean_path(s["auto_file"])
    if not path or not Path(path).is_file():
        return fail("There's no weekly list yet. On Home, upload the Excel and click \"Save as the weekly list\" first.")
    if s["auto_template"] not in [t["name"] for t in list_templates()]:
        return fail("The saved message for automatic sending no longer exists. Choose one in Settings.")
    save_settings({"auto_enabled": "1", "auto_paused_until": ""})
    ok, msg = schedule_task(int(s["auto_day"]), s["auto_time"])
    auto_log("TURNED ON (by user)")
    return jsonify(ok=ok, message="" if ok else "Turned on, but " + msg[0].lower() + msg[1:], **auto_status())


@app.post("/api/auto/stop")
@login_required
def api_auto_stop():
    save_settings({"auto_enabled": "0", "auto_paused_until": ""})
    unschedule_task()
    auto_log("STOPPED (by user)")
    return jsonify(ok=True, **auto_status())


@app.get("/api/auto")
@login_required
def api_auto_get():
    s = get_settings()
    keys = ("auto_enabled", "auto_day", "auto_time", "auto_file", "auto_template", "auto_max_age_days", "auto_last_result",
            "auto_source", "auto_resend", "auto_skip_sunday")
    weekly = None
    if s.get("auto_list_name") and s["auto_list_date"]:
        weekly = {"name": s["auto_list_name"], "count": s.get("auto_list_count", ""),
                  "date": date.fromisoformat(s["auto_list_date"]).strftime("%d-%b-%Y"),
                  "exists": Path(s["auto_file"]).is_file() if s["auto_source"] == "upload" else True}
    return jsonify(auto={k: s[k] for k in keys}, templates=[t["name"] for t in list_templates()], weekly=weekly,
                   task=task_status(), max_per_run=int(s["max_per_run"]), mode=s["sender_mode"], status=auto_status(s))


@app.post("/api/auto")
@login_required
def api_auto_save():
    p = request.get_json(force=True)
    enabled = bool(p.get("enabled"))
    try:
        day = int(p.get("day", 0))
        assert 0 <= day <= 6
        hhmm = (p.get("time") or "10:30").strip()
        hh, mm = (int(x) for x in hhmm.split(":"))
        assert 0 <= hh <= 23 and 0 <= mm <= 59
        hhmm = f"{hh:02d}:{mm:02d}"
        max_age = float(p.get("max_age_days", 3))
        assert 0.5 <= max_age <= 30
    except (ValueError, AssertionError, TypeError):
        return fail("Check the day, time (e.g. 10:30) and the number of days.")
    source = "path" if p.get("source") == "path" else "upload"
    resend = "1" if p.get("resend", True) else "0"
    old = get_settings()
    if source == "upload":
        weekly_files = [APP_DIR / "weekly-list.xlsx", APP_DIR / "weekly-list.csv"]
        path = next((str(f) for f in weekly_files if f.is_file()), "")
    else:
        path = clean_path(p.get("file", ""))
    template = (p.get("template") or "").strip()
    if enabled:
        if source == "upload" and not path:
            return fail("No weekly list yet. Go to Send reminders, upload the Excel, and click "
                        "'Use this list for automatic weekly sending'.")
        if source == "path" and not path:
            return fail("Enter the full path of the Excel file, e.g. C:\\Reminders\\outstanding.xlsx")
        if source == "path" and not Path(path).is_file():
            return fail(f"No file found at: {path}")
        if source == "path" and not path.lower().endswith((".xlsx", ".xlsm", ".csv")):
            return fail("The file must be an .xlsx or .csv file.")
        if template not in [t["name"] for t in list_templates()]:
            return fail("Choose which saved message to send.")
        if not 9 <= hh < 19:
            return fail("Pick a time between 09:00 and 18:59, so reminders go out in business hours.")
    changes = {"auto_enabled": "1" if enabled else "0", "auto_day": day, "auto_time": hhmm, "auto_file": path,
               "auto_template": template or old["auto_template"], "auto_max_age_days": max_age,
               "auto_source": source, "auto_resend": resend,
               "auto_skip_sunday": "1" if p.get("skip_sunday", True) else "0"}
    if not enabled:
        changes["auto_paused_until"] = ""
    if path != old["auto_file"]:
        changes["auto_last_hash"] = ""
    if source == "path" and old["auto_source"] != "path":
        changes["auto_mapping"] = ""            # columns are detected from the file's headings
    save_settings(changes)
    ok, msg = schedule_task(day, hhmm) if enabled else unschedule_task()
    if enabled and not ok:
        return jsonify(ok=False, message="Settings saved, but " + msg[0].lower() + msg[1:], task=task_status())
    return jsonify(ok=True, message=msg or "Automatic sending is off.", task=task_status())


@app.post("/api/auto/use-upload")
@login_required
def api_auto_use_upload():
    payload = request.get_json(force=True)
    upload = UPLOADS.get(payload.get("upload_id") or "")
    if not upload or "raw" not in upload:
        return fail("The uploaded file has expired. Upload it again.")
    try:
        idx = mapping_from(payload, upload)
    except ValueError as exc:
        return fail(str(exc))
    v = validate(upload["rows"], *idx)
    if not v["ready_count"]:
        return fail("This list has no valid rows to send.")
    ext = ".csv" if upload["file_name"].lower().endswith(".csv") else ".xlsx"
    target = APP_DIR / f"weekly-list{ext}"
    for old in (APP_DIR / "weekly-list.xlsx", APP_DIR / "weekly-list.csv"):
        if old != target and old.exists():
            old.unlink()
    target.write_bytes(upload["raw"])
    n_rules, rule_problems = rules_from_excel(upload["headers"], upload["rows"], list(idx))
    today = now_ist().date()
    save_settings({"auto_source": "upload", "auto_file": str(target), "auto_list_date": today.isoformat(),
                   "auto_mapping": json.dumps([upload["headers"][i] for i in idx]), "auto_last_hash": "",
                   "auto_list_name": upload["file_name"], "auto_list_count": v["ready_count"]})
    s = get_settings()
    on = s["auto_enabled"] == "1"
    extra = (f" Reminder schedules read from the Excel for {n_rules} vendors." if n_rules else "") + \
            (" " + "; ".join(rule_problems[:3]) + ("…" if len(rule_problems) > 3 else "") + "." if rule_problems else "")
    return jsonify(ok=True, rule_problems=rule_problems, message=f"Saved as the vendor list: {v['ready_count']} vendors, amounts as of "
                                    f"{today.strftime('%d-%b-%Y')}. " + (
                                    f"Reminders go out on each vendor's schedule until you upload a new list."
                                    if on else "To send automatically, press \"Turn on automatic sending\".") + extra)


@app.get("/api/vendors")
@login_required
def api_vendors():
    s = get_settings()
    try:
        wl = load_weekly_list(s)
    except AutoStop as stop:
        return jsonify(vendors=[], message=str(stop), weekly_day=DAYS[int(s["auto_day"])], time=s["auto_time"])
    rules, last, now = get_rules(), get_last(), now_ist()
    paused = date.fromisoformat(s["auto_paused_until"]) if s["auto_paused_until"] else None
    out = []
    for r in wl["validation"]["ready"]:
        rule = rules.get(r["phone"]) or {"freq": "weekly", "start": ""}
        nd = next_due(rule, last.get(r["phone"]), s, paused, now) if s["auto_enabled"] == "1" else None
        out.append({"name": r["name"], "phone": r["phone"], "amount": r["amount_text"], "freq": rule["freq"],
                    "start": rule["start"], "last": last[r["phone"]].strftime("%d-%b-%Y") if r["phone"] in last else "",
                    "next": nd.strftime("%a %d-%b-%Y") if nd else ("" if s["auto_enabled"] == "1" else "reminders are off")})
    return jsonify(vendors=out, weekly_day=DAYS[int(s["auto_day"])], time=s["auto_time"],
                   skip_sunday=s["auto_skip_sunday"] == "1")


@app.post("/api/vendors/<phone>")
@login_required
def api_vendor_set(phone):
    p = request.get_json(force=True)
    freq = p.get("freq") or "weekly"
    if freq not in FREQ_LABELS:
        return fail("Choose how often.")
    start = (p.get("start") or "").strip()
    if start:
        try:
            date.fromisoformat(start)
        except ValueError:
            return fail("Choose a valid date.")
    num, _ = normalize_mobile(phone)
    if not num:
        return fail("Unknown mobile number.")
    set_rule(num, p.get("name") or "", freq, start)
    s = get_settings()
    paused = date.fromisoformat(s["auto_paused_until"]) if s["auto_paused_until"] else None
    nd = next_due({"freq": freq, "start": start}, get_last().get(num), s, paused, now_ist()) if s["auto_enabled"] == "1" else None
    return jsonify(ok=True, next=nd.strftime("%a %d-%b-%Y") if nd else ("" if s["auto_enabled"] == "1" else "reminders are off"))


@app.post("/api/auto/test")
@login_required
def api_auto_test():
    return jsonify(run_auto(test=True))


# ===========================================================================
# Screens
# ===========================================================================
def _resource_dir() -> Path:
    """Folder holding ui/: next to app.py, or inside the PyInstaller bundle."""
    return Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))


UI_DIR = _resource_dir() / "ui"
LOGIN_HTML = (UI_DIR / "login.html").read_text(encoding="utf-8")
PAGE_HTML = (UI_DIR / "index.html").read_text(encoding="utf-8")


def create_app() -> Flask:
    init_db()
    s = get_settings()
    if os.name == "nt" and s["auto_enabled"] == "1" and s["auto_task_kind"] != "daily":
        schedule_task(int(s["auto_day"]), s["auto_time"])     # older versions ran only on the weekly day
    app.secret_key = get_settings(include_secret=True)["secret_key"]
    return app


def main() -> None:
    parser = argparse.ArgumentParser(description="Outstanding reminders")
    parser.add_argument("--host", default="127.0.0.1", help="0.0.0.0 lets others on the network use it")
    parser.add_argument("--port", type=int, default=5050)
    parser.add_argument("--no-browser", action="store_true")
    parser.add_argument("--auto-send", action="store_true", help="send the reminders due today and exit")
    args = parser.parse_args()
    create_app()
    if args.auto_send:
        print("\n  Outstanding Reminders: daily check for reminders that are due. This window closes by itself.\n")
        result = run_auto()
        print(f"  {result['status'].upper()}: {result['reason']}\n")
        time.sleep(5)
        return
    s = get_settings()
    if s["auto_enabled"] == "1" and os.name == "nt":       # older versions created a weekly task; refresh it
        ok, _ = schedule_task(int(s["auto_day"]), s["auto_time"])
        if ok:
            auto_log("Windows schedule refreshed (daily check)")
    import logging
    import flask.cli
    logging.getLogger("werkzeug").setLevel(logging.ERROR)     # no request logs or dev-server banner
    flask.cli.show_server_banner = lambda *a, **k: None
    url = f"http://127.0.0.1:{args.port}"
    print("\n  Outstanding Reminders is running.")
    print(f"  Your browser should open by itself. If not, go to {url}")
    if args.host == "0.0.0.0":
        print(f"  Others on this network can use http://<this computer's IP>:{args.port}")
    if using_default_password():
        print("  First-time password: change-me (change it in Settings).")
    print("\n  Keep this window open while you use the app. Close it to stop.\n")
    if not args.no_browser:
        threading.Timer(1.2, lambda: webbrowser.open(url)).start()
    try:
        app.run(host=args.host, port=args.port, threaded=True)
    except OSError:
        print(f"  Port {args.port} is busy: the app may already be running. Opening it in the browser.")
        webbrowser.open(url)
        time.sleep(3)


if __name__ == "__main__":
    main()
