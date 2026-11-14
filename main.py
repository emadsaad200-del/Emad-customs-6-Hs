import os
import json
import sqlite3
import requests
import datetime
from bs4 import BeautifulSoup
from telegram import Update
from telegram import InlineKeyboardButton
from telegram import InlineKeyboardMarkup
from telegram.ext import Application
from telegram.ext import CommandHandler
from telegram.ext import MessageHandler
from telegram.ext import CallbackQueryHandler
from telegram.ext import filters

TOKEN = os.environ.get("BOT_TOKEN", "")
ADMIN_ID = int(os.environ.get("ADMIN_ID", "0") or 0)
ALLOWED_FILE = "allowed.json"
PENDING_FILE = "pending.json"

def load_json(path, default):
    try:
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return default

def save_json(path, data):
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print("save error", e, flush=True)

allowed_users = set(load_json(ALLOWED_FILE, []))
pending_users = load_json(PENDING_FILE, {})
pending_invoices = {}

def clean_desc(s):
    if not s:
        return ""
    return str(s).strip()

def get_db():
    conn = sqlite3.connect("items.db")
    conn.row_factory = sqlite3.Row
    return conn

def get_tariff(code):
    code = str(code).replace(".", "").strip()
    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute("SELECT * FROM tariffs WHERE code=?", (code,))
        row = cur.fetchone()
        if not row:
            cur.execute(
                "SELECT * FROM tariffs WHERE code LIKE ?",
                (code[:4] + "%",)
            )
            row = cur.fetchone()
        conn.close()
        if row:
            return dict(row)
    except Exception as e:
        print("tariff error", e, flush=True)
    return None

def calc_customs(cif, duty_p, vat_p, profit_p, dev_p, table_p, dump_p):
    duty = cif * duty_p / 100
    dev = cif * dev_p / 100
    table = cif * table_p / 100
    dumping = cif * dump_p / 100
    base = cif + duty + dev + table + dumping
    vat = base * vat_p / 100
    profit = base * profit_p / 100
    total = duty + dev + table + dumping + vat + profit
    return duty, vat, profit, total, dev, table, dumping

def search_combined(text):
    text = text.strip()
    out = []
    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute(
            "SELECT code, desc_ar, desc_en FROM items "
            "WHERE code LIKE ? OR desc_ar LIKE ? "
            "OR desc_en LIKE ? LIMIT 5",
            ("%" + text + "%", "%" + text + "%", "%" + text + "%")
        )
        rows = cur.fetchall()
        conn.close()
        for r in rows:
            code = r[0]
            t = get_tariff(code)
            d = t.get("duty", 0) if t else 0
            out.append({
                "code": code,
                "clean": str(code).replace(".", ""),
                "desc_ar": r[1] or "",
                "desc_en": r[2] or "",
                "duty": d
            })
    except Exception as e:
        print("search error", e, flush=True)
    return out, []

def search_gs1(q):
    q = q.strip()
    out = []
    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute(
            "SELECT brick_code, brick_name, segment, "
            "family, class_name, includes_text, "
            "excludes_text FROM gs1 "
            "WHERE brick_code LIKE ? OR brick_name LIKE ? LIMIT 3",
            ("%" + q + "%", "%" + q + "%")
        )
        rows = cur.fetchall()
        conn.close()
        for r in rows:
            out.append(r)
    except Exception as e:
        print("gs1 error", e, flush=True)
    return out

def get_usd_rate():
    try:
        r = requests.get(
            "https://open.er-api.com/v6/latest/USD",
            timeout=10
        )
        return float(r.json()["rates"].get("EGP", 48.0))
    except Exception:
        return 48.0

def read_invoice_image(image_bytes):
    try:
        import base64
        api_key = os.environ.get("OPENAI_API_KEY", "")
        if not api_key:
            return None
        b64 = base64.b64encode(image_bytes).decode("utf-8")
        headers = {
            "Authorization": "Bearer " + api_key,
            "Content-Type": "application/json"
        }
        prompt = (
            "اقرا الفاتورة وطلع JSON بس: "
            "{supplier, importer, origin, currency, "
            "incoterm, total, items: "
            "[{name, code, qty, price, total}]}"
        )
        payload = {
            "model": "gpt-4o-mini",
            "messages": [{
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {"type": "image_url", "image_url": {
                        "url": "data:image/jpeg;base64," + b64
                    }}
                ]
            }],
            "max_tokens": 2000
        }
        r = requests.post(
            "https://api.openai.com/v1/chat/completions",
            headers=headers, json=payload, timeout=60
        )
        txt = r.json()["choices"][0]["message"]["content"]
        txt = txt.replace("```