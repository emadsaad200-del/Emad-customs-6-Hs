import os
import json
import re
import base64
import unicodedata
import sqlite3

_orig_connect = sqlite3.connect
def _patched_connect(path, *a, **k):
    if path == "items.db":
        path = "items_v2.db"
    return _orig_connect(path, *a, **k)
sqlite3.connect = _patched_connect

import openpyxl
import requests
from bs4 import BeautifulSoup
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, CallbackQueryHandler, filters

TOKEN = os.getenv("BOT_TOKEN") or os.getenv("TELEGRAM_TOKEN")
OPENAI_KEY = os.getenv("OPENAI_KEY") or os.getenv("OPENAI_API_KEY") or ""
ADMIN_ID = 1280808746
ALLOWED_FILE = "allowed_users.json"
PENDING_FILE = "pending_users.json"
print("Starting Bot - Easy Mode...", flush=True)

def load_json(file, default):
    if os.path.exists(file):
        try:
            with open(file, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return default
    return default

def save_json(file, data):
    with open(file, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)

allowed_users = set(load_json(ALLOWED_FILE, [ADMIN_ID]))
pending_users = load_json(PENDING_FILE, {})
pending_invoices = {}

def clean_desc(text):
    if not text:
        return ""
    t = str(text).strip()
    t = unicodedata.normalize('NFKC', t)
    t = unicodedata.normalize('NFC', t)
    t = t.replace('ـ', '')
    t = ' '.join(t.split())
    return t

# ============ بناء قاعدة البيانات ============
fname = None
for n in ["Book1LOGIST TATTEF.xlsx", "Book1LOGIST_TATTEF.xlsx"]:
    if os.path.exists(n):
        fname = n
        break

if fname and not os.path.exists("items_v2.db"):
    wb = openpyxl.load_workbook(fname, read_only=True, data_only=True)
    ws = wb.active
    conn = sqlite3.connect("items.db")
    cur = conn.cursor()
    cur.execute("DROP TABLE IF EXISTS items")
    cur.execute("CREATE TABLE items (code TEXT PRIMARY KEY, desc_ar TEXT, desc_en TEXT, desc TEXT, clean TEXT, duty REAL)")
    conn.commit()
    ar = {}
    en = {}
    for row in ws.iter_rows(values_only=True):
        if not row or not row[0]:
            continue
        t = str(row[0]).strip()
        if len(t) < 7:
            continue
        code = ""
        for part in t.replace("-", " ").split():
            p = part.strip()
            if len(p) >= 6 and p[0].isdigit() and "." in p:
                if re.match(r"^\d{4}\.\d+", p):
                    code = p
                    break
        if not code:
            continue
        duty = None
        if "معفاه" in t or "معفاة" in t or "معفي" in t or "معفى" in t:
            duty = 0
        else:
            w = t.split()
            if w and w[-1].replace(".", "").isdigit():
                try:
                    duty = float(w[-1])
                except Exception:
                    duty = None
        desc = t.replace(code, "").replace("-", " ").strip()
        desc = re.sub(r"\s+\d+(\.\d+)?\s*$", "", desc).strip()
        if len(desc) < 3:
            desc = code
        is_ar = any('\u0600' <= ch <= '\u06FF' for ch in desc)
        if is_ar:
            ar[code] = (desc, duty)
        else:
            if len(desc) > 3:
                en[code] = desc
    all_codes = set(ar.keys()) | set(en.keys())
    for code in all_codes:
        desc_ar, duty = ar.get(code, ("", None))
        desc_en = en.get(code, "")
        desc_combined = (desc_ar + " | " + desc_en).strip(" |")
        clean = code.replace(".", "").replace(" ", "")
        cur.execute("INSERT INTO items VALUES (?,?,?,?,?,?)", (code, desc_ar, desc_en, desc_combined + (f" | فئة: {duty}" if duty is not None else ""), clean, duty if duty is not None else 0))
    conn.commit()
    conn.close()
    print(f"Built items.db: {len(all_codes)} items", flush=True)

def build_gs1():
    gs1_file = None
    for n in ["gs1_exle.xlsx", "gs1 exle.xlsx", "gs1_exle.XLSX", "GS1.xlsx"]:
        if os.path.exists(n):
            gs1_file = n
            break
    if not gs1_file:
        return
    try:
        conn = sqlite3.connect("items.db")
        cur = conn.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='gs1_bricks'")
        if cur.fetchone():
            conn.close()
            return
        cur.execute("DROP TABLE IF EXISTS gs1_bricks")
        cur.execute("CREATE TABLE gs1_bricks (brick_code TEXT PRIMARY KEY, brick_title TEXT, segment_title TEXT, family_title TEXT, class_title TEXT, includes_text TEXT, excludes_text TEXT)")
        wb = openpyxl.load_workbook(gs1_file, read_only=True, data_only=True)
        ws = wb.active
        seen = set()
        count = 0
        first = True
        for row in ws.iter_rows(values_only=True):
            if first:
                first = False
                continue
            if not row or len(row) < 11:
                continue
            brick_code = str(row[9]).strip() if row[9] else ""
            if not brick_code or brick_code in seen or brick_code == "BrickCode":
                continue
            seen.add(brick_code)
            brick_title = str(row[10] or "").strip()
            segment_title = str(row[1] or "").strip()
            family_title = str(row[4] or "").strip()
            class_title = str(row[7] or "").strip()
            inc = str(row[11] or "").strip()[:800]
            exc = str(row[12] or "").strip()[:800]
            cur.execute("INSERT OR IGNORE INTO gs1_bricks VALUES (?,?,?,?,?,?,?)", (brick_code, brick_title, segment_title, family_title, class_title, inc, exc))
            count += 1
            if count % 1000 == 0:
                conn.commit()
        conn.commit()
        conn.close()
        print(f"GS1 built: {count} bricks", flush=True)
    except Exception as e:
        print(f"GS1 build error: {e}", flush=True)

build_gs1()

def search_gs1(query):
    query = query.strip()
    if len(query) < 2:
        return []
    try:
        conn = sqlite3.connect("items.db")
        cur = conn.cursor()
        if query.replace(" ", "").isdigit():
            cur.execute("SELECT brick_code, brick_title, segment_title, family_title, class_title, includes_text, excludes_text FROM gs1_bricks WHERE brick_code LIKE ? LIMIT 5", (query + "%",))
        else:
            cur.execute("SELECT brick_code, brick_title, segment_title, family_title, class_title, includes_text, excludes_text FROM gs1_bricks WHERE brick_title LIKE ? LIMIT 5", ("%" + query + "%",))
        rows = cur.fetchall()
        conn.close()
        return rows
    except Exception as e:
        print(f"GS1 search error: {e}", flush=True)
        return []

def init_invoices_db():
    try:
        conn = sqlite3.connect("items.db")
        cur = conn.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS invoices (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, supplier TEXT, importer TEXT, origin TEXT, currency TEXT, incoterm TEXT, total_foreign REAL, usd_rate REAL, cif_egp REAL, duty_egp REAL, vat_egp REAL, profit_egp REAL, total_tax_egp REAL, grand_total_egp REAL, items_json TEXT, created_at TEXT DEFAULT CURRENT_TIMESTAMP)")
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"Invoices DB error: {e}", flush=True)

init_invoices_db()

def save_invoice(user_id, data):
    try:
        conn = sqlite3.connect("items.db")
        cur = conn.cursor()
        cur.execute("INSERT INTO invoices (user_id, supplier, importer, origin, currency, incoterm, total_foreign, usd_rate, cif_egp, duty_egp, vat_egp, profit_egp, total_tax_egp, grand_total_egp, items_json) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", (user_id, data.get("supplier", ""), data.get("importer", ""), data.get("origin", ""), data.get("currency", "USD"), data.get("incoterm", "FOB"), data.get("total_foreign", 0), data.get("usd_rate", 0), data.get("cif_egp", 0), data.get("duty_egp", 0), data.get("vat_egp", 0), data.get("profit_egp", 0), data.get("total_tax_egp", 0), data.get("grand_total_egp", 0), json.dumps(data.get("items", []), ensure_ascii=False)))
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"Save invoice error: {e}", flush=True)

def get_usd_rate():
    try:
        r = requests.get("https://open.er-api.com/v6/latest/USD", timeout=8)
        j = r.json()
        if j.get("rates", {}).get("EGP"):
            return float(j["rates"]["EGP"])
    except Exception as e:
        print(f"USD rate error: {e}", flush=True)
    return 53.48

# ============ قواعد الرسوم ============
VAT_DEFAULT = 14.0
VAT_EXEMPT_PREFIX = ["2309.9010"]
DEV_DEFAULT = 0.0
PROFIT_DEFAULT = 1.0
DEV_EXEMPT_CHAPTERS = ["01", "02", "03", "04", "07", "08", "10", "11", "12"]
TABLE_TAX = {}
ANTI_DUMPING = {}

def get_dev_fee(code):
    ch = code.replace(".", "")[:2] if code else ""
    if ch in DEV_EXEMPT_CHAPTERS:
        return 0.0
    return DEV_DEFAULT

def get_profit_fee(code, is_factory):
    if is_factory:
        return 0.0
    return PROFIT_DEFAULT

def get_vat(code):
    if any(code.startswith(p) for p in VAT_EXEMPT_PREFIX):
        return 0.0
    return VAT_DEFAULT

def get_table_tax(code):
    clean = code.replace(".", "")
    return TABLE_TAX.get(code, TABLE_TAX.get(clean, TABLE_TAX.get(code[:4], 0)))

def get_dumping(code):
    clean = code.replace(".", "")
    return ANTI_DUMPING.get(code, ANTI_DUMPING.get(clean, ANTI_DUMPING.get(code[:4], 0)))

tariff_items = []
if os.path.exists("Tariff.json"):
    tariff_items = load_json("Tariff.json", [])

def get_tariff(code_clean):
    for