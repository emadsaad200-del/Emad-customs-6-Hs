import os
import json
import sqlite3
import requests
from bs4 import BeautifulSoup
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, CallbackQueryHandler, filters, ContextTypes

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
        print(f"save error: {e}", flush=True)

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
            cur.execute("SELECT * FROM tariffs WHERE code LIKE ?", (code[:4] + "%",))
            row = cur.fetchone()
        conn.close()
        if row:
            return dict(row)
    except Exception as e:
        print(f"tariff error: {e}", flush=True)
    return None

def get_dev_fee(code):
    return 0

def calc_customs(cif, duty_p, vat_p, profit_p, dev_p, table_p, dumping_p):
    duty = cif * duty_p / 100
    dev = cif * dev_p / 100
    table = cif * table_p / 100
    dumping = cif * dumping_p / 100
    vat_base = cif + duty + dev + table + dumping
    vat = vat_base * vat_p / 100
    profit = vat_base * profit_p / 100
    total = duty + dev + table + dumping + vat + profit
    return duty, vat, profit, total, dev, table, dumping

def search_combined(text):
    text = text.strip()
    results = []
    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute("SELECT code, desc_ar, desc_en FROM items WHERE code LIKE ? OR desc_ar LIKE ? OR desc_en LIKE ? LIMIT 5", (f"%{text}%", f"%{text}%", f"%{text}%"))
        rows = cur.fetchall()
        conn.close()
        for r in rows:
            code = r[0]
            tariff = get_tariff(code)
            results.append({"code": code, "clean": str(code).replace(".", ""), "desc_ar": r[1] or "", "desc_en": r[2] or "", "duty": tariff.get("duty", 0) if tariff else 0})
    except Exception as e:
        print(f"search error: {e}", flush=True)
    return results, []

def format_item_full(code, desc_ar, desc_en, tariff):
    msg = f"📦 البند: {code}\n"
    if desc_ar:
        msg += f"{clean_desc(desc_ar)}\n"
    if desc_en:
        msg += f"{clean_desc(desc_en)}\n"
    return msg

def search_gs1(q):
    q = q.strip()
    out = []
    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute("SELECT brick_code, brick_name, segment, family, class_name, includes_text, excludes_text FROM gs1 WHERE brick_code LIKE ? OR brick_name LIKE ? LIMIT 3", (f"%{q}%", f"%{q}%"))
        rows = cur.fetchall()
        conn.close()
        for r in rows:
            out.append(r)
    except Exception as e:
        print(f"gs1 error: {e}", flush=True)
    return out

def get_usd_rate():
    try:
        r = requests.get("https://open.er-api.com/v6/latest/USD", timeout=10)
        data = r.json()
        return float(data["rates"].get("EGP", 48.0))
    except Exception:
        return 48.0

def read_invoice_image(image_bytes):
    try:
        import base64
        api_key = os.environ.get("OPENAI_API_KEY", "")
        if not api_key:
            return None
        b64 = base64.b64encode(image_bytes).decode("utf-8")
        headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
        payload = {
            "model": "gpt-4o-mini",
            "messages": [{"role": "user", "content": [
                {"type": "text", "text": "اقرا الفاتورة دي وطلع JSON بس: {supplier, importer, origin, currency, incoterm, total, items: [{name, code, qty, price, total}]}. لو مش واضح حط قيم فاضية."},
                {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64}"}}
            ]}],
            "max_tokens": 2000
        }
        r = requests.post("https://api.openai.com/v1/chat/completions", headers=headers, json=payload, timeout=60)
        txt = r.json()["choices"][0]["message"]["content"]
        txt = txt.replace("```json", "").replace("```", "").strip()
        return json.loads(txt)
    except Exception as e:
        print(f"OCR error: {e}", flush=True)
        return None

def calc_invoice_egp(items, usd_rate):
    cif_egp = 0
    duty_egp = 0
    vat_egp = 0
    profit_egp = 0
    dev_egp = 0
    details = []
    for it in items:
        code = str(it.get("code", "")).replace(".", "")
        total_foreign = float(it.get("total", 0) or 0)
        cif = total_foreign * usd_rate
        tariff = get_tariff(code) if code else None
        duty_p = float(tariff.get("duty", 0)) if tariff else 0
        vat_p = float(tariff.get("vat", 14)) if tariff else 14
        profit_p = float(tariff.get("profit", 0) or 0) if tariff else 0
        dev_p = float(tariff.get("dev", 0) or 0) if tariff else 0
        duty, vat, profit, total, dev, table, dumping = calc_customs(cif, duty_p, vat_p, profit_p, dev_p, 0, 0)
        cif_egp += cif
        duty_egp += duty
        vat_egp += vat
        profit_egp += profit
        dev_egp += dev
        details.append({"name": it.get("name", ""), "code": code, "cif_egp": cif, "duty_p": duty_p, "duty": duty})
    total_tax = duty_egp + vat_egp + profit_egp + dev_egp
    grand = cif_egp + total_tax
    return {"cif_egp": cif_egp, "duty_egp": duty_egp, "vat_egp": vat_egp, "profit_egp": profit_egp, "dev_egp": dev_egp, "total_tax_egp": total_tax, "grand_total_egp": grand, "details": details}

def format_invoice_preview(inv):
    msg = "📋 مراجعة الفاتورة:\n\n"
    msg += f"المورد: {inv.get('supplier','')}\n"
    msg += f"المنشأ: {inv.get('origin','')}\n"
    msg += f"القيمة: {inv.get('total_foreign',0):,.0f} {inv.get('currency','USD')}\n"
    msg += f"التسليم: {inv.get('incoterm','FOB')}\n"
    msg += f"سعر الدولار: {inv.get('usd_rate',0):.2f} ج\n\n"
    for it in inv.get("items", []):
        msg += f"• {it.get('name','')} | بند {it.get('code','')} | {it.get('qty',0)} × {it.get('price',0)} = {it.get('total',0):,.0f}\n"
    msg += "\nلو صح دوس ✅ صح كمل"
    return msg

def invoice_confirm_keyboard():
    b1 = InlineKeyboardButton(text="✅ صح كمل", callback_data="inv_confirm")
    b2 = InlineKeyboardButton(text="✏️ تعديل", callback_data="inv_edit")
    return InlineKeyboardMarkup([[b1, b2]])

def save_invoice(user_id, data):
    try:
        conn = sqlite3.connect("items.db")
        cur = conn.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS invoices (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, supplier TEXT, importer TEXT, origin TEXT, currency TEXT, incoterm TEXT, total_foreign REAL, usd_rate REAL, cif_egp REAL, duty_egp REAL, vat_egp REAL, profit_egp REAL, total_tax_egp REAL, grand_total_egp REAL, created_at TEXT)")
        import datetime
        cur.execute("INSERT INTO invoices (user_id, supplier, importer, origin, currency, incoterm, total_foreign, usd_rate, cif_egp, duty_egp, vat_egp, profit_egp, total_tax_egp, grand_total_egp, created_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", (user_id, data.get("supplier",""), data.get("importer",""), data.get("origin",""), data.get("currency",""), data.get("incoterm","