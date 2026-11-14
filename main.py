import os
import json
import re
import base64
import unicodedata
import sqlite3
import urllib.request

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
GEMINI_KEY = os.getenv("GEMINI_KEY") or os.getenv("GOOGLE_API_KEY") or ""
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
    for t in tariff_items:
        tc = str(t.get("code", "")).replace(".", "").replace(" ", "")
        if tc == code_clean or code_clean.startswith(tc) or tc.startswith(code_clean):
            return t
    try:
        conn = sqlite3.connect("items.db")
        cur = conn.cursor()
        cur.execute("SELECT code, desc_ar, desc_en, duty FROM items WHERE REPLACE(code,'.','')=?", (code_clean,))
        row = cur.fetchone()
        if not row:
            cur.execute("SELECT code, desc_ar, desc_en, duty FROM items WHERE REPLACE(code,'.','') LIKE ?", (code_clean[:6] + "%",))
            row = cur.fetchone()
        conn.close()
        if row:
            code_full = row[0]
            duty = float(row[3] or 0)
            return {"code": code_full, "duty": duty, "vat": get_vat(code_full), "dev": get_dev_fee(code_full), "table": get_table_tax(code_full), "dumping": get_dumping(code_full), "profit": 0, "desc_ar": row[1] or "", "desc_en": row[2] or ""}
    except Exception:
        pass
    return None

def calc_customs(cif, duty_p, vat_p, profit_p, dev_p=0, table_p=0, dumping_p=0):
    duty = cif * duty_p / 100
    dev = cif * dev_p / 100
    table = (cif + duty) * table_p / 100 if table_p else 0
    dumping = cif * dumping_p / 100 if dumping_p else 0
    vat = (cif + duty + dev + table + dumping) * vat_p / 100
    profit = (cif + duty) * profit_p / 100
    total = duty + vat + profit + dev + table + dumping
    return duty, vat, profit, total, dev, table, dumping

def calc_invoice_egp(items, usd_rate, freight_egp=0, insurance_egp=0):
    total_cif_egp = 0
    total_duty = 0
    total_vat = 0
    total_profit = 0
    total_dev = 0
    details = []
    for it in items:
        cif_foreign = float(it.get("total", 0))
        cif_egp = cif_foreign * usd_rate + freight_egp + insurance_egp
        code = str(it.get("code", "")).replace(".", "")
        tariff = get_tariff(code) if code else None
        duty_p = float(tariff.get("duty", 0)) if tariff else 0
        vat_p = float(tariff.get("vat", 14)) if tariff else 14
        profit_p = float(tariff.get("profit") or 1.0) if tariff else 1.0
        dev_p = float(tariff.get("dev", 0)) if tariff else 0
        duty, vat, profit, total, dev, table, dumping = calc_customs(cif_egp, duty_p, vat_p, profit_p, dev_p)
        total_cif_egp += cif_egp
        total_duty += duty
        total_vat += vat
        total_profit += profit
        total_dev += dev
        details.append({"name": it.get("name", ""), "code": code, "cif_egp": cif_egp, "duty_p": duty_p, "duty": duty, "vat": vat, "profit": profit, "total_tax": total})
    total_tax = total_duty + total_vat + total_profit + total_dev
    return {"details": details, "cif_egp": total_cif_egp, "duty_egp": total_duty, "vat_egp": total_vat, "profit_egp": total_profit, "dev_egp": total_dev, "total_tax_egp": total_tax, "grand_total_egp": total_cif_egp + total_tax}

def format_item_full(code, desc_ar, desc_en, tariff):
    msg = f"📦 البند: {code}\n"
    if desc_ar:
        msg += f"{clean_desc(desc_ar)}\n"
    if desc_en:
        msg += f"{clean_desc(desc_en)}\n"
    msg += "\n"
    if tariff:
        msg += f"🧾 الفئة الجمركية: {tariff.get('duty',0)}%\n"
        msg += f"🧾 رسم التنمية: {tariff.get('dev', get_dev_fee(code))}%\n"
        msg += f"🧾 ض. القيمة المضافة: {tariff.get('vat',14)}%\n"
        if tariff.get('table'):
            msg += f"🧾 ضريبة جدول: {tariff.get('table')}%\n"
        if tariff.get('dumping'):
            msg += f"⚠️ رسم إغراق / وقائي: {tariff.get('dumping')}% - انتبه فيه منشور إغراق على البند ده\n"
        else:
            msg += f"🧾 رسم إغراق: لا يوجد\n"
        if tariff.get('vat', 14) == 0:
            msg += "✅ معفي من القيمة المضافة حسب الملف\n"
    return msg

def search_combined(query):
    query = clean_desc(query.strip())
    q_lower = query.lower()
    digits = "".join([ch for ch in query if ch.isdigit()])
    if not os.path.exists("items.db"):
        return [], []
    conn = sqlite3.connect("items.db")
    cur = conn.cursor()
    result = []
    try:
        if digits and len(digits) >= 2 and len(query.replace(" ", "")) <= 8:
            cur.execute("SELECT code, desc_ar, desc_en, duty FROM items WHERE clean = ? LIMIT 1", (digits,))
            row = cur.fetchone()
            if row:
                conn.close()
                return [{'code': row[0], 'desc_ar': clean_desc(row[1]), 'desc_en': clean_desc(row[2]), 'desc': clean_desc((row[1] or "") + " | " + (row[2] or "")), 'duty': row[3], 'clean': digits}], []
            cur.execute("SELECT code, desc_ar, desc_en, duty FROM items WHERE clean LIKE ? LIMIT 10", (digits + "%",))
            rows = cur.fetchall()
            conn.close()
            for r in rows:
                result.append({'code': r[0], 'desc_ar': clean_desc(r[1]), 'desc_en': clean_desc(r[2]), 'desc': clean_desc((r[1] or "") + " | " + (r[2] or "")), 'duty': r[3], 'clean': r[0].replace(".", "").replace(" ", "")})
            return result, []
        words = [w for w in q_lower.split() if len(w) > 1]
        if not words:
            conn.close()
            return [], []
        cur.execute("SELECT code, desc_ar, desc_en, duty FROM items WHERE desc_ar LIKE ? OR desc_en LIKE ? LIMIT 7", ("%" + words[0] + "%", "%" + words[0] + "%"))
        rows = cur.fetchall()
        conn.close()
        for r in rows:
            result.append({'code': r[0], 'desc_ar': clean_desc(r[1]), 'desc_en': clean_desc(r[2]), 'desc': clean_desc((r[1] or "") + " | " + (r[2] or "")), 'duty': r[3], 'clean': r[0].replace(".", "").replace(" ", "")})
        return result, []
    except Exception as e:
        print(f"Search error: {e}", flush=True)
        try:
            conn.close()
        except Exception:
            pass
        return [], []

# ============ قراءة صورة الفاتورة - Gemini مجاني ============
def read_invoice_image(image_bytes):
    if not GEMINI_KEY:
        print("Gemini OCR error: GEMINI_KEY is empty", flush=True)
        return None
    try:
        b64 = base64.b64encode(image_bytes).decode("utf-8")
        prompt = """Read this commercial invoice. Return ONLY valid JSON like:
        {"supplier": "", "importer": "", "origin": "", "currency": "USD", "incoterm": "FOB", "total": 0, "items": [{"name": "", "code": "", "qty": 0, "price": 0, "total": 0}]}
        If HS code not on invoice, leave code empty. If origin not found, leave empty."""
        data = {
            "contents": [{
                "parts": [
                    {"text": prompt},
                    {"inline_data": {"mime_type": "image/jpeg", "data": b64}}
                ]
            }]
        }
        for model in ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash"]:
            try:
                req = urllib.request.Request(
                    f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={GEMINI_KEY}",
                    data=json.dumps(data).encode("utf-8"),
                    headers={"Content-Type": "application/json"},
                    method="POST"
                )
                with urllib.request.urlopen(req, timeout=60) as resp:
                    res = json.loads(resp.read().decode("utf-8"))
                text = res["candidates"][0]["content"]["parts"][0]["text"]
                text = text.replace("```json", "").replace("```", "").strip()
                return json.loads(text)
            except Exception as e:
                print(f"Gemini model {model} failed: {e}", flush=True)
                continue
        return None
    except Exception as e:
        print(f"Gemini OCR error: {e}", flush=True)
        return None

def invoice_confirm_keyboard():
    return InlineKeyboardMarkup([[InlineKeyboardButton("✅ صح كمل", callback_data="inv_confirm"), InlineKeyboardButton("✏️ تعديل", callback_data="inv_edit")]])

def format_invoice_preview(inv):
    msg = "📋 قريت الفاتورة:\n\n"
    msg += f"المورد: {inv.get('supplier','')}\n"
    msg += f"المستورد: {inv.get('importer','')}\n"
    msg += f"المنشأ: {inv.get('origin','')}\n"
    msg += f"القيمة: {inv.get('total_foreign',0):,.0f} {inv.get('currency','USD')}\n"
    msg += f"التسليم: {inv.get('incoterm','FOB')}\n"
    msg += f"سعر الدولار: {inv.get('usd_rate',0):.2f} ج\n\n"
    if inv.get("items"):
        msg += "الأصناف:\n"
        for it in inv["items"]:
            msg += f"- {it.get('name','')} | بند {it.get('code','')} | {it.get('qty',0)} x {it.get('price',0)} = {it.get('total',0):,.0f}\n"
    else:
        msg += "مفيش أصناف، ابعت صنف كده:\nصنف: اسم الصنف | البند | الكمية | السعر\n"
    msg += "\nلو صح دوس ✅ صح كمل"
    return msg

# ============ الهاندلرز ============
async def request_approval(update, context):
    user = update.effective_user
    if user.id in allowed_users or user.id == ADMIN_ID:
        return True
    if str(user.id) in pending