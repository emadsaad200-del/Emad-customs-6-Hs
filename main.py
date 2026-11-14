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
        req = urllib.request.Request(
            f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={GEMINI_KEY}",
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
    if str(user.id) in pending_users:
        await update.message.reply_text("طلبك قيد المراجعة")
        return False
    pending_users[str(user.id)] = {"name": user.full_name, "username": user.username}
    save_json(PENDING_FILE, pending_users)
    await update.message.reply_text("البوت خاص. تم ارسال طلبك للادمن.")
    btn1 = InlineKeyboardButton(text="موافقة", callback_data="approve_" + str(user.id))
    btn2 = InlineKeyboardButton(text="رفض", callback_data="reject_" + str(user.id))
    kb = InlineKeyboardMarkup([[btn1, btn2]])
    try:
        await context.bot.send_message(chat_id=ADMIN_ID, text="طلب جديد: " + user.full_name + " ID:" + str(user.id), reply_markup=kb)
    except Exception:
        pass
    return False

async def start(update, context):
    ok = await request_approval(update, context)
    if not ok:
        return
    await update.message.reply_text("اهلا بيك في بوت Logistics Group\n\nجرب تبعت بس من غير اوامر:\n\n8471 -> هيجيب البند\nزجاج -> هيجيب البند بالاسم\n8471 100000 -> هيحسبلك الجمارك\nمنشور 28 -> هيجيب المنشور\n/gpc paint -> هيجيب تصنيف GS1\n📸 ابعت صورة فاتورة -> هقراها وأحسبها بالمصري\n/invoices -> آخر الفواتير\n\nبس كده!")

async def gpc_command(update, context):
    ok = await request_approval(update, context)
    if not ok:
        return
    if not context.args:
        await update.message.reply_text("اكتب كده:\n/gpc paint\n/gpc milk\n/gpc 10001686")
        return
    q = " ".join(context.args).strip()
    rows = search_gs1(q)
    if not rows:
        await update.message.reply_text(f"❌ مفيش Brick لـ: {q}\nجرب كلمة انجليزي زي paint / milk / cheese")
        return
    for r in rows:
        msg = f"🧱 GS1 Brick: {r[1]}\n"
        msg += f"🔢 Brick Code: {r[0]}\n"
        msg += f"📂 Segment: {r[2]}\n"
        msg += f"📁 Family: {r[3]}\n"
        msg += f"📄 Class: {r[4]}\n"
        if r[5]:
            msg += f"\n✅ Includes: {r[5][:350]}\n"
        if r[6]:
            msg += f"\n⛔ Excludes: {r[6][:350]}"
        await update.message.reply_text(msg[:4000])

async def invoices_command(update, context):
    ok = await request_approval(update, context)
    if not ok:
        return
    try:
        conn = sqlite3.connect("items.db")
        cur = conn.cursor()
        cur.execute("SELECT supplier, total_foreign, currency, total_tax_egp, grand_total_egp, created_at FROM invoices WHERE user_id=? ORDER BY id DESC LIMIT 10", (update.effective_user.id,))
        rows = cur.fetchall()
        conn.close()
    except Exception:
        rows = []
    if not rows:
        await update.message.reply_text("لسه مفيش فواتير محفوظة.\nابعت صورة فاتورة عشان نبدأ.")
        return
    msg = "🧾 آخر الفواتير:\n\n"
    for r in rows:
        msg += f"• {r[0] or 'بدون مورد'} | {r[1]:,.0f} {r[2]} | ضرائب: {r[3]:,.0f} ج | إجمالي: {r[4]:,.0f} ج | {r[5]}\n"
    await update.message.reply_text(msg[:4000])

async def handle_photo(update, context):
    ok = await request_approval(update, context)
    if not ok:
        return
    await update.message.reply_text("📸 وصلت صورة الفاتورة، بقراها... استنى ثواني")
    try:
        photo = update.message.photo[-1]
        f = await context.bot.get_file(photo.file_id)
        image_bytes = await f.download_as_bytearray()
        image_bytes = bytes(image_bytes)
        data = read_invoice_image(image_bytes)
        if not data:
            await update.message.reply_text("⚠️ مش قادر أقرا الصورة أوتوماتيك (مفيش GEMINI_KEY).\nابعت البيانات كده:\nفاتورة: المورد | المنشأ | القيمة | العملة | FOB\nمثال:\nفاتورة: Rollmed | الصين | 261152 | USD | FOB")
            return
        usd_rate = get_usd_rate()
        items = []
        for it in data.get("items", []) or []:
            qty = float(it.get("qty", 0) or 0)
            price = float(it.get("price", 0) or 0)
            total = float(it.get("total", 0) or qty * price)
            items.append({"name": str(it.get("name", "")), "code": str(it.get("code", "")).replace(".", ""), "qty": qty, "price": price, "total": total})
        inv = {"supplier": str(data.get("supplier", "")), "importer": str(data.get("importer", "")), "origin": str(data.get("origin", "")), "currency": str(data.get("currency", "USD")).upper() or "USD", "incoterm": str(data.get("incoterm", "FOB")).upper() or "FOB", "total_foreign": float(data.get("total", 0) or 0), "usd_rate": usd_rate, "items": items}
        pending_invoices[update.effective_user.id] = inv
        await update.message.reply_text(format_invoice_preview(inv)[:4000], reply_markup=invoice_confirm_keyboard())
    except Exception as e:
        print(f"Photo error: {e}", flush=True)
        await update.message.reply_text("حصل مشكلة في قراءة الصورة. ابعت البيانات يدوي:\nفاتورة: المورد | المنشأ | القيمة | العملة | FOB")

async def handle_invoice_text(text, update):
    try:
        parts = text.replace("فاتورة:", "").replace("فاتورة", "").strip().split("|")
        supplier = parts[0].strip() if len(parts) > 0 else ""
        origin = parts[1].strip() if len(parts) > 1 else ""
        total = float(parts[2].strip().replace(",", "")) if len(parts) > 2 else 0
        currency = parts[3].strip().upper() if len(parts) > 3 else "USD"
        incoterm = parts[4].strip().upper() if len(parts) > 4 else "FOB"
    except Exception:
        await update.message.reply_text("صيغة الفاتورة غلط. ابعت كده:\nفاتورة: Rollmed | الصين | 261152 | USD | FOB")
        return
    usd_rate = get_usd_rate()
    pending_invoices[update.effective_user.id] = {"supplier": supplier, "importer": "", "origin": origin, "total_foreign": total, "currency": currency, "incoterm": incoterm, "usd_rate": usd_rate, "items": []}
    await update.message.reply_text(f"📋 قريت الفاتورة:\n\nالمورد: {supplier}\nالمنشأ: {origin}\nالقيمة: {total:,.0f} {currency}\nالتسليم: {incoterm}\nسعر الدولار: {usd_rate:.2f} ج\n\nدلوقتي ابعت الأصناف كده:\nصنف: اسم الصنف | البند | الكمية | السعر\nمثال: صنف: أقراص طحن | 6804.22 | 1000 | 50\n\nولما تخلص، دوس ✅ صح كمل", reply_markup=invoice_confirm_keyboard())

async def handle_smart(update, context):
    ok = await request_approval(update, context)
    if not ok:
        return
    text = update.message.text.strip()
    low = text.lower()
    if low.startswith("فاتورة"):
        await handle_invoice_text(text, update)
        return
    if low.startswith("صنف:") or low.startswith("صنف"):
        uid = update.effective_user.id
        if uid not in pending_invoices:
            await update.message.reply_text("ابعت الفاتورة الأول كده:\nفاتورة: المورد | المنشأ | القيمة | العملة | FOB")
            return
        try:
            p = text.replace("صنف:", "").replace("صنف", "").strip().split("|")
            name = p[0].strip()
            code = p[1].strip().replace(".", "") if len(p) > 1 else ""
            qty = float(p[2].strip()) if len(p) > 2 else 0
            price = float(p[3].strip()) if len(p) > 3 else 0
            pending_invoices[uid]["items"].append({"name": name, "code": code, "qty": qty, "price": price, "total": qty * price})
            await update.message.reply_text(f"✅ أضفت: {name} | بند {code} | {qty} × {price} = {qty*price:,.0f}\nابعت صنف تاني أو دوس ✅ صح كمل", reply_markup=invoice_confirm_keyboard())
        except Exception:
            await update.message.reply_text("صيغة الصنف غلط. ابعت كده:\nصنف: أقراص طحن | 6804.22 | 1000 | 50")
        return
    if low.startswith("gpc ") or low.startswith("gs1 ") or low.startswith("/gpc"):
        q = text.replace("/gpc", "").replace("gpc", "").replace("gs1", "").replace("GS1", "").strip()
        if not q:
            await update.message.reply_text("اكتب كده:\n/gpc paint")
            return
        rows = search_gs1(q)
        if not rows:
            await update.message.reply_text(f"❌ مفيش Brick لـ: {q}")
            return
        for r in rows:
            msg = f"🧱 GS1 Brick: {r[1]}\n🔢 Brick Code: {r[0]}\n📂 Segment: {r[2]}\n📁 Family: {r[3]}\n📄 Class: {r[4]}"
            await update.message.reply_text(msg[:4000])
        return
    if "منشور" in low or "manshour" in low or "قرار" in low:
        num = "".join([c for c in text if c.isdigit()])
        if not num:
            await update.message.reply_text("اكتب رقم المنشور: مثال منشور 28")
            return
        await update.message.reply_text(f"بدور على منشور {num} ...")
        try:
            q = f"منشور جمارك {num} نافذة"
            headers = {"User-Agent": "Mozilla/5.0"}
            r = requests.get("https://www.google.com/search?q=" + q, headers=headers, timeout=12)
            soup = BeautifulSoup(r.text, "html.parser")
            found = []
            for a in soup.find_all("a", href=True):
                t = a.get_text(strip=True)
                if len(t) > 20 and num in t:
                    found.append("• " + t[:120])
                if len(found) >= 5:
                    break
            if found:
                msg = f"📜 منشور {num} - نتائج البحث:\n\n" + "\n\n".join(found)
                msg += f"\n\n🔗 شوفه كامل في نافذة:\nhttps://www.nafeza.gov.eg/"
                msg += f"\n🔗 بحث جوجل:\nhttps://www.google.com/search?q={q.replace(' ', '+')}"
            else:
                msg = f"📜 منشور {num}\n\nملقيتوش في البحث التلقائي.\n\nشوفه هنا:\n1- نافذة: https://www.nafeza.gov.eg/\n2- جوجل: https://www.google.com/search?q=منشور+جمارك+{num}\n\nلو لقيته ابعتهولي وانا احفظه."
            await update.message.reply_text(msg)
        except Exception:
            await update.message.reply_text(f"📜 منشور {num}\n\nشوفه هنا:\nhttps://www.nafeza.gov.eg/\nhttps://www.google.com/search?q=منشور+جمارك+{num}")
        return
    parts = text.replace(",", " ").split()
    if len(parts) == 2 and parts[0].replace(".", "").isdigit() and parts[1].replace(".", "").isdigit():
        code = parts[0].replace(".", "")
        try:
            cif = float(parts[1])
        except Exception:
            cif = 0
        tariff = get_tariff(code)
        if not tariff:
            await update.message.reply_text(f"البند {code} مش موجود\nجرب بند تاني او ابعت للادمن النسب")
            return
        duty_p = float(tariff.get("duty", 0))
        vat_p = float(tariff.get("vat", 14))
        profit_p = float(tariff.get("profit") or 1.0)
        dev_p = float(tariff.get("dev", get_dev_fee(tariff.get("code", code))))
        table_p = float(tariff.get("table", 0))
        dumping_p = float(tariff.get("dumping", 0))
        duty, vat, profit, total, dev, table, dumping = calc_customs(cif, duty_p, vat_p, profit_p, dev_p, table_p, dumping_p)
        msg = f"🧾 حساب - البند: {code}\n"
        if tariff.get("desc_ar"):
            msg += f"{clean_desc(tariff.get('desc_ar'))}\n"
        if tariff.get("desc_en"):
            msg += f"{clean_desc(tariff.get('desc_en'))}\n"
        msg += f"قيمة CIF: {cif:,.0f} ج\n\n"
        msg += f"وارد {duty_p}% = {duty:,.0f} ج\n"
        msg += f"رسم تنمية {dev_p}% = {dev:,.0f} ج\n"
        if table_p:
            msg += f"ضريبة جدول {table_p}% = {table:,.0f} ج\n"
        if dumping_p:
            msg += f"رسم إغراق {dumping_p}% = {dumping:,.0f} ج\n"
        msg += f"قيمة مضافة {vat_p}% = {vat:,.0f} ج\n"
        if profit_p:
            msg += f"ارباح تجارية {profit_p}% = {profit:,.0f} ج\n"
        msg += "----------------\n"
        msg += f"اجمالي تقديري: {total:,.0f} ج\n"
        msg += f"الاجمالي + البضاعة: {cif + total:,.0f} ج\n"
        if tariff.get('notes'):
            msg += f"\nملاحظة: {tariff.get('notes','')}"
        await update.message.reply_text(msg)
        return
    ex_found, _ = search_combined(text)
    if not ex_found:
        await update.message.reply_text("مفيش نتائج لـ: " + text + "\n\nتأكد ان ملف items.db موجود في FadeHost\nاو جرب: منشور 28 - او - 8471 100000")
        return
    it = ex_found[0]
    tariff = get_tariff(it['clean'])
    msg = format_item_full(it['code'], it.get('desc_ar', ''), it.get('desc_en', ''), tariff)
    if tariff:
        msg += f"\n💰 وارد: {tariff.get('duty',0)}% | تنمية: {tariff.get('dev',0)}% | قيمة مضافة: {tariff.get('vat',14)}% | ارباح: {tariff.get('profit',0)}%\n"
        if tariff.get('notes'):
            msg += f"📋 {tariff.get('notes','')}\n"
    else:
        msg += "\n💰 النسب: مش موجودة\n"
    msg += f"\n🔗 التأكد من نافذة:\nhttps://www.nafeza.gov.eg/\n\n"
    msg += "💡 عشان تحسبها ابعت:\n" + it['clean'] + " 100000"
    if len(ex_found) > 1:
        msg += "\n\n🔍 مشابهة:\n"
        for it2 in ex_found[1:4]:
            msg += f"- [{it2['code']}] {it2.get('desc_ar','')[:40]} | {it2.get('desc_en','')[:40]} | فئة: {it2.get('duty',0)}%\n"
    await update.message.reply_text(msg[:4000])

async def button_callback(update, context):
    q = update.callback_query
    await q.answer()
    if q.data in ["inv_confirm", "inv_edit"]:
        uid = q.from_user.id
        if uid not in pending_invoices:
            await q.edit_message_text("مفيش فاتورة معلقة، ابعت فاتورة جديدة.")
            return
        if q.data == "inv_edit":
            await q.edit_message_text("ابعت التعديل كده:\nفاتورة: المورد | المنشأ | القيمة | العملة | FOB\nأو ابعت صنف جديد:\nصنف: اسم الصنف | البند | الكمية | السعر")
            return
        inv = pending_invoices[uid]
        if not inv.get("items"):
            inv["items"] = [{"name": "إجمالي الفاتورة", "code": "", "qty": 1, "price": inv["total_foreign"], "total": inv["total_foreign"]}]
        result = calc_invoice_egp(inv["items"], inv["usd_rate"])
        save_invoice(uid, {"supplier": inv["supplier"], "importer": inv.get("importer", ""), "origin": inv["origin"], "currency": inv["currency"], "incoterm": inv["incoterm"], "total_foreign": inv["total_foreign"], "usd_rate": inv["usd_rate"], "cif_egp": result["cif_egp"], "duty_egp": result["duty_egp"], "vat_egp": result["vat_egp"], "profit_egp": result["profit_egp"], "total_tax_egp": result["total_tax_egp"], "grand_total_egp": result["grand_total_egp"], "items": inv["items"]})
        msg = "✅ تم التأكيد - الحسبة بالمصري\n\n"
        msg += f"المورد: {inv['supplier']}\nالمنشأ: {inv['origin']}\n"
        msg += f"سعر الدولار: {inv['usd_rate']:.2f} ج\n\n"
        for d in result["details"]:
            msg += f"• {d['name']} | بند {d['code']} | CIF: {d['cif_egp']:,.0f} ج | وارد {d['duty_p']}% = {d['duty']:,.0f} ج\n"
        msg += f"\nCIF: {result['cif_egp']:,.0f} ج\n"
        msg += f"الوارد: {result['duty_egp']:,.0f} ج\n"
        msg += f"قيمة مضافة: {result['vat_egp']:,.0f} ج\n"
        msg += f"أ.ت.ص: {result['profit_egp']:,.0f} ج\n"
        msg += f"تنمية: {result['dev_egp']:,.0f} ج\n"
        msg += f"----------------\n"
        msg += f"إجمالي الضرائب: {result['total_tax_egp']:,.0f} ج\n"
        msg += f"الإجمالي + البضاعة: {result['grand_total_egp']:,.0f} ج\n"
        msg += f"\n💾 اتحفظت في /invoices"
        pending_invoices.pop(uid, None)
        await q.edit_message_text(msg[:4000])
        return
    if q.from_user.id != ADMIN_ID:
        await q.edit_message_text("لست الادمن")
        return
    data = q.data
    if data.startswith("approve_"):
        uid = int(data.split("_")[1])
        allowed_users.add(uid)
        save_json(ALLOWED_FILE, list(allowed_users))
        pending_users.pop(str(uid), None)
        save_json(PENDING_FILE, pending_users)
        await q.edit_message_text("تمت الموافقة على " + str(uid))
        try:
            await context.bot.send_message(chat_id=uid, text="تمت الموافقة! ابعت /start")
        except Exception:
            pass
    else:
        uid = data.split("_")[1]
        pending_users.pop(uid, None)
        save_json(PENDING_FILE, pending_users)
        await q.edit_message_text("تم رفض " + uid)

async def admin_list(update, context):
    if update.effective_user.id != ADMIN_ID:
        return
    text = "👥 المصرح لهم:\n"
    for uid in allowed_users:
        text += f"• {uid}\n"
    text += f"\n⏳ المنتظرين:\n"
    for pid, info in pending_users.items():
        text += f"• {pid} - {info.get('name','')} @{info.get('username','')}\n"
    text += "\nعشان تشيل حد ابعت:\n/remove + الـ ID\nمثال: /remove 123456789"
    await update.message.reply_text(text)

async def remove_user(update, context):
    if update.effective_user.id != ADMIN_ID:
        return
    if not context.args:
        await update.message.reply_text("ابعت كده:\n/remove 123456789")
        return
    try:
        uid = int(context.args[0])
    except Exception:
        await update.message.reply_text("الـ ID لازم يكون رقم")
        return
    allowed_users.discard(uid)
    pending_users.pop(str(uid), None)
    save_json(ALLOWED_FILE, list(allowed_users))
    save_json(PENDING_FILE, pending_users)
    await update.message.reply_text(f"✅ تمت إزالة {uid}، لو بعت تاني هيجيلك طلب موافقة جديد.")
    try:
        await context.bot.send_message(chat_id=uid, text="تمت إزالتك من البوت، لو عايز تدخل تاني ابعت /start وانتظر الموافقة.")
    except Exception:
        pass

if TOKEN:
    print("Bot polling started...", flush=True)
    app = Application.builder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("gpc", gpc_command))
    app.add_handler(CommandHandler("invoices", invoices_command))
    app.add_handler(CommandHandler("users", admin_list))
    app.add_handler(CommandHandler("remove", remove_user))
    app.add_handler(CallbackQueryHandler(button_callback))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_smart))
    app.run_polling()