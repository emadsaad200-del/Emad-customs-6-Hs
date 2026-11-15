import os
import json
import sqlite3
import requests
import datetime
import glob
import re
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

def fmt0(x):
    try:
        return format(float(x), ",.0f")
    except Exception:
        return "0"

def fmt2(x):
    try:
        return format(float(x), ",.2f")
    except Exception:
        return "0.00"

allowed_users = set(load_json(ALLOWED_FILE, []))
pending_users = load_json(PENDING_FILE, {})
pending_invoices = {}

def clean_desc(s):
    if not s:
        return ""
    return str(s).strip()

def norm_col(c):
    c = str(c).strip().lower()
    c = re.sub(r'[\s\-_]+', '', c)
    c = c.replace('é','e')
    return c

def find_file(*keywords):
    files = glob.glob("*") + glob.glob("**/*", recursive=True)
    files = [f for f in files if os.path.isfile(f)]
    for f in files:
        low = f.lower()
        if all(k.lower() in low for k in keywords):
            return f
    best = None
    for f in files:
        low = f.lower()
        if any(k.lower() in low for k in keywords):
            best = f
            break
    return best

def _is_header_row(vals):
    joined = " ".join([str(v).lower() for v in vals if v is not None])
    joined_n = norm_col(joined)
    keys = ["بند", "code", "hs", "وصف", "description", "brick", "item", "goods", "وارد", "duty", "قيمة"]
    for k in keys:
        if norm_col(k) in joined_n:
            return True
    return False

def _rows_to_dicts(raw_rows):
    # raw_rows: list of lists, first find header row
    header_idx = 0
    for i, r in enumerate(raw_rows[:12]):
        if _is_header_row(r):
            header_idx = i
            break
    headers = [norm_col(c) for c in raw_rows[header_idx]]
    out = []
    for r in raw_rows[header_idx+1:]:
        d = {}
        for h, v in zip(headers, r):
            if not h: continue
            d[h] = "" if v is None else str(v)
        # تجاهل سطور فاضية تماما
        if any(str(v).strip() != "" for v in d.values()):
            out.append(d)
    return out

def load_table_file(path):
    try:
        if path.lower().endswith(".csv"):
            import csv
            with open(path, encoding="utf-8-sig", errors="ignore") as f:
                reader = list(csv.reader(f))
            return _rows_to_dicts(reader)
        else:
            try:
                import pandas as pd
                df = pd.read_excel(path, dtype=str, header=None)
                df = df.fillna("")
                raw = df.values.tolist()
                return _rows_to_dicts(raw)
            except Exception as e:
                print("pandas fail", path, e, flush=True)
                import openpyxl
                wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
                ws = wb.active
                rows = list(ws.iter_rows(values_only=True))
                return _rows_to_dicts(rows)
    except Exception as e:
        print("load file error", path, e, flush=True)
        return []
    return []

def pick(row, *candidates):
    for c in candidates:
        nc = norm_col(c)
        if nc in row:
            v = row[nc]
            if v is not None and str(v).strip() != "":
                return str(v).strip()
    return ""

def pick_code_fallback(row):
    code = pick(row, "code", "hs", "hscode", "hs_code", "بند", "itemcode", "tariffcode", "hscode6", "hscode4", "tariff", "chapter", "heading", "customscode", "البندالجمركي", "البند")
    if code:
        return code
    for v in row.values():
        vv = str(v).strip().replace(".", "").replace(" ", "")
        if vv.isdigit() and len(vv) >= 4 and len(vv) <= 10:
            return vv
    return ""

def init_db():
    # امسح الداتا القديمة البايظة خالص وابني من الاول
    try:
        if os.path.exists("items.db"):
            os.remove("items.db")
            print("deleted old items.db", flush=True)
    except Exception as e:
        print("delete db error", e, flush=True)
    conn = sqlite3.connect("items.db")
    cur = conn.cursor()
    cur.execute("CREATE TABLE IF NOT EXISTS tariffs (code TEXT PRIMARY KEY, desc_ar TEXT, desc_en TEXT, duty REAL, vat REAL, profit REAL, dev REAL, table_rate REAL, dumping REAL, notes TEXT)")
    cur.execute("CREATE TABLE IF NOT EXISTS items (code TEXT, desc_ar TEXT, desc_en TEXT)")
    cur.execute("CREATE TABLE IF NOT EXISTS gs1 (brick_code TEXT, brick_name TEXT, segment TEXT, family TEXT, class_name TEXT, includes_text TEXT, excludes_text TEXT)")
    cur.execute("CREATE TABLE IF NOT EXISTS invoices (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, supplier TEXT, importer TEXT, origin TEXT, currency TEXT, incoterm TEXT, total_foreign REAL, usd_rate REAL, cif_egp REAL, duty_egp REAL, vat_egp REAL, profit_egp REAL, total_tax_egp REAL, grand_total_egp REAL, created_at TEXT)")
    conn.commit()
    tariff_path = find_file("tariff")
    print("tariff file:", tariff_path, flush=True)
    if tariff_path:
        rows = load_table_file(tariff_path)
        print(f"tariff rows {len(rows)}", flush=True)
        if rows:
            print("tariff headers:", list(rows[0].keys())[:12], flush=True)
            print("tariff sample:", rows[0], flush=True)
        for r in rows:
            code = pick_code_fallback(r)
            if not code: continue
            code = re.sub(r'\.0$', '', code.strip())
            desc_ar = pick(r, "desc_ar", "descriptionar", "arabic", "ar", "الوصف", "وصف", "descriptionarabic", "desc", "goods", "item", "البيان", "بيان", "الصنف")
            desc_en = pick(r, "desc_en", "descriptionen", "english", "en", "description")
            def num(*names):
                v = pick(r, *names)
                try: return float(str(v).replace("%","").replace(",","").strip() or 0)
                except: return 0
            duty = num("duty", "وارد", "customs", "duty%", "rate", "customsduty", "فئةالوارد", "فئة")
            vat = num("vat", "قيمةمضافة", "valueadded", "vat%")
            if vat == 0: vat = 14
            profit = num("profit", "ارباح", "ارباح تجارية", "profit%", "at", "profits", "ارباح تجاريه")
            dev = num("dev", "تنمية", "development", "dev%", "developmenttax", "رسم تنمية")
            table_rate = num("table", "جدول", "tabletax", "table%", "ضريبةجدول")
            dumping = num("dumping", "اغراق", "dumping%")
            notes = pick(r, "notes", "note", "ملاحظات", "remarks", "ملحوظات")
            cur.execute("INSERT OR REPLACE INTO tariffs VALUES (?,?,?,?,?,?,?,?,?,?)", (code, desc_ar, desc_en, duty, vat, profit, dev, table_rate, dumping, notes))
        conn.commit()
    book_path = find_file("book") or find_file("logist") or find_file("tattef")
    print("book file:", book_path, flush=True)
    if book_path:
        rows = load_table_file(book_path)
        print(f"book rows {len(rows)}", flush=True)
        if rows:
            print("book headers:", list(rows[0].keys())[:12], flush=True)
            print("book sample:", rows[0], flush=True)
        for r in rows:
            code = pick_code_fallback(r)
            if not code: continue
            code = re.sub(r'\.0$', '', code.strip())
            desc_ar = pick(r, "desc_ar", "descriptionar", "arabic", "الوصف", "desc", "item", "namear", "البيان", "بيان", "الصنف")
            desc_en = pick(r, "desc_en", "descriptionen", "english", "nameen", "name")
            if not desc_ar and not desc_en: continue
            cur.execute("INSERT INTO items VALUES (?,?,?)", (code, desc_ar, desc_en))
        conn.commit()
    cur.execute("SELECT COUNT(*) FROM items")
    if cur.fetchone()[0] == 0:
        cur.execute("INSERT INTO items SELECT code, desc_ar, desc_en FROM tariffs")
        conn.commit()
    gs1_path = find_file("gs1")
    print("gs1 file:", gs1_path, flush=True)
    if gs1_path:
        rows = load_table_file(gs1_path)
        print(f"gs1 rows {len(rows)}", flush=True)
        if rows:
            print("gs1 headers:", list(rows[0].keys())[:12], flush=True)
        cur.execute("DELETE FROM gs1")
        for r in rows:
            brick_code = pick(r, "brick_code", "brickcode", "brick", "code")
            brick_name = pick(r, "brick_name", "brickname", "brick", "name", "title")
            if not brick_code and not brick_name: continue
            segment = pick(r, "segment", "segmentname")
            family = pick(r, "family", "familyname")
            class_name = pick(r, "class_name", "classname", "class")
            includes = pick(r, "includes_text", "includes", "include")
            excludes = pick(r, "excludes_text", "excludes", "exclude")
            cur.execute("INSERT INTO gs1 VALUES (?,?,?,?,?,?,?)", (brick_code, brick_name, segment, family, class_name, includes, excludes))
        conn.commit()
    cur.execute("SELECT COUNT(*) FROM tariffs")
    print("DB built tariffs:", cur.fetchone()[0], flush=True)
    cur.execute("SELECT COUNT(*) FROM gs1")
    print("DB built gs1:", cur.fetchone()[0], flush=True)
    cur.execute("SELECT COUNT(*) FROM items")
    print("DB built items:", cur.fetchone()[0], flush=True)
    conn.close()

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
        cur.execute("SELECT code, desc_ar, desc_en FROM items WHERE code LIKE ? OR desc_ar LIKE ? OR desc_en LIKE ? LIMIT 5", ("%" + text + "%", "%" + text + "%", "%" + text + "%"))
        rows = cur.fetchall()
        conn.close()
        for r in rows:
            code = r[0]
            t = get_tariff(code)
            d = t.get("duty", 0) if t else 0
            out.append({"code": code, "clean": str(code).replace(".", ""), "desc_ar": r[1] or "", "desc_en": r[2] or "", "duty": d})
    except Exception as e:
        print("search error", e, flush=True)
    return out, []

def search_gs1(q):
    q = q.strip()
    out = []
    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute("SELECT brick_code, brick_name, segment, family, class_name, includes_text, excludes_text FROM gs1 WHERE brick_code LIKE ? OR brick_name LIKE ? OR family LIKE ? OR class_name LIKE ? OR includes_text LIKE ? LIMIT 3", ("%" + q + "%", "%" + q + "%", "%" + q + "%", "%" + q + "%", "%" + q + "%"))
        rows = cur.fetchall()
        conn.close()
        for r in rows:
            out.append(r)
    except Exception as e:
        print("gs1 error", e, flush=True)
    return out

def get_usd_rate():
    try:
        r = requests.get("https://open.er-api.com/v6/latest/USD", timeout=10)
        return float(r.json()["rates"].get("EGP", 48.0))
    except Exception:
        return 48.0

def clean_ai_text(txt):
    txt = txt.strip()
    txt = txt.replace(chr(96), "").strip()
    low = txt.lower()
    if low.startswith("json"):
        txt = txt[4:].strip()
    return txt

def read_invoice_image(image_bytes):
    try:
        import base64
        api_key = os.environ.get("GEMINI_KEY", "") or os.environ.get("GEMINI_API_KEY", "") or os.environ.get("GOOGLE_API_KEY", "")
        if not api_key:
            print("GEMINI_KEY missing", flush=True)
            return None
        b64 = base64.b64encode(image_bytes).decode("utf-8")
        prompt = "اقرا الفاتورة وطلع JSON بس بدون اي كلام زيادة: {\"supplier\":\"\",\"importer\":\"\",\"origin\":\"\",\"currency\":\"USD\",\"incoterm\":\"FOB\",\"total\":0,\"items\":[{\"name\":\"\",\"code\":\"\",\"qty\":0,\"price\":0,\"total\":0}]} لو مش شايف كود اكتب code فارغ. لو الصورة مش فاتورة رجع {\"supplier\":\"\",\"total\":0,\"items\":[]}"
        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={api_key}"
        payload = {"contents": [{"parts": [{"text": prompt}, {"inline_data": {"mime_type": "image/jpeg", "data": b64}}]}]}
        r = requests.post(url, json=payload, timeout=60)
        print("gemini status", r.status_code, flush=True)
        if r.status_code != 200:
            print("gemini error", r.text[:500], flush=True)
            return None
        j = r.json()
        txt = j["candidates"][0]["content"]["parts"][0]["text"]
        txt = clean_ai_text(txt)
        s = txt.find("{"); e = txt.rfind("}")
        if s != -1 and e != -1:
            txt = txt[s:e+1]
        return json.loads(txt)
    except Exception as e:
        print("OCR error", e, flush=True)
        return None

def calc_invoice_egp(items, usd_rate):
    cif_egp = 0; duty_egp = 0; vat_egp = 0; profit_egp = 0; dev_egp = 0
    details = []
    for it in items:
        code = str(it.get("code", "")).replace(".", "")
        tot = float(it.get("total", 0) or 0)
        cif = tot * usd_rate
        t = get_tariff(code) if code else None
        duty_p = float(t.get("duty", 0)) if t else 0
        vat_p = float(t.get("vat", 14)) if t else 14
        profit_p = float(t.get("profit", 0) or 0) if t else 0
        dev_p = float(t.get("dev", 0) or 0) if t else 0
        duty, vat, profit, tot2, dev, tb, dm = calc_customs(cif, duty_p, vat_p, profit_p, dev_p, 0, 0)
        cif_egp += cif; duty_egp += duty; vat_egp += vat; profit_egp += profit; dev_egp += dev
        details.append({"name": it.get("name", ""), "code": code, "cif_egp": cif, "duty_p": duty_p, "duty": duty})
    total_tax = duty_egp + vat_egp + profit_egp + dev_egp
    grand = cif_egp + total_tax
    return {"cif_egp": cif_egp, "duty_egp": duty_egp, "vat_egp": vat_egp, "profit_egp": profit_egp, "dev_egp": dev_egp, "total_tax_egp": total_tax, "grand_total_egp": grand, "details": details}

def format_invoice_preview(inv):
    msg = "📋 مراجعة الفاتورة:\n\n"
    msg += "المورد: " + str(inv.get("supplier", "")) + "\n"
    msg += "المنشأ: " + str(inv.get("origin", "")) + "\n"
    msg += "القيمة: " + fmt0(inv.get("total_foreign", 0)) + " " + str(inv.get("currency", "USD")) + "\n"
    msg += "التسليم: " + str(inv.get("incoterm", "FOB")) + "\n"
    msg += "سعر الدولار: " + fmt2(inv.get("usd_rate", 0)) + " ج\n\n"
    for it in inv.get("items", []):
        msg += "• " + str(it.get("name", "")) + " | بند " + str(it.get("code", "")) + " | " + str(it.get("qty", 0)) + " × " + str(it.get("price", 0)) + " = " + fmt0(it.get("total", 0)) + "\n"
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
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        cur.execute("INSERT INTO invoices (user_id, supplier, importer, origin, currency, incoterm, total_foreign, usd_rate, cif_egp, duty_egp, vat_egp, profit_egp, total_tax_egp, grand_total_egp, created_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", (user_id, data.get("supplier", ""), data.get("importer", ""), data.get("origin", ""), data.get("currency", ""), data.get("incoterm", ""), data.get("total_foreign", 0), data.get("usd_rate", 0), data.get("cif_egp", 0), data.get("duty_egp", 0), data.get("vat_egp", 0), data.get("profit_egp", 0), data.get("total_tax_egp", 0), data.get("grand_total_egp", 0), now))
        conn.commit(); conn.close()
    except Exception as e:
        print("save invoice error", e, flush=True)

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
    if not ok: return
    msg = ("اهلا بيك في بوت Logistics Group\n\n8471 -> هيجيب البند\nزجاج -> هيجيب البند بالاسم\n8471 100000 -> هيحسبلك الجمارك\nمنشور 28 -> هيجيب المنشور\n/gpc paint -> هيجيب تصنيف GS1\n📸 ابعت صورة فاتورة -> هقراها وأحسبها\n/invoices -> آخر الفواتير")
    await update.message.reply_text(msg)

async def gpc_command(update, context):
    ok = await request_approval(update, context)
    if not ok: return
    if not context.args:
        await update.message.reply_text("اكتب كده:\n/gpc paint\n/gpc milk")
        return
    q = " ".join(context.args).strip()
    rows = search_gs1(q)
    if not rows:
        await update.message.reply_text("❌ مفيش Brick لـ: " + q)
        return
    for r in rows:
        msg = "🧱 GS1 Brick: " + str(r[1]) + "\n🔢 Brick Code: " + str(r[0]) + "\n📂 Segment: " + str(r[2]) + "\n📁 Family: " + str(r[3]) + "\n📄 Class: " + str(r[4]) + "\n"
        if r[5]: msg += "\n✅ Includes: " + str(r[5])[:350] + "\n"
        if r[6]: msg += "\n⛔ Excludes: " + str(r[6])[:350]
        await update.message.reply_text(msg[:4000])

async def invoices_command(update, context):
    ok = await request_approval(update, context)
    if not ok: return
    try:
        conn = sqlite3.connect("items.db")
        cur = conn.cursor()
        cur.execute("SELECT supplier, total_foreign, currency, total_tax_egp, grand_total_egp, created_at FROM invoices WHERE user_id=? ORDER BY id DESC LIMIT 10", (update.effective_user.id,))
        rows = cur.fetchall(); conn.close()
    except Exception:
        rows = []
    if not rows:
        await update.message.reply_text("لسه مفيش فواتير.\nابعت صورة فاتورة عشان نبدأ.")
        return
    msg = "🧾 آخر الفواتير:\n\n"
    for r in rows:
        msg += "• " + str(r[0] or "بدون مورد") + " | " + fmt0(r[1]) + " " + str(r[2]) + " | ضرائب: " + fmt0(r[3]) + " ج | إجمالي: " + fmt0(r[4]) + " ج | " + str(r[5]) + "\n"
    await update.message.reply_text(msg[:4000])

async def handle_photo(update, context):
    ok = await request_approval(update, context)
    if not ok: return
    await update.message.reply_text("📸 وصلت صورة الفاتورة، بقراها... استنى ثواني")
    try:
        photo = update.message.photo[-1]
        f = await context.bot.get_file(photo.file_id)
        image_bytes = await f.download_as_bytearray()
        image_bytes = bytes(image_bytes)
        data = read_invoice_image(image_bytes)
        if not data:
            await update.message.reply_text("⚠️ مش قادر أقرا الصورة أوتوماتيك.\nابعت البيانات يدوي:\nفاتورة: المورد | المنشأ | القيمة | العملة | FOB")
            return
        usd_rate = get_usd_rate()
        items = []
        for it in data.get("items", []) or []:
            qty = float(it.get("qty", 0) or 0); price = float(it.get("price", 0) or 0); total = float(it.get("total", 0) or qty * price)
            items.append({"name": str(it.get("name", "")), "code": str(it.get("code", "")).replace(".", ""), "qty": qty, "price": price, "total": total})
        inv = {"supplier": str(data.get("supplier", "")), "importer": str(data.get("importer", "")), "origin": str(data.get("origin", "")), "currency": str(data.get("currency", "USD")).upper() or "USD", "incoterm": str(data.get("incoterm", "FOB")).upper() or "FOB", "total_foreign": float(data.get("total", 0) or 0), "usd_rate": usd_rate, "items": items}
        pending_invoices[update.effective_user.id] = inv
        await update.message.reply_text(format_invoice_preview(inv)[:4000], reply_markup=invoice_confirm_keyboard())
    except Exception as e:
        print("Photo error", e, flush=True)
        await update.message.reply_text("حصل مشكلة في قراءة الصورة. ابعت يدوي:\nفاتورة: المورد | المنشأ | القيمة | العملة | FOB")

async def handle_invoice_text(text, update):
    try:
        parts = text.replace("فاتورة:", "").replace("فاتورة", "").strip().split("|")
        supplier = parts[0].strip() if len(parts) > 0 else ""
        origin = parts[1].strip() if len(parts) > 1 else ""
        total = 0
        if len(parts) > 2: total = float(parts[2].strip().replace(",", ""))
        currency = "USD"
        if len(parts) > 3: currency = parts[3].strip().upper()
        incoterm = "FOB"
        if len(parts) > 4: incoterm = parts[4].strip().upper()
    except Exception:
        await update.message.reply_text("صيغة الفاتورة غلط. ابعت كده:\nفاتورة: Rollmed | الصين | 261152 | USD | FOB")
        return
    usd_rate = get_usd_rate()
    pending_invoices[update.effective_user.id] = {"supplier": supplier, "importer": "", "origin": origin, "total_foreign": total, "currency": currency, "incoterm": incoterm, "usd_rate": usd_rate, "items": []}
    msg = "📋 قريت الفاتورة:\n\nالمورد: " + supplier + "\nالمنشأ: " + origin + "\nالقيمة: " + fmt0(total) + " " + currency + "\nالتسليم: " + incoterm + "\nسعر الدولار: " + fmt2(usd_rate) + " ج\n\nدلوقتي ابعت الأصناف كده:\nصنف: اسم الصنف | البند | الكمية | السعر\nمثال: صنف: أقراص طحن | 6804.22 | 1000 | 50"
    await update.message.reply_text(msg, reply_markup=invoice_confirm_keyboard())

async def handle_smart(update, context):
    ok = await request_approval(update, context)
    if not ok: return
    text = update.message.text.strip()
    low = text.lower()
    if low.startswith("فاتورة"):
        await handle_invoice_text(text, update); return
    if low.startswith("صنف:") or low.startswith("صنف"):
        uid = update.effective_user.id
        if uid not in pending_invoices:
            await update.message.reply_text("ابعت الفاتورة الأول:\nفاتورة: المورد | المنشأ | القيمة | العملة | FOB"); return
        try:
            p = text.replace("صنف:", "").replace("صنف", "").strip().split("|")
            name = p[0].strip(); code = ""
            if len(p) > 1: code = p[1].strip().replace(".", "")
            qty = float(p[2].strip()) if len(p) > 2 else 0
            price = float(p[3].strip()) if len(p) > 3 else 0
            pending_invoices[uid]["items"].append({"name": name, "code": code, "qty": qty, "price": price, "total": qty * price})
            await update.message.reply_text("✅ أضفت: " + name + " | بند " + code + " | " + str(qty) + " × " + str(price), reply_markup=invoice_confirm_keyboard())
        except Exception:
            await update.message.reply_text("صيغة الصنف غلط. ابعت كده:\nصنف: أقراص طحن | 6804.22 | 1000 | 50")
        return
    if low.startswith("gpc ") or low.startswith("gs1 "):
        q = text.replace("gpc", "").replace("gs1", "").replace("GS1", "").strip()
        if not q:
            await update.message.reply_text("اكتب كده:\n/gpc paint"); return
        rows = search_gs1(q)
        if not rows:
            await update.message.reply_text("❌ مفيش Brick لـ: " + q); return
        for r in rows:
            msg = "🧱 GS1 Brick: " + str(r[1]) + "\n🔢 Brick Code: " + str(r[0]) + "\n📂 Segment: " + str(r[2]) + "\n📁 Family: " + str(r[3]) + "\n📄 Class: " + str(r[4])
            await update.message.reply_text(msg[:4000])
        return
    if "منشور" in low or "manshour" in low or "قرار" in low:
        num = "".join([c for c in text if c.isdigit()])
        if not num:
            await update.message.reply_text("اكتب رقم المنشور: مثال منشور 28"); return
        await update.message.reply_text("بدور على منشور " + num + " ...")
        try:
            q = "منشور جمارك " + num + " نافذة"
            headers = {"User-Agent": "Mozilla/5.0"}
            r = requests.get("https://www.google.com/search?q=" + q, headers=headers, timeout=12)
            soup = BeautifulSoup(r.text, "html.parser")
            found = []
            for a in soup.find_all("a", href=True):
                t = a.get_text(strip=True)
                if len(t) > 20 and num in t: found.append("• " + t[:120])
                if len(found) >= 5: break
            if found:
                msg = "📜 منشور " + num + " - نتائج:\n\n" + "\n\n".join(found) + "\n\n🔗 نافذة:\nhttps://www.nafeza.gov.eg/"
            else:
                msg = "📜 منشور " + num + "\n\nشوفه هنا:\nhttps://www.nafeza.gov.eg/\nhttps://www.google.com/search?q=منشور+جمارك+" + num
            await update.message.reply_text(msg)
        except Exception:
            await update.message.reply_text("📜 منشور " + num + "\nhttps://www.nafeza.gov.eg/")
        return
    parts = text.replace(",", " ").split()
    if len(parts) == 2:
        if parts[0].replace(".", "").isdigit():
            if parts[1].replace(".", "").isdigit():
                code = parts[0].replace(".", ""); cif = float(parts[1])
                tariff = get_tariff(code)
                if not tariff:
                    await update.message.reply_text("البند " + code + " مش موجود"); return
                duty_p = float(tariff.get("duty", 0)); vat_p = float(tariff.get("vat", 14)); profit_p = float(tariff.get("profit") or 1.0); dev_p = float(tariff.get("dev", 0)); table_p = float(tariff.get("table_rate", 0) or tariff.get("table", 0) or 0); dump_p = float(tariff.get("dumping", 0))
                duty, vat, profit, total, dev, tb, dm = calc_customs(cif, duty_p, vat_p, profit_p, dev_p, table_p, dump_p)
                msg = "🧾 حساب - البند: " + code + "\n"
                if tariff.get("desc_ar"): msg += clean_desc(tariff.get("desc_ar")) + "\n"
                if tariff.get("desc_en"): msg += clean_desc(tariff.get("desc_en")) + "\n"
                msg += "قيمة CIF: " + fmt0(cif) + " ج\n\nوارد " + str(duty_p) + "% = " + fmt0(duty) + " ج\nرسم تنمية " + str(dev_p) + "% = " + fmt0(dev) + " ج\n"
                if table_p: msg += "ضريبة جدول " + str(table_p) + "% = " + fmt0(tb) + " ج\n"
                if dump_p: msg += "رسم إغراق " + str(dump_p) + "% = " + fmt0(dm) + " ج\n"
                msg += "قيمة مضافة " + str(vat_p) + "% = " + fmt0(vat) + " ج\n"
                if profit_p: msg += "ارباح تجارية " + str(profit_p) + "% = " + fmt0(profit) + " ج\n"
                msg += "----------------\nاجمالي تقديري: " + fmt0(total) + " ج\nالاجمالي + البضاعة: " + fmt0(cif + total) + " ج\n"
                if tariff.get("notes"): msg += "\nملاحظة: " + str(tariff.get("notes", ""))
                await update.message.reply_text(msg); return
    ex_found, _ = search_combined(text)
    if not ex_found:
        await update.message.reply_text("مفيش نتائج لـ: " + text + "\nجرب: منشور 28 - او - 8471 100000"); return
    it = ex_found[0]
    tariff = get_tariff(it["clean"])
    msg = "📦 البند: " + str(it["code"]) + "\n" + clean_desc(it.get("desc_ar", "")) + "\n" + clean_desc(it.get("desc_en", "")) + "\n"
    if tariff:
        msg += "\n💰 وارد: " + str(tariff.get("duty", 0)) + "% | تنمية: " + str(tariff.get("dev", 0)) + "% | قيمة مضافة: " + str(tariff.get("vat", 14)) + "% | ارباح: " + str(tariff.get("profit", 0)) + "%\n"
        if tariff.get("notes"): msg += "📋 " + str(tariff.get("notes", "")) + "\n"
    else:
        msg += "\n💰 النسب: مش موجودة\n"
    msg += "\n🔗 التأكد من نافذة:\nhttps://www.nafeza.gov.eg/\n\n💡 عشان تحسبها ابعت:\n" + it["clean"] + " 100000"
    if len(ex_found) > 1:
        msg += "\n\n🔍 مشابهة:\n"
        for it2 in ex_found[1:4]:
            msg += "- [" + str(it2["code"]) + "] " + str(it2.get("desc_ar", ""))[:40] + " | " + str(it2.get("desc_en", ""))[:40] + " | فئة: " + str(it2.get("duty", 0)) + "%\n"
    await update.message.reply_text(msg[:4000])

async def button_callback(update, context):
    q = update.callback_query
    await q.answer()
    if q.data in ["inv_confirm", "inv_edit"]:
        uid = q.from_user.id
        if uid not in pending_invoices:
            await q.edit_message_text("مفيش فاتورة معلقة، ابعت فاتورة جديدة."); return
        if q.data == "inv_edit":
            await q.edit_message_text("ابعت التعديل كده:\nفاتورة: المورد | المنشأ | القيمة | العملة | FOB"); return
        inv = pending_invoices[uid]
        if not inv.get("items"):
            inv["items"] = [{"name": "إجمالي الفاتورة", "code": "", "qty": 1, "price": inv["total_foreign"], "total": inv["total_foreign"]}]
        result = calc_invoice_egp(inv["items"], inv["usd_rate"])
        save_invoice(uid, {"supplier": inv["supplier"], "importer": inv.get("importer", ""), "origin": inv["origin"], "currency": inv["currency"], "incoterm": inv["incoterm"], "total_foreign": inv["total_foreign"], "usd_rate": inv["usd_rate"], "cif_egp": result["cif_egp"], "duty_egp": result["duty_egp"], "vat_egp": result["vat_egp"], "profit_egp": result["profit_egp"], "total_tax_egp": result["total_tax_egp"], "grand_total_egp": result["grand_total_egp"], "items": inv["items"]})
        msg = "✅ تم التأكيد - الحسبة بالمصري\n\nالمورد: " + str(inv["supplier"]) + "\nالمنشأ: " + str(inv["origin"]) + "\nسعر الدولار: " + fmt2(inv["usd_rate"]) + " ج\n\n"
        for d in result["details"]:
            msg += "• " + str(d["name"]) + " | بند " + str(d["code"]) + " | CIF: " + fmt0(d["cif_egp"]) + " ج | وارد " + str(d["duty_p"]) + "% = " + fmt0(d["duty"]) + " ج\n"
        msg += "\nCIF: " + fmt0(result["cif_egp"]) + " ج\nالوارد: " + fmt0(result["duty_egp"]) + " ج\nقيمة مضافة: " + fmt0(result["vat_egp"]) + " ج\nأ.ت.ص: " + fmt0(result["profit_egp"]) + " ج\nتنمية: " + fmt0(result["dev_egp"]) + " ج\n----------------\nإجمالي الضرائب: " + fmt0(result["total_tax_egp"]) + " ج\nالإجمالي + البضاعة: " + fmt0(result["grand_total_egp"]) + " ج\n\n💾 اتحفظت في /invoices"
        pending_invoices.pop(uid, None)
        await q.edit_message_text(msg[:4000]); return
    if q.from_user.id != ADMIN_ID:
        await q.edit_message_text("لست الادمن"); return
    data = q.data
    if data.startswith("approve_"):
        uid = int(data.split("_")[1]); allowed_users.add(uid); save_json(ALLOWED_FILE, list(allowed_users)); pending_users.pop(str(uid), None); save_json(PENDING_FILE, pending_users)
        await q.edit_message_text("تمت الموافقة على " + str(uid))
        try: await context.bot.send_message(chat_id=uid, text="تمت الموافقة! ابعت /start")
        except Exception: pass
    else:
        uid = data.split("_")[1]; pending_users.pop(uid, None); save_json(PENDING_FILE, pending_users); await q.edit_message_text("تم رفض " + uid)

async def admin_list(update, context):
    if update.effective_user.id != ADMIN_ID: return
    text = "👥 المصرح لهم:\n"
    for uid in allowed_users: text += "• " + str(uid) + "\n"
    text += "\n⏳ المنتظرين:\n"
    for pid, info in pending_users.items(): text += "• " + str(pid) + " - " + str(info.get("name", "")) + " @" + str(info.get("username", "")) + "\n"
    text += "\nعشان تشيل حد ابعت:\n/remove + الـ ID"
    await update.message.reply_text(text)

async def remove_user(update, context):
    if update.effective_user.id != ADMIN_ID: return
    if not context.args:
        await update.message.reply_text("ابعت كده:\n/remove 123456789"); return
    try: uid = int(context.args[0])
    except Exception:
        await update.message.reply_text("الـ ID لازم يكون رقم"); return
    allowed_users.discard(uid); pending_users.pop(str(uid), None); save_json(ALLOWED_FILE, list(allowed_users)); save_json(PENDING_FILE, pending_users)
    await update.message.reply_text("✅ تمت إزالة " + str(uid))
    try: await context.bot.send_message(chat_id=uid, text="تمت إزالتك من البوت، ابعت /start تاني")
    except Exception: pass

if TOKEN:
    print("Bot polling started...", flush=True)
    init_db()
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