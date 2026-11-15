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
    try: return format(float(x), ",.0f")
    except Exception: return "0"
def fmt2(x):
    try: return format(float(x), ",.2f")
    except Exception: return "0.00"

allowed_users = set(load_json(ALLOWED_FILE, []))
pending_users = load_json(PENDING_FILE, {})
pending_invoices = {}

def clean_desc(s):
    return str(s or "").strip().strip('"').strip()

def find_file(*keywords):
    files = glob.glob("*") + glob.glob("**/*", recursive=True)
    files = [f for f in files if os.path.isfile(f)]
    for f in files:
        low = f.lower()
        if all(k.lower() in low for k in keywords):
            return f
    return None

def parse_duty(raw):
    s = str(raw or "").strip()
    if not s: return 0, ""
    if "ممنوع" in s: return -1, s
    if "معف" in s: return 0, "معفاة"
    m = re.search(r'([\d\.]+)', s.replace(",", "."))
    if m:
        try: return float(m.group(1)), s
        except: return 0, s
    return 0, s

def short_code(code12):
    c = str(code12).strip().replace(".", "")
    while len(c) > 4 and c.endswith("00"):
        c = c[:-2]
    if len(c) == 4: return c
    if len(c) == 6: return f"{c[0:4]}.{c[4:6]}"
    if len(c) == 8: return f"{c[0:4]}.{c[4:6]}.{c[6:8]}"
    if len(c) == 10: return f"{c[0:4]}.{c[4:6]}.{c[6:8]}.{c[8:10]}"
    return c

def init_db():
    try:
        if os.path.exists("items.db"):
            os.remove("items.db")
            print("deleted old items.db", flush=True)
    except Exception as e:
        print("delete db error", e, flush=True)
    conn = sqlite3.connect("items.db")
    cur = conn.cursor()
    cur.execute("CREATE TABLE tariffs (code TEXT PRIMARY KEY, code12 TEXT, desc_ar TEXT, desc_en TEXT, duty REAL, duty_raw TEXT, procedures TEXT)")
    cur.execute("CREATE TABLE book_items (code TEXT, desc_ar TEXT)")
    cur.execute("CREATE TABLE gs1 (brick_code TEXT PRIMARY KEY, brick_name TEXT, segment TEXT, family TEXT, class_name TEXT, includes_text TEXT, excludes_text TEXT)")
    cur.execute("CREATE TABLE invoices (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, supplier TEXT, importer TEXT, origin TEXT, currency TEXT, incoterm TEXT, total_foreign REAL, usd_rate REAL, cif_egp REAL, duty_egp REAL, vat_egp REAL, profit_egp REAL, total_tax_egp REAL, grand_total_egp REAL, created_at TEXT)")
    conn.commit()
    # 1 - التعريفة - الملف الاساسي للنسب
    tariff_path = find_file("tariff")
    print("tariff file:", tariff_path, flush=True)
    if tariff_path:
        try:
            import pandas as pd
            try:
                df = pd.read_excel(tariff_path, dtype=str)
            except Exception:
                df = pd.read_csv(tariff_path, dtype=str, encoding="utf-8-sig")
            df = df.fillna("")
            print(f"tariff shape {df.shape} cols {list(df.columns)}", flush=True)
            for _, row in df.iterrows():
                vals = [str(x).strip() for x in row.tolist()]
                if len(vals) < 4: continue
                code_raw = vals[0].replace(".", "").replace(" ", "").strip()
                if not code_raw.isdigit(): continue
                code12 = code_raw.ljust(12, "0")[:12]
                desc_ar = clean_desc(vals[1])
                desc_en = clean_desc(vals[2])
                duty, duty_raw = parse_duty(vals[3])
                proc = vals[4] if len(vals) > 4 else ""
                cur.execute("INSERT OR REPLACE INTO tariffs VALUES (?,?,?,?,?,?,?)", (code12, code12, desc_ar, desc_en, duty, duty_raw, proc))
            conn.commit()
        except Exception as e:
            print("tariff load error", e, flush=True)
    # 2 - Book1 - للبحث فقط - عمود واحد نصي
    book_path = find_file("book")
    print("book file:", book_path, flush=True)
    if book_path:
        try:
            import pandas as pd
            df = pd.read_excel(book_path, dtype=str, header=None)
            df = df.fillna("")
            print(f"book shape {df.shape}", flush=True)
            count_book = 0
            for _, row in df.iterrows():
                txt = str(row.tolist()[0] or "").strip()
                if not txt: continue
                # دور على كود زي 8471.30 او 8471.49 او 8471.70 في اول السطر
                m = re.search(r'(\d{4}(?:\.\d{2}){1,3})', txt)
                if not m: continue
                code_dot = m.group(1)
                code_digits = code_dot.replace(".", "")
                # الوصف = اللي بعد الكود
                desc = txt[m.end():].strip()
                desc = re.sub(r'^[-–—\s]+', '', desc).strip()
                if not desc: continue
                # شيل كلمة معفاه من الوصف وخليها معلومة لوحدها لو حابب
                cur.execute("INSERT INTO book_items VALUES (?,?)", (code_digits, desc))
                count_book += 1
            conn.commit()
            print(f"book parsed items: {count_book}", flush=True)
        except Exception as e:
            print("book load error", e, flush=True)
    # 3 - GS1 - 5249 بريك من غير تكرار
    gs1_path = find_file("gs1")
    print("gs1 file:", gs1_path, flush=True)
    if gs1_path:
        try:
            import pandas as pd
            df = pd.read_excel(gs1_path, dtype=str)
            df = df.fillna("")
            print(f"gs1 raw shape {df.shape} cols {list(df.columns)}", flush=True)
            df = df.drop_duplicates(subset=["BrickCode"])
            print(f"gs1 distinct bricks {len(df)}", flush=True)
            for _, row in df.iterrows():
                bc = str(row.get("BrickCode", "")).strip()
                bn = str(row.get("BrickTitle", "")).strip()
                if not bc and not bn: continue
                cur.execute("INSERT OR REPLACE INTO gs1 VALUES (?,?,?,?,?,?,?)", (bc, bn, str(row.get("SegmentTitle", "")), str(row.get("FamilyTitle", "")), str(row.get("ClassTitle", "")), str(row.get("BrickDefinition_Includes", "")), str(row.get("BrickDefinition_Excludes", ""))))
            conn.commit()
        except Exception as e:
            print("gs1 load error", e, flush=True)
    cur.execute("SELECT COUNT(*) FROM tariffs")
    print("DB built tariffs:", cur.fetchone()[0], flush=True)
    cur.execute("SELECT COUNT(*) FROM book_items")
    print("DB built book_items:", cur.fetchone()[0], flush=True)
    cur.execute("SELECT COUNT(*) FROM gs1")
    print("DB built gs1:", cur.fetchone()[0], flush=True)
    conn.close()

def get_db():
    conn = sqlite3.connect("items.db")
    conn.row_factory = sqlite3.Row
    return conn

def get_tariff(code_input):
    q = str(code_input).replace(".", "").replace(" ", "").strip()
    if not q.isdigit(): return None
    try:
        conn = get_db()
        cur = conn.cursor()
        q12 = q.ljust(12, "0")[:12]
        cur.execute("SELECT * FROM tariffs WHERE code=?", (q12,))
        row = cur.fetchone()
        if not row:
            cur.execute("SELECT * FROM tariffs WHERE code LIKE ? ORDER BY code LIMIT 1", (q + "%",))
            row = cur.fetchone()
        conn.close()
        return dict(row) if row else None
    except Exception as e:
        print("tariff error", e, flush=True)
        return None

def calc_customs(cif, duty_p):
    if duty_p < 0: duty_p = 0
    duty = cif * duty_p / 100
    base = cif + duty
    vat = base * 14 / 100
    profit = base * 1 / 100
    total = duty + vat + profit
    return duty, vat, profit, total

def search_combined(text):
    text = text.strip()
    out = []
    try:
        conn = get_db()
        cur = conn.cursor()
        if text.replace(".", "").replace(" ", "").isdigit():
            q = text.replace(".", "").replace(" ", "")
            cur.execute("SELECT code, desc_ar, desc_en, duty, duty_raw FROM tariffs WHERE code LIKE ? ORDER BY code LIMIT 5", (q + "%",))
            rows = cur.fetchall()
            # لو ملقاش في التعريفة، دور في Book1
            if not rows:
                cur.execute("SELECT code, desc_ar FROM book_items WHERE code LIKE ? LIMIT 5", (q + "%",))
                brow = cur.fetchall()
                for r in brow:
                    t = get_tariff(r[0])
                    duty = t.get("duty", 0) if t else 0
                    duty_raw = t.get("duty_raw", "") if t else ""
                    out.append({"code12": r[0], "code": short_code(r[0]), "desc_ar": r[1] or "", "desc_en": "", "duty": duty, "duty_raw": duty_raw, "from_book": True})
                conn.close()
                return out
        else:
            cur.execute("SELECT code, desc_ar, desc_en, duty, duty_raw FROM tariffs WHERE desc_ar LIKE ? OR desc_en LIKE ? LIMIT 5", ("%" + text + "%", "%" + text + "%"))
            rows = cur.fetchall()
            # كمل من Book1 لو النتائج اقل من 5
            if len(rows) < 5:
                cur.execute("SELECT code, desc_ar FROM book_items WHERE desc_ar LIKE ? LIMIT 5", ("%" + text + "%",))
                brow = cur.fetchall()
                for r in brow:
                    # متكررش اللي موجود في التعريفة
                    if any(str(rr[0]).startswith(str(r[0])[:4]) and str(r[1])[:20] in str(rr[1]) for rr in rows):
                        continue
                    t = get_tariff(r[0])
                    duty = t.get("duty", 0) if t else 0
                    duty_raw = t.get("duty_raw", "") if t else ""
                    out.append({"code12": r[0], "code": short_code(r[0]), "desc_ar": r[1] or "", "desc_en": "", "duty": duty, "duty_raw": duty_raw, "from_book": True})
                    if len(rows) + len(out) >= 5: break
        conn.close()
        for r in rows:
            out.insert(0, {"code12": r[0], "code": short_code(r[0]), "desc_ar": r[1] or "", "desc_en": r[2] or "", "duty": r[3], "duty_raw": r[4], "from_book": False})
        # رتب: التعريفة الاول
        out = sorted(out, key=lambda x: x.get("from_book", False))
        return out[:5]
    except Exception as e:
        print("search error", e, flush=True)
    return out

def search_gs1(q):
    q = q.strip()
    out = []
    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute("SELECT brick_code, brick_name, segment, family, class_name, includes_text, excludes_text FROM gs1 WHERE brick_code LIKE ? OR brick_name LIKE ? OR family LIKE ? OR class_name LIKE ? OR includes_text LIKE ? LIMIT 3", ("%" + q + "%", "%" + q + "%", "%" + q + "%", "%" + q + "%", "%" + q + "%"))
        rows = cur.fetchall()
        conn.close()
        for r in rows: out.append(r)
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
    txt = txt.strip().replace(chr(96), "").strip()
    if txt.lower().startswith("json"): txt = txt[4:].strip()
    return txt

def read_invoice_image(image_bytes):
    try:
        import base64
        api_key = os.environ.get("GEMINI_KEY", "") or os.environ.get("GEMINI_API_KEY", "") or os.environ.get("GOOGLE_API_KEY", "")
        if not api_key:
            print("GEMINI_KEY missing", flush=True)
            return None
        b64 = base64.b64encode(image_bytes).decode("utf-8")
        prompt = "اقرا الفاتورة وطلع JSON بس: {\"supplier\":\"\",\"importer\":\"\",\"origin\":\"\",\"currency\":\"USD\",\"incoterm\":\"FOB\",\"total\":0,\"items\":[{\"name\":\"\",\"code\":\"\",\"qty\":0,\"price\":0,\"total\":0}]} لو مش شايف كود اكتب code فارغ"
        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={api_key}"
        payload = {"contents": [{"parts": [{"text": prompt}, {"inline_data": {"mime_type": "image/jpeg", "data": b64}}]}]}
        r = requests.post(url, json=payload, timeout=60)
        if r.status_code != 200:
            print("gemini error", r.text[:500], flush=True)
            return None
        j = r.json()
        txt = j["candidates"][0]["content"]["parts"][0]["text"]
        txt = clean_ai_text(txt)
        s = txt.find("{"); e = txt.rfind("}")
        if s != -1 and e != -1: txt = txt[s:e+1]
        return json.loads(txt)
    except Exception as e:
        print("OCR error", e, flush=True)
        return None

def calc_invoice_egp(items, usd_rate):
    cif_egp = 0; duty_egp = 0; vat_egp = 0; profit_egp = 0
    details = []
    for it in items:
        code = str(it.get("code", "")).replace(".", "")
        tot = float(it.get("total", 0) or 0)
        cif = tot * usd_rate
        t = get_tariff(code) if code else None
        duty_p = float(t.get("duty", 0)) if t and t.get("duty", 0) >= 0 else 0
        duty, vat, profit, tot2 = calc_customs(cif, duty_p)
        cif_egp += cif; duty_egp += duty; vat_egp += vat; profit_egp += profit
        details.append({"name": it.get("name", ""), "code": code, "cif_egp": cif, "duty_p": duty_p, "duty": duty})
    total_tax = duty_egp + vat_egp + profit_egp
    return {"cif_egp": cif_egp, "duty_egp": duty_egp, "vat_egp": vat_egp, "profit_egp": profit_egp, "total_tax_egp": total_tax, "grand_total_egp": cif_egp + total_tax, "details": details}

def format_invoice_preview(inv):
    msg = "📋 مراجعة الفاتورة:\n\nالمورد: " + str(inv.get("supplier", "")) + "\nالمنشأ: " + str(inv.get("origin", "")) + "\nالقيمة: " + fmt0(inv.get("total_foreign", 0)) + " " + str(inv.get("currency", "USD")) + "\nالتسليم: " + str(inv.get("incoterm", "FOB")) + "\nسعر الدولار: " + fmt2(inv.get("usd_rate", 0)) + " ج\n\n"
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
    if user.id in allowed_users or user.id == ADMIN_ID: return True
    if str(user.id) in pending_users:
        await update.message.reply_text("طلبك قيد المراجعة")
        return False
    pending_users[str(user.id)] = {"name": user.full_name, "username": user.username}
    save_json(PENDING_FILE, pending_users)
    await update.message.reply_text("البوت خاص. تم ارسال طلبك للادمن.")
    btn1 = InlineKeyboardButton(text="موافقة", callback_data="approve_" + str(user.id))
    btn2 = InlineKeyboardButton(text="رفض", callback_data="reject_" + str(user.id))
    kb = InlineKeyboardMarkup([[btn1, btn2]])
    try: await context.bot.send_message(chat_id=ADMIN_ID, text="طلب جديد: " + user.full_name + " ID:" + str(user.id), reply_markup=kb)
    except Exception: pass
    return False

async def start(update, context):
    ok = await request_approval(update, context)
    if not ok: return
    await update.message.reply_text("اهلا بيك في بوت Logistics Group\n\n8471 -> هيجيب البند\nزجاج -> هيجيب البند بالاسم\n8471 100000 -> هيحسبلك الجمارك\nمنشور 28 -> هيجيب المنشور\n/gpc paint -> هيجيب تصنيف GS1\n📸 ابعت صورة فاتورة -> هقراها وأحسبها\n/invoices -> آخر الفواتير")

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
        data = read_invoice_image(bytes(image_bytes))
        if not data:
            await update.message.reply_text("⚠️ مش قادر أقرا الصورة أوتوماتيك.\nابعت يدوي:\nفاتورة: المورد | المنشأ | القيمة | العملة | FOB")
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
    if len(parts) == 2 and parts[0].replace(".", "").isdigit() and parts[1].replace(".", "").isdigit():
        code = parts[0].replace(".", ""); cif = float(parts[1])
        tariff = get_tariff(code)
        if not tariff:
            await update.message.reply_text("البند " + code + " مش موجود"); return
        if tariff.get("duty", 0) < 0:
            await update.message.reply_text("⛔ البند " + short_code(tariff["code"]) + " ممنوع استيراده\n" + clean_desc(tariff.get("desc_ar", "")))
            return
        duty_p = float(tariff.get("duty", 0))
        duty, vat, profit, total = calc_customs(cif, duty_p)
        msg = "🧾 حساب - البند: " + short_code(tariff["code"]) + "\n" + clean_desc(tariff.get("desc_ar")) + "\n" + clean_desc(tariff.get("desc_en")) + "\nقيمة CIF: " + fmt0(cif) + " ج\n\nوارد " + str(duty_p) + "% = " + fmt0(duty) + " ج\nقيمة مضافة 14% = " + fmt0(vat) + " ج\nارباح تجارية 1% = " + fmt0(profit) + " ج\n----------------\nاجمالي تقديري: " + fmt0(total) + " ج\nالاجمالي + البضاعة: " + fmt0(cif + total) + " ج\n"
        if tariff.get("procedures"): msg += "\n📋 إجراءات: " + str(tariff.get("procedures"))
        await update.message.reply_text(msg); return
    ex_found = search_combined(text)
    if not ex_found:
        await update.message.reply_text("مفيش نتائج لـ: " + text + "\nجرب: منشور 28 - او - 8471 100000"); return
    it = ex_found[0]
    if it["duty"] is not None and it["duty"] < 0:
        msg = "⛔ البند: " + str(it["code"]) + " ممنوع استيراده\n" + clean_desc(it.get("desc_ar", "")) + "\n" + clean_desc(it.get("desc_en", "")) + "\n"
    else:
        msg = "📦 البند: " + str(it["code"]) + "\n" + clean_desc(it.get("desc_ar", "")) + "\n" + clean_desc(it.get("desc_en", "")) + "\n\n💰 وارد: " + str(it.get("duty_raw") or str(it.get("duty", 0)) + "%") + " | قيمة مضافة: 14% | ارباح: 1%\n"
        if it.get("from_book"):
            msg += "(من كتاب التعريفة - Book1)\n"
    msg += "\n🔗 التأكد من نافذة:\nhttps://www.nafeza.gov.eg/\n\n💡 عشان تحسبها ابعت:\n" + it["code"] + " 100000"
    if len(ex_found) > 1:
        msg += "\n\n🔍 مشابهة:\n"
        for it2 in ex_found[1:4]:
            msg += "- [" + str(it2["code"]) + "] " + str(it2.get("desc_ar", ""))[:40] + " | " + str(it2.get("desc_en", ""))[:40] + " | وارد: " + str(it2.get("duty_raw") or it2.get("duty")) + "\n"
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
        msg += "\nCIF: " + fmt0(result["cif_egp"]) + " ج\nالوارد: " + fmt0(result["duty_egp"]) + " ج\nقيمة مضافة: " + fmt0(result["vat_egp"]) + " ج\nأ.ت.ص: " + fmt0(result["profit_egp"]) + " ج\n----------------\nإجمالي الضرائب: " + fmt0(result["total_tax_egp"]) + " ج\nالإجمالي + البضاعة: " + fmt0(result["grand_total_egp"]) + " ج\n\n💾 اتحفظت في /invoices"
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