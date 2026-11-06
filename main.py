import os
import json
import re
import unicodedata
import sqlite3
# تجاوز ملف items.db المقفول - استخدم ملف جديد
_orig_connect = sqlite3.connect
def _patched_connect(path, *a, **k):
    if path == "items.db":
        path = "items_v2.db"
    return _orig_connect(path, *a, **k)
sqlite3.connect = _patched_connect
import openpyxl

# =========================================================
# 1- بناء قاعدة البيانات: عربي + انجليزي من ملف الموسوعة
# =========================================================
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

    ar = {}  # code -> (desc_ar, duty)
    en = {}  # code -> desc_en

    for row in ws.iter_rows(values_only=True):
        if not row or not row[0]:
            continue
        t = str(row[0]).strip()
        if len(t) < 7:
            continue
        # الكود بيكون زي 2309.10 او 5516.11
        code = ""
        for part in t.replace("-", " ").split():
            p = part.strip()
            if len(p) >= 6 and p[0].isdigit() and "." in p:
                # خد اول حاجة شكلها كود
                if re.match(r"^\d{4}\.\d+", p):
                    code = p
                    break
        if not code:
            continue
        # الفئة: معفاه او رقم في اخر السطر
        duty = None
        if "معفاه" in t or "معفاة" in t or "معفي" in t or "معفى" in t:
            duty = 0
        else:
            w = t.split()
            if w and w[-1].replace(".", "").isdigit():
                try:
                    duty = float(w[-1])
                except:
                    duty = None
        desc = t.replace(code, "").replace("-", " ").strip()
        # شيل رقم الفئة من اخر الوصف لو موجود
        desc = re.sub(r"\s+\d+(\.\d+)?\s*$", "", desc).strip()
        if len(desc) < 3:
            desc = code
        is_ar = any('\u0600' <= ch <= '\u06FF' for ch in desc)
        if is_ar:
            ar[code] = (desc, duty)
        else:
            # الانجليزي - تجاهل العناوين الطويلة جدا
            if len(desc) > 3:
                en[code] = desc

    all_codes = set(ar.keys()) | set(en.keys())
    for code in all_codes:
        desc_ar, duty = ar.get(code, ("", None))
        desc_en = en.get(code, "")
        # الوصف الموحد للبحث القديم
        desc_combined = (desc_ar + " | " + desc_en).strip(" |")
        clean = code.replace(".", "").replace(" ", "")
        cur.execute("INSERT INTO items VALUES (?,?,?,?,?,?)",
                    (code, desc_ar, desc_en, desc_combined + (f" | فئة: {duty}" if duty is not None else ""), clean, duty if duty is not None else 0))
    conn.commit()
    conn.close()
    print(f"Built items.db: {len(all_codes)} items | AR: {len(ar)} | EN: {len(en)}", flush=True)

import requests
from bs4 import BeautifulSoup
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, CallbackQueryHandler, filters

TOKEN = os.getenv("BOT_TOKEN") or os.getenv("TELEGRAM_TOKEN")
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

def clean_desc(text):
    if not text:
        return ""
    t = str(text).strip()
    t = unicodedata.normalize('NFKC', t)
    t = unicodedata.normalize('NFC', t)
    t = t.replace('ـ', '')
    t = ' '.join(t.split())
    return t

# =========================================================
# 2- قواعد الرسوم الاضافية - من الملف والقانون
# الفئة نفسها بتيجي من الموسوعة، الباقي قواعد عامة
# =========================================================
VAT_DEFAULT = 14.0
VAT_EXEMPT_PREFIX = ["2309.9010"]  # اضافات فيتامينات - معفاه في ملفك، ضيف عليها لو عندك

# رسم التنمية: الافتراضي 10% - فصول غذائية اساسية معفاة
DEV_DEFAULT = 0.0
PROFIT_DEFAULT = 1.0
DEV_EXEMPT_CHAPTERS = ["01", "02", "03", "04", "07", "08", "10", "11", "12"]

# ضريبة جدول - لبعض السلع فقط، ضيف هنا لو عندك بند
TABLE_TAX = {
    # مثال: "2402.20": 50,
}

# رسم اغراق / وقائي - جدول منفصل، لما يجيلك منشور ضيف الكود هنا بس
ANTI_DUMPING = {
    # مثال: "7210.70": 15,
}

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
    # 1- Tariff.json لو موجود
    for t in tariff_items:
        tc = str(t.get("code", "")).replace(".", "").replace(" ", "")
        if tc == code_clean or code_clean.startswith(tc) or tc.startswith(code_clean):
            return t
    # 2- من الموسوعة نفسها
    try:
        conn = sqlite3.connect("items.db")
        cur = conn.cursor()
        cur.execute("SELECT code, desc_ar, desc_en, duty FROM items WHERE REPLACE(code,'.','')=?", (code_clean,))
        row = cur.fetchone()
        if not row:
            cur.execute("SELECT code, desc_ar, desc_en, duty FROM items WHERE REPLACE(code,'.','') LIKE ?", (code_clean[:6]+"%",))
            row = cur.fetchone()
        conn.close()
        if row:
            code_full = row[0]
            duty = float(row[3] or 0)
            return {
                "code": code_full,
                "duty": duty,
                "vat": get_vat(code_full),
                "dev": get_dev_fee(code_full),
                "table": get_table_tax(code_full),
                "dumping": get_dumping(code_full),
                "profit": 0,
                "desc_ar": row[1] or "",
                "desc_en": row[2] or "",
            }
    except Exception:
        pass
    return None

def calc_customs(cif, duty_p, vat_p, profit_p, dev_p=0, table_p=0, dumping_p=0):
    # الترتيب الصح: جمرك -> تنمية -> جدول -> اغراق -> قيمة مضافة على الكل
    duty = cif * duty_p / 100
    dev = cif * dev_p / 100
    table = (cif + duty) * table_p / 100 if table_p else 0
    dumping = cif * dumping_p / 100 if dumping_p else 0
    vat = (cif + duty + dev + table + dumping) * vat_p / 100
profit = (cif + duty) * profit_p / 100
    total = duty + vat + profit + dev + table + dumping
    return duty, vat, profit, total, dev, table, dumping

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
        # بحث عربي او انجليزي
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
    await update.message.reply_text("اهلا بيك في بوت Logistics Group\n\nجرب تبعت بس من غير اوامر:\n\n8471 -> هيجيب البند\nزجاج -> هيجيب البند بالاسم\n8471 100000 -> هيحسبلك الجمارك\nمنشور 28 -> هيجيب المنشور\n\nبس كده!")

async def handle_smart(update, context):
    ok = await request_approval(update, context)
    if not ok:
        return
    text = update.message.text.strip()
    low = text.lower()
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
        profit_p = float(tariff.get("profit", 0))
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
            # زي التطبيق: كود + عربي + انجليزي + رسم
            msg += f"- [{it2['code']}] {it2.get('desc_ar','')[:40]} | {it2.get('desc_en','')[:40]} | فئة: {it2.get('duty',0)}%\n"
    await update.message.reply_text(msg[:4000])

async def button_callback(update, context):
    q = update.callback_query
    await q.answer()
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
    app.add_handler(CommandHandler("users", admin_list))
    app.add_handler(CommandHandler("remove", remove_user))
    app.add_handler(CallbackQueryHandler(button_callback))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_smart))
    app.run_polling()