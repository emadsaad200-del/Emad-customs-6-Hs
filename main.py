import os
import json
import re
import gc
import unicodedata
import sqlite3
import requests
from bs4 import BeautifulSoup
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, CallbackQueryHandler, filters

TOKEN = os.getenv("BOT_TOKEN") or os.getenv("TELEGRAM_TOKEN")
ADMIN_ID = 1280808746
ALLOWED_FILE = "allowed_users.json"
PENDING_FILE = "pending_users.json"

print("Starting Bot - SQLite Light Version...", flush=True)

def load_json(file, default):
    if os.path.exists(file):
        try:
            with open(file, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return default
    return default

def save_json(file, data):
    with open(file, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)

allowed_users = set(load_json(ALLOWED_FILE, [ADMIN_ID]))
pending_users = load_json(PENDING_FILE, {})

# ========= إصلاح حرف العين =========
def clean_desc(text):
    if not text:
        return ""
    t = str(text).strip()
    try:
        if 'حع' in t or 'إحعادات' in t or 'البعرودة' in t:
            try:
                b = t.encode('latin1', errors='ignore')
                for enc in ['windows-1256', 'cp1256']:
                    try:
                        dec = b.decode(enc)
                        if 'إعدادات' in dec or 'البرودة' in dec:
                            t = dec
                            break
                    except:
                        continue
            except:
                pass
    except:
        pass
    t = unicodedata.normalize('NFKC', t)
    t = unicodedata.normalize('NFC', t)
    t = t.replace('ـ', '')
    full_map = {
        'إعدادات البعرودة': 'إعدادات البرودة',
        'إحعادات البعرودة': 'إعدادات البرودة',
        'إحعادات': 'إعدادات',
        'احعادات': 'إعدادات',
        'إحعدادات': 'إعدادات',
        'البعرودة': 'البرودة',
        'البعروده': 'البرودة',
        'مععدات': 'معدات',
        'مكونعين': 'مكونين',
        'مكونع': 'مكون',
        'أكعثر': 'أكثر',
        'كعثر': 'كثر',
        'أرلفة': 'أغلفة',
        'الحمعل': 'الحمل',
        'أدويعة': 'أدوية',
        'زجععاج': 'زجاج',
        'زجعاج': 'زجاج',
        'الاليات': 'الآلات',
        'بعرد': 'برد',
    }
    for bad in sorted(full_map, key=len, reverse=True):
        if bad in t:
            t = t.replace(bad, full_map[bad])
    t = t.replace('حع', 'ع')
    t = re.sub(r'(\w)حع(\w)', r'\1ع\2', t)
    t = re.sub(r'عع+', 'ع', t)
    t = t.replace('ع ع ع','').replace('ع ع','')
    t = re.sub(r'\bرير\b', 'غير', t)
    t = re.sub(r'\bريرها\b', 'غيرها', t)
    t = re.sub(r'\bرير مشربة\b', 'غير مشربة', t)
    t = re.sub(r'(\w)عين\b', r'\1ين', t)
    t = re.sub(r'ويعة\b', 'وية', t)
    return ' '.join(t.split())

def is_allowed(uid):
    return uid in allowed_users or uid == ADMIN_ID

# ========= التعريفة من Tariff.json =========
tariff_items = []
if os.path.exists("Tariff.json"):
    tariff_items = load_json("Tariff.json", [])

def get_tariff(code_clean):
    for t in tariff_items:
        tc = str(t.get("code","")).replace(".","").replace(" ","")
        if tc == code_clean or code_clean.startswith(tc) or tc.startswith(code_clean):
            return t
    return None

def calc_customs(cif, duty_p, vat_p, profit_p):
    duty = cif * duty_p / 100
    vat_base = cif + duty
    vat = vat_base * vat_p / 100
    profit = cif * profit_p / 100
    total = duty + vat + profit
    return duty, vat, profit, total

# ========= البحث من SQLite - خفيف جدا بدون EXLL.xlsx =========
def search_combined(query):
    query = clean_desc(query.strip())
    q_lower = query.lower()
    digits = "".join([ch for ch in query if ch.isdigit()])

    if not os.path.exists("items.db"):
        return [], []

    conn = sqlite3.connect("items.db")
    cur = conn.cursor()

    try:
        if digits and len(digits) >= 2:
            cur.execute("SELECT code, desc FROM items WHERE clean = ? LIMIT 1", (digits,))
            row = cur.fetchone()
            if row:
                conn.close()
                return [{'code': row[0], 'desc': clean_desc(row[1]), 'clean': digits}], []
            cur.execute("SELECT code, desc FROM items WHERE clean LIKE ? LIMIT 5", (digits + "%",))
            rows = cur.fetchall()
            conn.close()
            return [{'code': r[0], 'desc': clean_desc(r[1]), 'clean': r[0].replace(".","").replace(" ","")} for r in rows], []

        words = [w for w in q_lower.split() if len(w) > 1]
        if not words:
            conn.close()
            return [], []
        cur.execute("SELECT code, desc FROM items WHERE desc LIKE ? LIMIT 7", ("%" + words[0] + "%",))
        rows = cur.fetchall()
        conn.close()
        return [{'code': r[0], 'desc': clean_desc(r[1]), 'clean': r[0].replace(".","").replace(" ","")} for r in rows], []
    except Exception as e:
        print(f"Search error: {e}", flush=True)
        try: conn.close()
        except: pass
        return [], []

async def request_approval(update, context):
    user = update.effective_user
    if is_allowed(user.id): return True
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
    except: pass
    return False

async def start(update, context):
    ok = await request_approval(update, context)
    if not ok: return
    await update.message.reply_text("اهلا بيك في بوت Logistics Group\nابحث ب:\n- الكود: 5506 - 0101\n- الاسم: زجاج - خشب - خيول\n\nاوامر جديدة:\n/calc 8471 100000 -> حاسبة تقديرية\n/qararat -> آخر قرارات الجمارك")

async def handle_search(update, context):
    ok = await request_approval(update, context)
    if not ok: return
    text = update.message.text.strip()
    if len(text) < 2:
        await update.message.reply_text("اكتب كلمة اطول")
        return
    ex_found, _ = search_combined(text)
    if not ex_found:
        await update.message.reply_text("مفيش نتائج لـ: " + text)
        return
    msg = ""
    it = ex_found[0]
    msg += "📦 البند: " + it['code'] + "\n" + it['desc'] + "\n\n"
    tariff = get_tariff(it['clean'])
    if tariff:
        msg += f"💰 وارد: {tariff.get('duty',0)}% | قيمة مضافة: {tariff.get('vat',14)}% | أرباح تجارية: {tariff.get('profit',0)}%\n"
        msg += f"📋 التعريفة: {tariff.get('notes','')}\n\n"
    else:
        msg += "💰 النسب: غير موجودة في Tariff.json\n\n"
    if len(ex_found) > 1:
        msg += "🔍 نتائج مشابهة:\n"
        for it2 in ex_found[1:4]:
            msg += "- [" + it2['code'] + "] " + it2['desc'][:50] + "\n"
    if len(msg) > 4000:
        msg = msg[:4000]
    await update.message.reply_text(msg, parse_mode=None)

async def calc_handler(update, context):
    ok = await request_approval(update, context)
    if not ok: return
    try:
        args = context.args
        if len(args) < 2:
            await update.message.reply_text("طريقة الحساب:\n/calc [البند] [قيمة الفاتورة بالجنيه]\n\nمثال:\n/calc 8471 100000")
            return
        code = args[0].replace(".","").replace(" ","")
        cif = float(args[1])
        tariff = get_tariff(code)
        if not tariff:
            await update.message.reply_text("البند مش موجود في ملف التعريفة Tariff.json")
            return
        duty_p = float(tariff.get("duty",0))
        vat_p = float(tariff.get("vat",14))
        profit_p = float(tariff.get("profit",0))
        duty, vat, profit, total = calc_customs(cif, duty_p, vat_p, profit_p)
        msg = f"🧾 حساب تقديري - البند: {code}\n"
        msg += f"قيمة CIF: {cif:,.0f} ج\n\n"
        msg += f"نسبة الوارد: {duty_p}% = {duty:,.0f} ج\n"
        msg += f"القيمة المضافة: {vat_p}% = {vat:,.0f} ج\n"
        msg += f"أرباح تجارية: {profit_p}% = {profit:,.0f} ج\n"
        msg += f"----------------\n"
        msg += f"إجمالي تقديري: {total:,.0f} ج\n\n"
        msg += f"ملاحظة: {tariff.get('notes','')}\n"
        msg += "⚠️ حساب استرشادي فقط - بدون رسم تنمية ومصاريف تخليص"
        await update.message.reply_text(msg)
    except Exception as e:
        await update.message.reply_text("خطأ في الحساب، اتأكد من الأرقام\nمثال: /calc 8471 100000")

async def decisions_handler(update, context):
    ok = await request_approval(update, context)
    if not ok: return
    await update.message.reply_text("جاري جلب آخر قرارات الجمارك من customs.gov.eg ...")
    try:
        url = "https://www.customs.gov.eg/"
        headers = {"User-Agent": "Mozilla/5.0"}
        r = requests.get(url, headers=headers, timeout=15)
        soup = BeautifulSoup(r.text, "html.parser")
        links = []
        for a in soup.find_all("a", href=True):
            txt = a.get_text(strip=True)
            if len(txt) > 20 and ("قرار" in txt or "تعريفة" in txt or "منشور" in txt or "جمارك" in txt):
                links.append(f"• {txt[:100]}")
            if len(links) >= 7:
                break
        if not links:
            msg = "موقع الجمارك مش راجع بيانات دلوقتي.\nادخل مباشرة: https://www.customs.gov.eg/"
        else:
            msg = "📢 آخر قرارات مصلحة الجمارك:\n\n" + "\n\n".join(links)
            msg