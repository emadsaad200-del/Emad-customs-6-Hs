import os
import json
import re
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
print("Starting Bot - SQLite Light...", flush=True)
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
    full_map = {'إعدادات البعرودة': 'إعدادات البرودة','إحعادات': 'إعدادات','البعرودة': 'البرودة','مععدات': 'معدات','مكونعين': 'مكونين','أكعثر': 'أكثر','أرلفة': 'أغلفة','الحمعل': 'الحمل','أدويعة': 'أدوية','زجععاج': 'زجاج','زجعاج': 'زجاج','بعرد': 'برد'}
    for bad in sorted(full_map, key=len, reverse=True):
        if bad in t:
            t = t.replace(bad, full_map[bad])
    t = t.replace('حع', 'ع')
    t = re.sub(r'عع+', 'ع', t)
    return ' '.join(t.split())
def is_allowed(uid):
    return uid in allowed_users or uid == ADMIN_ID
tariff_items = []
if os.path.exists("Tariff.json"):
    tariff_items = load_json("Tariff.json", [])
def get_tariff(code_clean):
    for t in tariff_items:
        tc = str(t.get("code", "")).replace(".", "").replace(" ", "")
        if tc == code_clean or code_clean.startswith(tc) or tc.startswith(code_clean):
            return t
    return None
def calc_customs(cif, duty_p, vat_p, profit_p):
    duty = cif * duty_p / 100
    vat = (cif + duty) * vat_p / 100
    profit = cif * profit_p / 100
    total = duty + vat + profit
    return duty, vat, profit, total
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
        if digits and len(digits) >= 2:
            cur.execute("SELECT code, desc FROM items WHERE clean = ? LIMIT 1", (digits,))
            row = cur.fetchone()
            if row:
                conn.close()
                return [{'code': row[0], 'desc': clean_desc(row[1]), 'clean': digits}], []
            cur.execute("SELECT code, desc FROM items WHERE clean LIKE ? LIMIT 5", (digits + "%",))
            rows = cur.fetchall()
            conn.close()
            for r in rows:
                result.append({'code': r[0], 'desc': clean_desc(r[1]), 'clean': r[0].replace(".", "").replace(" ", "")})
            return result, []
        words = [w for w in q_lower.split() if len(w) > 1]
        if not words:
            conn.close()
            return [], []
        cur.execute("SELECT code, desc FROM items WHERE desc LIKE ? LIMIT 7", ("%" + words[0] + "%",))
        rows = cur.fetchall()
        conn.close()
        for r in rows:
            result.append({'code': r[0], 'desc': clean_desc(r[1]), 'clean': r[0].replace(".", "").replace(" ", "")})
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
    if is_allowed(user.id):
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
    await update.message.reply_text("اهلا بيك في بوت Logistics Group\nابحث ب:\n- الكود: 5506 - 0101\n- الاسم: زجاج - خشب - خيول\n\nاوامر جديدة:\n/calc 8471 100000 -> حاسبة تقديرية\n/qararat -> آخر قرارات الجمارك")
async def handle_search(update, context):
    ok = await request_approval(update, context)
    if not ok:
        return
    text = update.message.text.strip()
    if len(text) < 2:
        await update.message.reply_text("اكتب كلمة اطول")
        return
    ex_found, _ = search_combined(text)
    if not ex_found:
        await update.message.reply_text("مفيش نتائج لـ: " + text)
        return
    it = ex_found[0]
    msg = "📦 البند: " + it['code'] + "\n" + it['desc'] + "\n\n"
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
    await update.message.reply_text(msg)
async def calc_handler(update, context):
    ok = await request_approval(update, context)
    if not ok:
        return
    try:
        args = context.args
        if len(args) < 2:
            await update.message.reply_text("طريقة الحساب:\n/calc [البند] [قيمة الفاتورة]\nمثال:\n/calc 8471 100000")
            return
        code = args[0].replace(".", "").replace(" ", "")
        cif = float(args[1])
        tariff = get_tariff(code)
        if not tariff:
            await update.message.reply_text("البند مش موجود في Tariff.json")
            return
        duty_p = float(tariff.get("duty", 0))
        vat_p = float(tariff.get("vat", 14))
        profit_p = float(tariff.get("profit", 0))
        duty, vat, profit, total = calc_customs(cif, duty_p, vat_p, profit_p)
        msg = f"🧾 حساب تقديري - البند: {code}\n"
        msg += f"قيمة CIF: {cif:,.0f} ج\n\n"
        msg += f"نسبة الوارد: {duty_p}% = {duty:,.0f} ج\n"
        msg += f"القيمة المضافة: {vat_p}% = {vat:,.0f} ج\n"
        msg += f"أرباح تجارية: {profit_p}% = {profit:,.0f} ج\n"
        msg += "----------------\n"
        msg += f"إجمالي تقديري: {total:,.0f} ج\n\n"
        msg += f"ملاحظة: {tariff.get('notes','')}\n"
        msg += "⚠️ حساب استرشادي فقط"
        await update.message.reply_text(msg)
    except Exception:
        await update.message.reply_text("خطأ في الحساب\nمثال: /calc 8471 100000")
async def decisions_handler(update, context):
    ok = await request_approval(update, context)
    if not ok:
        return
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
            msg += "\n\nالمصدر: customs.gov.eg"
        await update.message.reply_text(msg)
    except Exception:
        await update.message.reply_text("موقع الجمارك مش متاح دلوقتي\nhttps://www.customs.gov.eg/")
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
if TOKEN:
    print("Bot polling started...", flush=True)
    app = Application.builder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("calc", calc_handler))
    app.add_handler(CommandHandler("qararat", decisions_handler))
    app.add_handler(CommandHandler("decisions", decisions_handler))
    app.add_handler(CallbackQueryHandler(button_callback))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_search))
    app.run_polling()