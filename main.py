import os
import json
import re
import unicodedata
import sqlite3
# تحويل موسوعة 2025
import openpyxl
fname = None
for n in ["Book1LOGIST TATTEF.xlsx", "Book1LOGIST_TATTEF.xlsx"]:
    if os.path.exists(n):
        fname = n
        break
if fname and not os.path.exists("items.db"):
    wb = openpyxl.load_workbook(fname, read_only=True, data_only=True)
    ws = wb.active
    conn = sqlite3.connect("items.db")
    cur = conn.cursor()
    cur.execute("CREATE TABLE IF NOT EXISTS items (code TEXT, desc TEXT, clean TEXT)")
    cur.execute("DELETE FROM items")
    for row in ws.iter_rows(values_only=True):
        if not row or not row[0]:
            continue
        t = str(row[0]).strip()
        if len(t) < 7:
            continue
        code = ""
        for part in t.replace("-", " ").split():
            p = part.strip()
            if len(p) == 7 and p[4] == "." and p[:4].isdigit() and p[5:].isdigit():
                code = p
                break
        if code:
            duty = "معفاه" if ("معفاه" in t or "معفاة" in t) else ""
            if not duty:
                w = t.split()
                if w and w[-1].isdigit():
                    duty = w[-1]
            desc = t.replace(code, "").replace(duty, "").replace("-", " ").strip()
            if len(desc) < 3:
                desc = code
            cur.execute("INSERT INTO items VALUES (?,?,?)", (code, desc + " | فئة: " + duty, code.replace(".", "")))
    conn.commit()
    conn.close()
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
        if digits and len(digits) >= 2 and len(query.replace(" ", "")) <= 8:
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
            await update.message.reply_text(f"البند {code} مش موجود في Tariff.json\nجرب بند تاني او ابعت للادمن النسب")
            return
        duty_p = float(tariff.get("duty", 0))
        vat_p = float(tariff.get("vat", 14))
        profit_p = float(tariff.get("profit", 0))
        duty, vat, profit, total = calc_customs(cif, duty_p, vat_p, profit_p)
        msg = f"🧾 حساب - البند: {code}\n"
        msg += f"قيمة CIF: {cif:,.0f} ج\n\n"
        msg += f"وارد {duty_p}% = {duty:,.0f} ج\n"
        msg += f"قيمة مضافة {vat_p}% = {vat:,.0f} ج\n"
        msg += f"ارباح تجارية {profit_p}% = {profit:,.0f} ج\n"
        msg += "----------------\n"
        msg += f"اجمالي تقديري: {total:,.0f} ج\n"
        msg += f"\nملاحظة: {tariff.get('notes','')}"
        await update.message.reply_text(msg)
        return
    ex_found, _ = search_combined(text)
    if not ex_found:
        await update.message.reply_text("مفيش نتائج لـ: " + text + "\n\nتأكد ان ملف items.db موجود في FadeHost\nاو جرب: منشور 28 - او - 8471 100000")
        return
    it = ex_found[0]
    msg = "📦 البند: " + it['code'] + "\n" + it['desc'] + "\n\n"
    tariff = get_tariff(it['clean'])
    if tariff:
        msg += f"💰 وارد: {tariff.get('duty',0)}% | قيمة مضافة: {tariff.get('vat',14)}% | ارباح: {tariff.get('profit',0)}%\n"
        msg += f"📋 {tariff.get('notes','')}\n\n"
    else:
        msg += "💰 النسب: مش موجودة في Tariff.json\n\n"
    msg += f"🔗 التأكد من نافذة:\nhttps://www.nafeza.gov.eg/\n\n"
    msg += "💡 عشان تحسبها ابعت:\n" + it['clean'] + " 100000"
    if len(ex_found) > 1:
        msg += "\n\n🔍 مشابهة:\n"
        for it2 in ex_found[1:3]:
            msg += "- [" + it2['code'] + "] " + it2['desc'][:50] + "\n"
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
if TOKEN:
    print("Bot polling started...", flush=True)
    app = Application.builder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CallbackQueryHandler(button_callback))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_smart))
    app.run_polling()
