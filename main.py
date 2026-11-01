import os
import json
import re
import gc
import unicodedata
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, CallbackQueryHandler, filters

TOKEN = os.getenv("BOT_TOKEN") or os.getenv("TELEGRAM_TOKEN")
ADMIN_ID = 1280808746
ALLOWED_FILE = "allowed_users.json"
PENDING_FILE = "pending_users.json"

print("Starting Bot - Ultra Light Mode...", flush=True)

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

def clean_desc(text):
    if not text: return ""
    t = str(text).strip()
    t = unicodedata.normalize('NFC', t)
    t = t.replace('ـ','')
    t = t.replace('إحعادات','إعدادات').replace('البعرودة','البرودة')
    t = t.replace('مكونعين','مكونين').replace('أكعثر','أكثر')
    t = t.replace('زجععاج','زجاج').replace('زجعاج','زجاج')
    t = re.sub(r'حع', 'ع', t)
    return ' '.join(t.split())

def has_arabic(s):
    return any('\u0600' <= c <= '\u06FF' for c in s)

def is_hs_code(s):
    t = s.strip().replace('.','').replace(' ','')
    return t.isdigit() and 2 <= len(t) <= 12

# === الحل السهل - بدون pandas ===
items = []
print("Loading EXLL.xlsx low memory...", flush=True)
try:
    from openpyxl import load_workbook
    wb = load_workbook("EXLL.xlsx", read_only=True, data_only=True)
    ws = wb.active
    for row in ws.iter_rows(values_only=True):
        if not row: continue
        code = None
        longest = ""
        # ندور على الكود من اليمين للشمال
        for val in reversed(row):
            s = str(val).strip() if val is not None else ""
            if is_hs_code(s) and s != '2007':
                code = s
                break
        for val in row:
            s = str(val).strip() if val is not None else ""
            if has_arabic(s) and len(s) > len(longest) and len(s) > 5 and len(s) < 400:
                if "جمهورية مصر" not in s and "وزارة المالية" not in s:
                    longest = s
        if code and longest:
            items.append({'code': code, 'desc': clean_desc(longest), 'clean': code.replace('.','').replace(' ','')})
    wb.close()
    gc.collect()
    print(f"Loaded {len(items)} - RAM: ~30MB", flush=True)
except Exception as e:
    print(f"Error: {e}", flush=True)

def is_allowed(uid):
    return uid in allowed_users or uid == ADMIN_ID

def search_combined(query):
    query = clean_desc(query.strip())
    q_lower = query.lower()
    digits = "".join([ch for ch in query if ch.isdigit()])
    if digits and len(digits) >= 2:
        for it in items:
            if it["clean"] == digits:
                return [it], []
        res = [it for it in items if it["clean"].startswith(digits) or digits.startswith(it["clean"])]
        if res: return res[:5], []
    words = [w for w in q_lower.split() if len(w) > 1]
    scored = []
    for it in items:
        dl = it["desc"].lower()
        score = 100 if q_lower in dl else 0
        for w in words:
            if w in dl: score += 10
        if score > 0: scored.append((score, it))
    scored.sort(key=lambda x: x[0], reverse=True)
    return [x[1] for x in scored[:7]], []

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
    await update.message.reply_text("اهلا بيك - البوت شغال وضع خفيف\nابحث: 5506 - 0101 - زجاج")

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
    msg = f"📦 البند: {ex_found[0]['code']}\n{ex_found[0]['desc']}\n\n"
    if len(ex_found) > 1:
        msg += "🔍 مشابهة:\n"
        for it in ex_found[1:3]:
            msg += f"- [{it['code']}] {it['desc'][:50]}\n"
    await update.message.reply_text(msg[:4000], parse_mode=None)

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
        try: await context.bot.send_message(chat_id=uid, text="تمت الموافقة! ابعت /start")
        except: pass
    else:
        uid = data.split("_")[1]
        pending_users.pop(uid, None)
        save_json(PENDING_FILE, pending_users)
        await q.edit_message_text("تم رفض " + uid)

if TOKEN:
    app = Application.builder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CallbackQueryHandler(button_callback))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_search))
    app.run_polling()