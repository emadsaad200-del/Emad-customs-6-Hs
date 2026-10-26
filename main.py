import os
import json
import pandas as pd
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, CallbackQueryHandler, filters

TOKEN = os.getenv("BOT_TOKEN") or os.getenv("TELEGRAM_TOKEN")
ADMIN_ID = 1280808746

ALLOWED_FILE = "allowed_users.json"
PENDING_FILE = "pending_users.json"

print("Starting Bot...", flush=True)

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

def has_arabic(s):
    for c in s:
        if '\u0600' <= c <= '\u06FF':
            return True
    return False

def is_hs_code(s):
    t = s.strip()
    if not t:
        return False
    cleaned = t.replace('.','').replace(' ','')
    if not cleaned.isdigit():
        return False
    if len(cleaned) < 2 or len(cleaned) > 8:
        return False
    return True

def clean_desc(text):
    if not text:
        return ""
    t = text.strip()
    t = t.replace('رير مشربة', 'غير مشربة')
    t = t.replace('ريرها', 'غيرها')
    t = t.replace('زجععاج', 'زجاج')
    t = t.replace('مجلعو', 'مجلو')
    t = t.replace('معصقول', 'مصقول')
    t = t.replace('ـ', '')
    return ' '.join(t.split())

def find_code_and_desc(row):
    code = None
    for val in reversed(list(row)):
        s = str(val).strip()
        if s == 'nan' or s == '':
            continue
        if is_hs_code(s):
            if len(s.replace('.','')) == 1:
                continue
            code = s
            break
    longest = ""
    for val in row:
        s = str(val).strip()
        if s == 'nan' or s == '':
            continue
        if has_arabic(s) and len(s) > len(longest) and len(s) > 5:
            longest = s
    return code, longest

print("Loading EXLL.xlsx ...", flush=True)
df = pd.read_excel("EXLL.xlsx", header=None)
items = []
for idx in range(len(df)):
    row = df.iloc[idx]
    c,d = find_code_and_desc(row)
    if c is None or d == "":
        continue
    if len(d) > 400:
        continue
    if "جمهورية مصر" in d or "وزارة المالية" in d:
        continue
    if c == '2007':
        continue
    cd = clean_desc(d)
    items.append({'code': c, 'desc': cd, 'clean': c.replace('.','')})

print(f"Loaded {len(items)} items", flush=True)

def is_allowed(uid):
    return uid in allowed_users or uid == ADMIN_ID

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
    except:
        pass
    return False

async def start(update, context):
    ok = await request_approval(update, context)
    if not ok:
        return
    await update.message.reply_text("اهلا بيك - ابعت كود البند مثل 4405 او 5506")

async def handle_code(update, context):
    ok = await request_approval(update, context)
    if not ok:
        return
    text = update.message.text.strip()
    digits = "".join([ch for ch in text if ch.isdigit()])
    if not digits:
        await update.message.reply_text("ابعت كود صحيح")
        return
    for item in items:
        if item["clean"] == digits:
            await update.message.reply_text("البند: " + item['code'] + "\n" + item['desc'])
            return
    for item in items:
        if item["clean"].startswith(digits) or digits.startswith(item["clean"]):
            await update.message.reply_text("البند: " + item['code'] + "\n" + item['desc'])
            return
    await update.message.reply_text("البند غير موجود")

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
        except:
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
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_code))
    app.run_polling()
