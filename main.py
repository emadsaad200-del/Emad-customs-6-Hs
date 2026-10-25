import os
import json
import re
import pandas as pd
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, CallbackQueryHandler, filters, ContextTypes

# ===== الإعدادات - تم تظبيطها ليك =====
TOKEN = os.getenv("BOT_TOKEN") or os.getenv("TELEGRAM_TOKEN")
ADMIN_ID = 1280808746  # الأدمن - Emad

ALLOWED_FILE = "allowed_users.json"
PENDING_FILE = "pending_users.json"

print("Starting Bot...")
print("Loading EXLL.xlsx ...")

try:
    df = pd.read_excel("EXLL.xlsx")
    print(f"File loaded: {len(df)} rows")
except Exception as e:
    print(f"Error loading: {e}")
    df = pd.DataFrame()

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

items = []
for _, row in df.iterrows():
    try:
        code = str(row.iloc[0]).strip()
        desc = str(row.iloc[1]).strip() if len(row) > 1 else ""
        if not code or code == 'nan': continue
        clean = re.sub(r'[^0-9]', '', code)
        items.append({'code': code, 'desc': desc, 'clean': clean, 'search': f"{desc} {code}".lower()})
    except: continue

print(f"Loaded {len(items)} items - FINAL FIXED PRIVATE")
print(f"Token exists: {TOKEN is not None}")
print(f"Admin ID: {ADMIN_ID}")

def is_allowed(user_id):
    return user_id in allowed_users or user_id == ADMIN_ID

async def request_approval(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    user_id = user.id
    if is_allowed(user_id):
        return True
    if str(user_id) in pending_users:
        await update.message.reply_text("⏳ طلبك قيد المراجعة، انتظر موافقة الأدمن.")
        return False
    pending_users[str(user_id)] = {"name": user.full_name, "username": user.username}
    save_json(PENDING_FILE, pending_users)
    await update.message.reply_text("🔒 البوت خاص.\nتم إرسال طلبك للأدمن، سيتم تفعيلك قريباً.")
    keyboard = [[InlineKeyboardButton("✅ موافقة", callback_data=f"approve_{user_id}"), InlineKeyboardButton("❌ رفض", callback_data=f"reject_{user_id}")]]
    try:
        await context.bot.send_message(chat_id=ADMIN_ID, text=f"🔔 طلب استخدام جديد:\n👤 {user.full_name}\n🆔 {user_id}\n🔗 @{user.username}", reply_markup=InlineKeyboardMarkup(keyboard))
    except: pass
    return False

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await request_approval(update, context): return
    await update.message.reply_text("أهلاً بيك في بوت التعريفة الجمركية ✅\nابعت كود البند مثل 73-02-10-00-00")

async def handle_code(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await request_approval(update, context): return
    text = update.message.text.strip()
    clean_input = re.sub(r'[^0-9]', '', text)
    if not clean_input:
        await update.message.reply_text("ابعت كود صحيح")
        return
    for item in items:
        if item['clean'] == clean_input or item['code'] == text:
            await update.message.reply_text(f"📋 البند: {item['code']}\n📝 {item['desc']}")
            return
    for item in items:
        if clean_input in item['clean']:
            await update.message.reply_text(f"📋 البند: {item['code']}\n📝 {item['desc']}")
            return
    await update.message.reply_text("❌ البند غير موجود")

async def button_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    if query.from_user.id != ADMIN_ID:
        await query.edit_message_text("❌ أنت لست الأدمن")
        return
    if data.startswith("approve_"):
        uid = int(data.split("_")[1])
        allowed_users.add(uid)
        save_json(ALLOWED_FILE, list(allowed_users))
        pending_users.pop(str(uid), None)
        save_json(PENDING_FILE, pending_users)
        await query.edit_message_text(f"✅ تمت الموافقة على {uid}")
        try:
            await context.bot.send_message(chat_id=uid, text="✅ تمت الموافقة! يمكنك الآن استخدام البوت.\nابعت /start")
        except: pass
    elif data.startswith("reject_"):
        uid = data.split("_")[1]
        pending_users.pop(uid, None)
        save_json(PENDING_FILE, pending_users)
        await query.edit_message_text(f"❌ تم رفض {uid}")

async def admin_list(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID: return
    await update.message.reply_text(f"✅ المصرح لهم: {list(allowed_users)}\n⏳ المنتظرين: {pending_users}")

if TOKEN:
   
