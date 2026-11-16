import json, os
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ApplicationBuilder, CommandHandler, CallbackQueryHandler, ContextTypes

TOKEN = "8631742684:AAE7rxr4sijmXu2IgwrrdkBffuQS3Kv9zkM"
ADMIN_ID = 1280808746
DB_FILE = "users.json"

def load_users():
    if not os.path.exists(DB_FILE):
        return {}
    with open(DB_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def save_users(data):
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def is_admin(user_id):
    return user_id == ADMIN_ID

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    users = load_users()
    uid = str(user.id)
    if uid in users and users[uid]["status"] == "approved":
        await update.message.reply_text(f"أهلا بيك يا {user.first_name} ✅\nالبوت جاهز")
        return
    users[uid] = {"name": user.first_name, "username": user.username or "مفيش يوزر", "code": user.id, "status": "pending"}
    save_users(users)
    await update.message.reply_text("في انتظار موافقة الأدمن Captain Emad ⏳\nأول ما يقبلك هتقدر تستخدم البوت")
    keyboard = [[InlineKeyboardButton("✅ قبول", callback_data=f"approve_{uid}"), InlineKeyboardButton("❌ رفض", callback_data=f"reject_{uid}")]]
    await context.bot.send_message(chat_id=ADMIN_ID, text=f"🔔 طلب دخول جديد\n\nالاسم: {user.first_name}\nالكود: {user.id}\nاليوزر: @{user.username or 'مفيش'}\nالحالة: في الانتظار", reply_markup=InlineKeyboardMarkup(keyboard))

async def handle_admin(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    action, uid = query.data.split("_")
    users = load_users()
    if uid not in users:
        await query.edit_message_text("المستخدم ده مش موجود")
        return
    if action == "approve":
        users[uid]["status"] = "approved"
        save_users(users)
        await query.edit_message_text(f"تم قبول {users[uid]['name']} ✅\nالكود: {uid}")
        await context.bot.send_message(chat_id=int(uid), text="تم قبولك ✅\nتقدر تستخدم البوت دلوقتي")
    else:
        users[uid]["status"] = "rejected"
        save_users(users)
        await query.edit_message_text(f"تم رفض {users[uid]['name']} ❌\nالكود: {uid}")
        await context.bot.send_message(chat_id=int(uid), text="تم رفض طلبك ❌\nتقدر تبعت /start تاني لو حابب تطلب تاني")

async def users_count(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id):
        return
    users = load_users()
    approved = sum(1 for u in users.values() if u["status"] == "approved")
    pending = sum(1 for u in users.values() if u["status"] == "pending")
    rejected = sum(1 for u in users.values() if u["status"] == "rejected")
    await update.message.reply_text(f"📊 عدد المستخدمين\n\nالكل: {len(users)}\n✅ مقبول: {approved}\n⏳ في الانتظار: {pending}\n❌ مرفوض: {rejected}")

async def list_users(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id):
        return
    users = load_users()
    if not users:
        await update.message.reply_text("مفيش مستخدمين لسه")
        return
    txt = "📋 قائمة المستخدمين:\n\n"
    for uid, u in users.items():
        emoji = "✅" if u["status"] == "approved" else "⏳" if u["status"] == "pending" else "❌"
        txt += f"{emoji} {u['name']} | @{u['username']} | كود: {uid} | {u['status']}\n"
    await update.message.reply_text(txt)

async def delete_user(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id):
        return
    if not context.args:
        await update.message.reply_text("اكتب كده:\n/delete 123456789")
        return
    uid = context.args[0]
    users = load_users()
    if uid not in users:
        await update.message.reply_text("الكود ده مش موجود")
        return
    name = users[uid]["name"]
    del users[uid]
    save_users(users)
    await update.message.reply_text(f"تم حذف {name} ❌\nالكود: {uid}")
    try:
        await context.bot.send_message(chat_id=int(uid), text="تم حذف حسابك من البوت ❌\nلو حابب ترجع، ابعت /start تاني")
    except:
        pass

app = ApplicationBuilder().token(TOKEN).build()
app.add_handler(CommandHandler("start", start))
app.add_handler(CommandHandler("users", users_count))
app.add_handler(CommandHandler("list", list_users))
app.add_handler(CommandHandler("delete", delete_user))
app.add_handler(CallbackQueryHandler(handle_admin))
app.run_polling()