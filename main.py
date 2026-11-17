import os, json, time, random, string, re
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, CallbackQueryHandler, ContextTypes, filters

BOT_TOKEN = os.getenv("BOT_TOKEN")
ADMIN_ID = os.getenv("ADMIN_ID")
USERS_FILE = "users.json"
REJECT_WAIT = 5 * 60  # 5 minutes

# --- الجديد: ملفات المستندات ---
DOCS_FILE = "docs_index.json"
os.makedirs("docs", exist_ok=True)

def load_docs():
    if not os.path.exists(DOCS_FILE):
        return []
    try:
        with open(DOCS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return []

def save_docs(docs):
    with open(DOCS_FILE, "w", encoding="utf-8") as f:
        json.dump(docs, f, ensure_ascii=False, indent=2)

WELCOME_MESSAGE = """Welcome to Captain Emad Customs Bot.

After the trial and subscription for 150 EGP, you can inquire about the customs item by code or by item name in Arabic or English.

You can also search for GS1, and you can calculate the approximate customs."""

def load_users():
    if not os.path.exists(USERS_FILE):
        return {}
    try:
        with open(USERS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return {}

def save_users(users):
    with open(USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(users, f, ensure_ascii=False, indent=2)

def generate_code(users):
    while True:
        code = "USR-" + "".join(random.choices(string.digits, k=4))
        if all(u.get("code") != code for u in users.values()):
            return code

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.message.from_user
    users = load_users()
    uid = str(user.id)
    now = time.time()
    if uid in users and users[uid].get("status") == "rejected":
        last = users[uid].get("rejected_at", 0)
        remaining = int((REJECT_WAIT - (now - last)) / 60) + 1
        if now - last < REJECT_WAIT:
            await update.message.reply_text(f"Your request was rejected. You can send again after {remaining} minutes.")
            return
        users[uid]["status"] = "pending"
        users[uid]["requested_at"] = now
        save_users(users)
    if uid in users and users[uid].get("status") == "accepted":
        await update.message.reply_text(WELCOME_MESSAGE + "\n\n✅ Your account is accepted. Send the item code or item name.")
        return
    if uid not in users:
        code = generate_code(users)
        users[uid] = {
            "code": code,
            "name": user.first_name or "",
            "username": user.username or "",
            "id": user.id,
            "status": "pending",
            "requested_at": now
        }
        save_users(users)
    else:
        code = users[uid]["code"]
        users[uid]["status"] = "pending"
        users[uid]["requested_at"] = now
        save_users(users)
    await update.message.reply_text(WELCOME_MESSAGE + f"\n\nYour Code: {code}\nWaiting for acceptance from Captain Emad.")
    if ADMIN_ID:
        keyboard = [[
            InlineKeyboardButton("✅ قبول / Accept", callback_data=f"accept_{uid}"),
            InlineKeyboardButton("❌ رفض / Reject", callback_data=f"reject_{uid}")
        ]]
        try:
            await context.bot.send_message(
                chat_id=int(ADMIN_ID),
                text=f"📦 New User Request\n\nCode: {code}\nName: {user.first_name}\nUsername: @{user.username}\nID: {user.id}",
                reply_markup=InlineKeyboardMarkup(keyboard)
            )
        except Exception as e:
            print(f"Admin send error: {e}")

async def button_click(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    users = load_users()
    data = query.data
    if data.startswith("accept_"):
        uid = data.split("_")[1]
        if uid in users:
            users[uid]["status"] = "accepted"
            save_users(users)
            await context.bot.send_message(chat_id=int(uid), text=WELCOME_MESSAGE + "\n\n✅ Your request has been accepted by Captain Emad.")
            await query.message.reply_text(f"Accepted ✅\nCode: {users[uid]['code']} - {users[uid]['name']}")
    elif data.startswith("reject_"):
        uid = data.split("_")[1]
        if uid in users:
            users[uid]["status"] = "rejected"
            users[uid]["rejected_at"] = time.time()
            save_users(users)
            await context.bot.send_message(chat_id=int(uid), text="❌ Your request was rejected by Captain Emad. You can send again after 5 minutes.")
            await query.message.reply_text(f"Rejected ❌\nCode: {users[uid]['code']} - {users[uid]['name']}")

async def save_doc(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if str(update.message.from_user.id) != str(ADMIN_ID):
        return
    caption = (update.message.caption or "").strip() or "بدون اسم"
    file = None
    ext = "jpg"
    if update.message.photo:
        file = await update.message.photo[-1].get_file()
    elif update.message.document:
        file = await update.message.document.get_file()
        ext = update.message.document.file_name.split(".")[-1] if "." in update.message.document.file_name else "pdf"
    else:
        return
    docs = load_docs()
    num = len(docs) + 1
    filename = f"docs/{num}.{ext}"
    await file.download_to_drive(filename)
    docs.append({"num": num, "name": caption, "path": filename})
    save_docs(docs)
    await update.message.reply_text(f"✅ اتحفظ برقم {num}\n📄 {caption}\n\nعشان تنزله اكتب: تزيل {num}")

async def list_docs(update: Update, context: ContextTypes.DEFAULT_TYPE):
    docs = load_docs()
    if not docs:
        await update.message.reply_text("مفيش مستندات لسه")
        return
    txt = "📁 المستندات:\n\n"
    for d in docs:
        txt += f"{d['num']} - {d['name']}\n"
    txt += "\nاكتب: تزيل 1"
    await update.message.reply_text(txt)

async def download_doc(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.message.from_user
    uid = str(user.id)
    if str(user.id) != str(ADMIN_ID):
        users = load_users()
        if uid not in users or users[uid].get("status") != "accepted":
            await update.message.reply_text("Please send /start first.")
            return
    txt = update.message.text.strip()
    m = re.search(r'(\d+)', txt)
    if not m:
        await update.message.reply_text("اكتب: تزيل 1")
        return
    num = int(m.group(1))
    docs = load_docs()
    for d in docs:
        if d["num"] == num:
            try:
                await update.message.reply_document(document=open(d["path"], "rb"), caption=f"📄 رقم {num} - {d['name']}")
            except FileNotFoundError:
                await update.message.reply_text(f"❌ الملف رقم {num} مش موجود على السيرفر، ارفعه تاني")
            return
    await update.message.reply_text(f"❌ مفيش مستند رقم {num}\nاكتب /docs عشان تشوف الارقام")

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.message.from_user
    uid = str(user.id)
    text = update.message.text or ""
    if "تزيل" in text or "تنزيل" in text:
        return await download_doc(update, context)
    users = load_users()
    if str(user.id) == str(ADMIN_ID):
        return
    if uid not in users:
        await update.message.reply_text("Please send /start first.")
        return
    if users[uid].get("status") == "pending":
        await update.message.reply_text("Your request is still waiting for acceptance from Captain Emad.")
        return
    if users[uid].get("status") == "rejected":
        await update.message.reply_text("Your request was rejected. Send /start again after 5 minutes.")
        return
    if ADMIN_ID:
        await context.bot.send_message(chat_id=int(ADMIN_ID), text=f"Inquiry from {users[uid]['code']} - {users[uid]['name']}:\n{update.message.text}")
    await update.message.reply_text("Received. Captain Emad will reply to you.")

async def list_users(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if str(update.message.from_user.id) != str(ADMIN_ID):
        return
    users = load_users()
    if not users:
        await update.message.reply_text("No users yet.")
        return
    text = f"Total Users: {len(users)}\n\n"
    for u in users.values():
        text += f"Code: {u['code']} | Name: {u['name']} | @{u.get('username','')} | ID: {u['id']} | Status: {u['status']}\n"
    await update.message.reply_text(text)

async def delete_user(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if str(update.message.from_user.id) != str(ADMIN_ID):
        return
    if not context.args:
        await update.message.reply_text("Usage: /delete USR-1234  or  /delete <ID>")
        return
    target = context.args[0]
    users = load_users()
    found = None
    for uid, u in list(users.items()):
        if u.get("code") == target or str(u.get("id")) == target:
            found = uid
            break
    if not found:
        await update.message.reply_text("User not found.")
        return
    del users[found]
    save_users(users)
    await update.message.reply_text(f"Deleted user {target} ✅")

async def stats(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if str(update.message.from_user.id) != str(ADMIN_ID):
        return
    users = load_users()
    accepted = 0
    pending = 0
    rejected = 0
    for u in users.values():
        s = u.get("status")
        if s == "accepted":
            accepted += 1
        elif s == "pending":
            pending += 1
        elif s == "rejected":
            rejected += 1
    await update.message.reply_text(f"Total: {len(users)}\nAccepted: {accepted}\nPending: {pending}\nRejected: {rejected}")

def main():
    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("users", list_users))
    app.add_handler(CommandHandler("delete", delete_user))
    app.add_handler(CommandHandler("stats", stats))
    app.add_handler(CommandHandler("docs", list_docs))
    app.add_handler(CallbackQueryHandler(button_click))
    app.add_handler(MessageHandler(filters.PHOTO | filters.Document.ALL, save_doc))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    print("Bot is running...")
    app.run_polling()

if __name__ == "__main__":
    main()