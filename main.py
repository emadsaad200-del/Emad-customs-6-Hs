import os, json, time, random, string
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, CallbackQueryHandler, ContextTypes, filters

BOT_TOKEN = os.getenv("BOT_TOKEN")
ADMIN_ID = os.getenv("ADMIN_ID")
USERS_FILE = "users.json"
REJECT_WAIT = 5 * 60  # 5 minutes

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

    # If user exists and rejected and less than 5 min -> block
    if uid in users and users[uid].get("status") == "rejected":
        last = users[uid].get("rejected_at", 0)
        remaining = int((REJECT_WAIT - (now - last)) / 60) + 1
        if now - last < REJECT_WAIT:
            await update.message.reply_text(f"Your request was rejected. You can send again after {remaining} minutes.")
            return
        # After 5 min, allow again -> reset to pending
        users[uid]["status"] = "pending"
        users[uid]["requested_at"] = now
        save_users(users)

    # If already accepted
    if uid in users and users[uid].get("status") == "accepted":
        await update.message.reply_text(WELCOME_MESSAGE + "\n\n✅ Your account is accepted. Send the item code or item name.")
        return

    # New user or re-request
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

    # Send welcome + waiting to user
    await update.message.reply_text(WELCOME_MESSAGE + f"\n\nYour Code: {code}\nWaiting for acceptance from Captain Emad.")

    # Send to admin with 2 buttons
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

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.message.from_user
    uid = str(user.id)
    users = load_users()
    # Admin messages are commands, ignore here
    if str(user.id) == str(ADMIN_ID):
        return
    # If not registered, tell to start
    if uid not in users:
        await update.message.reply_text("Please send /start first.")
        return
    if users[uid].get("status") == "pending":
        await update.message.reply_text("Your request is still waiting for acceptance from Captain Emad.")
        return
    if users[uid].get("status") == "rejected":
        await update.message.reply_text("Your request was rejected. Send /start again after 5 minutes.")
        return
    # Accepted user -> forward to admin as inquiry
    if ADMIN_ID:
        await context.bot.send_message(chat_id=int(ADMIN_ID), text=f"Inquiry from {users[uid]['code']} - {users[uid]['name']}:\n{update.message.text}")
    await update.message.reply_text("Received. Captain Emad will reply to you.")

# --- Admin Commands ---
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
    accepted = sum(1 for u in users.values() if u.get("status")=="accepted")
    pending = sum(1 for u in users.values() if u.get("status")=="pending")
    rejected = sum(1 for u in users.values() if u.get("status")=="rejected")
    await update.message.reply_text(f"Total: {len(users)}\nAccepted: {accepted}\nPending: {pending}\nRejected: {rejected}")

def main():
    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("users", list_users))
    app.add_handler(CommandHandler("delete", delete_user))
    app.add_handler(CommandHandler("stats", stats))
    app.add_handler(CallbackQueryHandler(button_click))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    print("Bot is running...")
    app.run_polling()

if __name__ == "__main__":
    main()