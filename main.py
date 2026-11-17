import os, json, time, random, string, re
import openpyxl
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, CallbackQueryHandler, ContextTypes, filters

BOT_TOKEN = os.getenv("BOT_TOKEN")
ADMIN_ID = os.getenv("ADMIN_ID")
USERS_FILE = "users.json"
REJECT_WAIT = 5 * 60

DOCS_FILE = "docs_index.json"
os.makedirs("docs", exist_ok=True)

# ملفين بس - شلنا gs1 exle.xlsx
EXCEL_FILES = [
    "gs1xlsx.xlsx",
    "GPC as of May 2023 v20230605 EN.xlsx",
]

BRICKS = {}
BRICK_BY_CODE = {}
ATTRIBUTES = {}

def load_gpc():
    print("Loading GPC Excel files...")
    for fname in EXCEL_FILES:
        if not os.path.exists(fname):
            print(f"Missing: {fname}")
            continue
        try:
            wb = openpyxl.load_workbook(fname, read_only=True, data_only=True)
            ws = wb["Schema"] if "Schema" in wb.sheetnames else wb[wb.sheetnames[0]]
            rows = list(ws.iter_rows(values_only=True))
            print(f"Loaded {fname}: {len(rows)} rows")
            for r in rows[1:]:
                try:
                    if len(r) < 11:
                        continue
                    seg_code, seg_title = r[0], r[1]
                    fam_code, fam_title = r[3], r[4]
                    cls_code, cls_title = r[6], r[7]
                    brick_code, brick_title = r[9], r[10]
                    includes, excludes = r[11], r[12] if len(r) > 12 else ""
                    attr_code = r[13] if len(r) > 13 else None
                    attr_title = r[14] if len(r) > 14 else None
                    val_title = r[17] if len(r) > 17 else None
                    if not brick_title:
                        continue
                    key = str(brick_title).strip().lower()
                    if key not in BRICKS:
                        BRICKS[key] = {
                            "brick_code": str(brick_code),
                            "brick_title": brick_title,
                            "segment": seg_title, "segment_code": str(seg_code),
                            "family": fam_title, "family_code": str(fam_code),
                            "class": cls_title, "class_code": str(cls_code),
                            "includes": includes or "",
                            "excludes": excludes or "",
                        }
                        if brick_code:
                            BRICK_BY_CODE[str(brick_code)] = key
                    if attr_code:
                        if key not in ATTRIBUTES:
                            ATTRIBUTES[key] = []
                        found = None
                        for a in ATTRIBUTES[key]:
                            if a["attribute_code"] == str(attr_code):
                                found = a; break
                        if not found:
                            found = {"attribute_code": str(attr_code), "attribute_title": attr_title, "values": []}
                            ATTRIBUTES[key].append(found)
                        if val_title and str(val_title) not in found["values"]:
                            found["values"].append(str(val_title))
                except:
                    continue
            wb.close()
        except Exception as e:
            print(f"Error loading {fname}: {e}")
    print(f"GPC Done: {len(BRICKS)} bricks")

load_gpc()

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

# ===== تصليح التزيل =====
def clean_text(text):
    if not text:
        return ""
    for ch in ["\u200e", "\u200f", "\u202a", "\u202b", "\u202c", "\u202d", "\u202e", "\u200c", "\u200d", "\u0640", "ـ"]:
        text = text.replace(ch, "")
    arabic_digits = "٠١٢٣٤٥٦٧٨٩"
    for i, d in enumerate(arabic_digits):
        text = text.replace(d, str(i))
    text = text.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا").replace("ة", "ه")
    return text.strip()

def is_download_request(text):
    t = clean_text(text)
    return bool(re.search(r'(تزيل|تنزيل|تحميل|download|نزل)', t))

def extract_number(text):
    t = clean_text(text)
    m = re.search(r'(\d+)', t)
    return int(m.group(1)) if m else None

def normalize(text):
    return re.sub(r'\s+', ' ', (text or "").strip().lower())

def search_gpc(query):
    q = normalize(query)
    if not q:
        return None, None
    q_digits = re.sub(r'\D', '', q)
    if q_digits and q_digits in BRICK_BY_CODE:
        key = BRICK_BY_CODE[q_digits]
        return BRICKS[key], ATTRIBUTES.get(key, [])
    if q in BRICKS:
        return BRICKS[q], ATTRIBUTES.get(q, [])
    matches = [k for k in BRICKS.keys() if q in k]
    if len(matches) == 1:
        return BRICKS[matches[0]], ATTRIBUTES.get(matches[0], [])
    if len(matches) > 1:
        return ("multiple", matches[:10]), None
    return None, None

def format_brick(brick, attrs):
    txt = f"📦 {brick['brick_title']}\n"
    txt += f"🔢 Brick Code: {brick['brick_code']}\n\n"
    txt += f"📂 Segment: {brick['segment']} ({brick['segment_code']})\n"
    txt += f"📁 Family: {brick['family']} ({brick['family_code']})\n"
    txt += f"📄 Class: {brick['class']} ({brick['class_code']})\n\n"
    if brick.get("includes"):
        txt += f"✅ Includes:\n{brick['includes']}\n\n"
    if brick.get("excludes"):
        txt += f"⛔ Excludes:\n{brick['excludes']}\n\n"
    if attrs:
        txt += "⚙️ Attributes:\n"
        for a in attrs:
            txt += f"\n• {a['attribute_title']} ({a['attribute_code']})\n"
            if a.get("values"):
                vals = a["values"][:15]
                txt += "  " + ", ".join(vals)
                if len(a["values"]) > 15:
                    txt += f" ... (+{len(a['values'])-15} more)"
                txt += "\n"
    return txt

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
        users[uid] = {"code": code, "name": user.first_name or "", "username": user.username or "", "id": user.id, "status": "pending", "requested_at": now}
        save_users(users)
    else:
        code = users[uid]["code"]
        users[uid]["status"] = "pending"
        users[uid]["requested_at"] = now
        save_users(users)
    await update.message.reply_text(WELCOME_MESSAGE + f"\n\nYour Code: {code}\nWaiting for acceptance from Captain Emad.")
    if ADMIN_ID:
        keyboard = [[InlineKeyboardButton("✅ قبول / Accept", callback_data=f"accept_{uid}"), InlineKeyboardButton("❌ رفض / Reject", callback_data=f"reject_{uid}")]]
        try:
            await context.bot.send_message(chat_id=int(ADMIN_ID), text=f"📦 New User Request\n\nCode: {code}\nName: {user.first_name}\nUsername: @{user.username}\nID: {user.id}", reply_markup=InlineKeyboardMarkup(keyboard))
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
    elif data.startswith("gpc_"):
        try:
            idx = int(data[4:])
            keys = context.user_data.get("gpc_matches", [])
            if 0 <= idx < len(keys):
                k = keys[idx]
                await query.message.reply_text(format_brick(BRICKS[k], ATTRIBUTES.get(k, [])))
        except Exception as e:
            print(f"gpc callback error: {e}")

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
    num = max([d["num"] for d in docs], default=0) + 1
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
    num = extract_number(update.message.text)
    if num is None:
        await update.message.reply_text("اكتب: تزيل 1")
        return
    docs = load_docs()
    for d in docs:
        if d["num"] == num:
            try:
                await update.message.reply_document(document=open(d["path"], "rb"), caption=f"📄 رقم {num} - {d['name']}")
            except FileNotFoundError:
                await update.message.reply_text(f"❌ الملف رقم {num} مش موجود على السيرفر، ارفعه تاني")
            return
    await update.message.reply_text(f"❌ مفيش مستند رقم {num}\nاكتب /docs عشان تشوف الارقام")

async def gpc_search_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    users = load_users()
    uid = str(update.message.from_user.id)
    if str(update.message.from_user.id) != str(ADMIN_ID):
        if uid not in users or users[uid].get("status") != "accepted":
            await update.message.reply_text("Please send /start first.")
            return
    if not context.args:
        await update.message.reply_text("اكتب: /gpc اسم الصنف او الكود\nمثال: /gpc Airbrushes\nاو: /gpc 10001686")
        return
    query = " ".join(context.args)
    brick, attrs = search_gpc(query)
    if brick is None:
        await update.message.reply_text(f"❌ ملقيتش صنف اسمه: {query}")
        return
    if isinstance(brick, tuple) and brick[0] == "multiple":
        matches = brick[1]
        context.user_data["gpc_matches"] = matches
        keyboard = [[InlineKeyboardButton(BRICKS[k]["brick_title"][:60], callback_data=f"gpc_{i}")] for i, k in enumerate(matches)]
        await update.message.reply_text(f"لقيت {len(matches)} صنف، اختار واحد:", reply_markup=InlineKeyboardMarkup(keyboard))
        return
    await update.message.reply_text(format_brick(brick, attrs))

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.message.from_user
    uid = str(user.id)
    text = update.message.text or ""
    # التزيل اول حاجة قبل GPC وقبل اي حاجة
    if is_download_request(text):
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
    if len(text.strip()) >= 3:
        brick, attrs = search_gpc(text)
        if brick is not None:
            if isinstance(brick, tuple) and brick[0] == "multiple":
                matches = brick[1]
                context.user_data["gpc_matches"] = matches
                keyboard = [[InlineKeyboardButton(BRICKS[k]["brick_title"][:60], callback_data=f"gpc_{i}")] for i, k in enumerate(matches)]
                await update.message.reply_text(f"لقيت {len(matches)} صنف قريب من: {text}\nاختار واحد:", reply_markup=InlineKeyboardMarkup(keyboard))
                return
            await update.message.reply_text(format_brick(brick, attrs))
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
    accepted = sum(1 for u in users.values() if u.get("status") == "accepted")
    pending = sum(1 for u in users.values() if u.get("status") == "pending")
    rejected = sum(1 for u in users.values() if u.get("status") == "rejected")
    await update.message.reply_text(f"Total: {len(users)}\nAccepted: {accepted}\nPending: {pending}\nRejected: {rejected}")

def main():
    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("users", list_users))
    app.add_handler(CommandHandler("delete", delete_user))
    app.add_handler(CommandHandler("stats", stats))
    app.add_handler(CommandHandler("docs", list_docs))
    app.add_handler(CommandHandler("gpc", gpc_search_command))
    app.add_handler(CallbackQueryHandler(button_click))
    app.add_handler(MessageHandler(filters.PHOTO | filters.Document.ALL, save_doc))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    print(f"Bot is running... GPC bricks loaded: {len(BRICKS)}")
    app.run_polling()

if __name__ == "__main__":
    main()
