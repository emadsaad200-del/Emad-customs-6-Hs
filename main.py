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

EXCEL_FILES = [
    "gs1xlsx.xlsx",
    "GPC as of May 2023 v20230605 EN.xlsx",
]

# ===== GPC =====
BRICK_INDEX = {}
BRICK_BY_CODE = {}
BRICK_ATTRS = {}

def build_index():
    print("Building light index...")
    for fname in EXCEL_FILES:
        if not os.path.exists(fname):
            continue
        try:
            wb = openpyxl.load_workbook(fname, read_only=True, data_only=True)
            ws = wb["Schema"] if "Schema" in wb.sheetnames else wb[wb.sheetnames[0]]
            for r in ws.iter_rows(values_only=True):
                try:
                    if not r or len(r) < 11:
                        continue
                    if str(r[0]).lower().startswith("segment"):
                        continue
                    brick_code = str(r[9]) if r[9] is not None else ""
                    brick_title = str(r[10]) if r[10] is not None else ""
                    if not brick_title:
                        continue
                    key = brick_title.strip().lower()
                    if key not in BRICK_INDEX:
                        BRICK_INDEX[key] = {
                            "brick_code": brick_code, "brick_title": brick_title,
                            "segment": r[1], "segment_code": str(r[0]),
                            "family": r[4], "family_code": str(r[3]),
                            "class": r[7], "class_code": str(r[6]),
                            "includes": r[11] or "", "excludes": r[12] if len(r) > 12 else "" or "",
                        }
                        if brick_code:
                            BRICK_BY_CODE[brick_code] = key
                    if len(r) > 14 and r[13] is not None:
                        attr_code = str(r[13])
                        attr_title = r[14]
                        val_title = str(r[17]) if len(r) > 17 and r[17] is not None else None
                        if key not in BRICK_ATTRS:
                            BRICK_ATTRS[key] = []
                        ex = None
                        for a in BRICK_ATTRS[key]:
                            if a["attribute_code"] == attr_code:
                                ex = a; break
                        if not ex:
                            ex = {"attribute_code": attr_code, "attribute_title": attr_title, "values": []}
                            BRICK_ATTRS[key].append(ex)
                        if val_title and val_title not in ex["values"]:
                            ex["values"].append(val_title)
                except:
                    continue
            wb.close()
            print(f"Indexed {fname}: {len(BRICK_INDEX)} bricks so far")
        except Exception as e:
            print(f"Index error {fname}: {e}")
    print(f"Index done: {len(BRICK_INDEX)} bricks")

build_index()

# ===== Cities - ملف المدن =====
CITIES_FILE = "cities.xlsx"
CITY_INDEX = {}
CITY_BY_CODE = {}

def build_city_index():
    if not os.path.exists(CITIES_FILE):
        print("cities.xlsx not found")
        return
    try:
        wb = openpyxl.load_workbook(CITIES_FILE, read_only=True, data_only=True)
        ws = wb["City List"] if "City List" in wb.sheetnames else wb[wb.sheetnames[0]]
        for r in ws.iter_rows(values_only=True):
            try:
                if not r or len(r) < 2 or not r[0] or not r[1]:
                    continue
                if str(r[0]).lower() == "citycode":
                    continue
                code = str(r[0]).strip().upper()
                name = str(r[1]).strip()
                country = str(r[2]).strip() if len(r) > 2 and r[2] else ""
                CITY_BY_CODE[code] = {"code": code, "name": name, "country": country}
                key = name.lower()
                if key not in CITY_INDEX:
                    CITY_INDEX[key] = []
                CITY_INDEX[key].append({"code": code, "name": name, "country": country})
            except:
                continue
        wb.close()
        print(f"Cities indexed: {len(CITY_BY_CODE)}")
    except Exception as e:
        print(f"City index error: {e}")

build_city_index()

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

def search_gpc_fast(query):
    q = normalize(query)
    if not q:
        return None, None
    q_digits = re.sub(r'\D', '', q)
    if q_digits and q_digits in BRICK_BY_CODE:
        key = BRICK_BY_CODE[q_digits]
        return BRICK_INDEX[key], None
    if q in BRICK_INDEX:
        return BRICK_INDEX[q], None
    matches = [BRICK_INDEX[k] for k in BRICK_INDEX.keys() if q in k]
    if len(matches) == 1:
        return matches[0], None
    if len(matches) > 1:
        return ("multiple", matches[:10]), None
    return None, None

def get_attrs_for_brick(brick_title):
    return BRICK_ATTRS.get(brick_title.strip().lower(), [])

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

# ===== البحث في المدن =====
def search_city(query):
    q = query.strip()
    qu = q.upper().strip()
    if qu in CITY_BY_CODE:
        return CITY_BY_CODE[qu], None
    ql = q.lower()
    matches = []
    for key, lst in CITY_INDEX.items():
        if ql in key:
            matches.extend(lst)
            if len(matches) >= 10:
                break
    if len(matches) == 1:
        return matches[0], None
    if len(matches) > 1:
        return ("multiple", matches[:10]), None
    return None, None

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
    try:
        await query.answer()
    except Exception:
        pass
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
            if not keys and context.bot_data.get(f"gpc_{query.from_user.id}"):
                keys = context.bot_data.get(f"gpc_{query.from_user.id}", [])
            if 0 <= idx < len(keys):
                b = keys[idx]
                attrs = get_attrs_for_brick(b["brick_title"])
                await query.message.reply_text(format_brick(b, attrs))
            else:
                await query.message.reply_text("القايمة دي انتهت، ابعت /gpc تاني واختار من جديد")
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
    brick, _ = search_gpc_fast(query)
    if brick is None:
        await update.message.reply_text(f"❌ ملقيتش صنف اسمه: {query}")
        return
    if isinstance(brick, tuple) and brick[0] == "multiple":
        matches = brick[1]
        context.user_data["gpc_matches"] = matches
        context.bot_data[f"gpc_{update.message.from_user.id}"] = matches
        keyboard = [[InlineKeyboardButton(m["brick_title"][:60], callback_data=f"gpc_{i}")] for i, m in enumerate(matches)]
        await update.message.reply_text(f"لقيت {len(matches)} صنف، اختار واحد:", reply_markup=InlineKeyboardMarkup(keyboard))
        return
    attrs = get_attrs_for_brick(brick["brick_title"])
    await update.message.reply_text(format_brick(brick, attrs))

# ===== امر المدن =====
async def city_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    users = load_users()
    uid = str(update.message.from_user.id)
    if str(update.message.from_user.id) != str(ADMIN_ID):
        if uid not in users or users[uid].get("status") != "accepted":
            await update.message.reply_text("Please send /start first.")
            return
    if not context.args:
        await update.message.reply_text("اكتب: /city اسم المدينة او الكود\nمثال: /city Cairo\nاو: /city EGCAI\nاو: /city Alexandria")
        return
    query = " ".join(context.args)
    result, _ = search_city(query)
    if result is None:
        await update.message.reply_text(f"❌ ملقيتش مدينة اسمها: {query}")
        return
    if isinstance(result, tuple) and result[0] == "multiple":
        matches = result[1]
        txt = f"لقيت {len(matches)} مدينة:\n\n"
        for c in matches:
            txt += f"🏙️ {c['name']}\n🔢 {c['code']} - {c['country']}\n\n"
        await update.message.reply_text(txt)
        return
    await update.message.reply_text(f"🏙️ {result['name']}\n🔢 City Code: {result['code']}\n🌍 Country: {result['country']}")

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.message.from_user
    uid = str(user.id)
    text = update.message.text or ""
    if is_download_request(text):
        return await download_doc(update, context)
    # لو كتب city + اسم المدينة من غير سلاش
    low = text.strip().lower()
    if low.startswith("city ") or low.startswith("مدينة ") or low.startswith("ميناء "):
        q = text.strip().split(" ", 1)[1] if " " in text.strip() else ""
        if q:
            result, _ = search_city(q)
            if result is not None:
                if isinstance(result, tuple) and result[0] == "multiple":
                    txt = f"لقيت {len(result[1])} مدينة:\n\n"
                    for c in result[1]:
                        txt += f"🏙️ {c['name']}\n🔢 {c['code']} - {c['country']}\n\n"
                    await update.message.reply_text(txt)
                    return
                await update.message.reply_text(f"🏙️ {result['name']}\n🔢 City Code: {result['code']}\n🌍 Country: {result['country']}")
                return
    users = load_users()
    if str(user.id) == str(ADMIN_ID):
        if len(text.strip()) >= 3:
            brick, _ = search_gpc_fast(text)
            if brick is not None:
                if isinstance(brick, tuple) and brick[0] == "multiple":
                    matches = brick[1]
                    context.user_data["gpc_matches"] = matches
                    context.bot_data[f"gpc_{user.id}"] = matches
                    keyboard = [[InlineKeyboardButton(m["brick_title"][:60], callback_data=f"gpc_{i}")] for i, m in enumerate(matches)]
                    await update.message.reply_text(f"لقيت {len(matches)} صنف قريب من: {text}\nاختار واحد:", reply_markup=InlineKeyboardMarkup(keyboard))
                    return
                attrs = get_attrs_for_brick(brick["brick_title"])
                await update.message.reply_text(format_brick(brick, attrs))
                return
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
        brick, _ = search_gpc_fast(text)
        if brick is not None:
            if isinstance(brick, tuple) and brick[0] == "multiple":
                matches = brick[1]
                context.user_data["gpc_matches"] = matches
                context.bot_data[f"gpc_{user.id}"] = matches
                keyboard = [[InlineKeyboardButton(m["brick_title"][:60], callback_data=f"gpc_{i}")] for i, m in enumerate(matches)]
                await update.message.reply_text(f"لقيت {len(matches)} صنف قريب من: {text}\nاختار واحد:", reply_markup=InlineKeyboardMarkup(keyboard))
                return
            attrs = get_attrs_for_brick(brick["brick_title"])
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
    app.add_handler(CommandHandler("city", city_command))
    app.add_handler(CallbackQueryHandler(button_click))
    app.add_handler(MessageHandler(filters.PHOTO | filters.Document.ALL, save_doc))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    print(f"Bot is running... fast index, {len(BRICK_INDEX)} bricks, {len(CITY_BY_CODE)} cities")
    app.run_polling()

if __name__ == "__main__":
    main()
