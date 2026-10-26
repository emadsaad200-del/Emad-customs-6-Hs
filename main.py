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

def search_items(query):
    query = query.strip()
    digits = "".join([ch for ch in query if ch.isdigit()])
    q_lower = query.lower()
    
    results = []
    
    # 1 - لو بعت كود رقمي - بحث دقيق بالكود
    if digits and len(digits) >= 2:
        for item in items:
            if item["clean"] == digits:
                return [item]  # مطابق تماما
        for item in items:
            if item["clean"].startswith(digits) or digits.startswith(item["clean"]):
                results.append(item)
        if results:
            return results[:10]
    
    # 2 - بحث بالاسم / الكلمة (عربي)
    # شيل المسافات الزيادة
    q_words = q_lower.split()
    q_words = [w for w in q_words if len(w) > 1]
    if not q_words:
        return []
    
    scored = []
    for item in items:
        desc_lower = item["desc"].lower()
        score = 0
        # لو الجملة كلها موجودة
        if q_lower in desc_lower:
            score += 100
        # عدد الكلمات المتطابقة
        for w in q_words:
            if w in desc_lower:
                score += 10
        if score > 0:
            scored.append((score, item))
    
    # رتب بالاعلى
    scored.sort(key=lambda x: x[0], reverse=True)
    return [x[1] for x in scored[:10]]

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
    await update.message.reply_text("اهلا بيك - تقدر تبحث ب:\n- الكود: 4405 - 5506 - 84.03\n- الكلمة: زجاج - خشب - نسيج - قطن\n- جزء من الوصف")

async def handle_search(update, context):
    ok = await request_approval(update, context)
    if not ok:
        return
    text = update.message.text.strip()
    if len(text) < 2:
        await update.message.reply_text("اكتب كلمة اطول")
        return
    
    found = search_items(text)
    
    if not found:
        await update.message.reply_text("مفيش نتائج لـ: " + text)
        return
    
    # لو نتيجة واحدة اعرضها كاملة
    if len(found) == 1:
        item = found[0]
        await update.message.reply_text("البند: " + item['code'] + "\n" + item['desc'])
        return
    
    # لو كذا نتيجة اعرضهم مختصرين
    msg = "لقيت " + str(len(found)) + " نتائج لـ " + text + ":\n\n"
    for i, item in enumerate(found, 1):
        short = item['desc'][:70] + "..." if len(item['desc']) > 70 else item['desc']
        msg += str(i) + "- [" + item['code'] + "] " + short + "\n\n"
    msg += "ابعت الكود اللي عايزه بالظبط عشان اجيبلك الوصف الكامل"
    await update.message.reply_text(msg)

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
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_search))
    app.run_polling()
