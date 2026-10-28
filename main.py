import os
import json
import re
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
    if not t: return False
    cleaned = t.replace('.','').replace(' ','')
    if not cleaned.isdigit(): return False
    if len(cleaned) < 2 or len(cleaned) > 12: return False
    return True

def clean_desc(text):
    if not text: 
        return ""
    t = str(text).strip()
    # شيل التطويلة
    t = t.replace('ـ','')
    
    # === إصلاحات حرف العين والغين اللي ظهرت في مثال 30.03 ===
    # إصلاحات مباشرة من رسالتك
    t = t.replace('مكونعين','مكونين')
    t = t.replace('أكعثر','أكثر')
    t = t.replace('كعثر','كثر')
    t = t.replace('أرلفة','أغلفة')
    t = t.replace('الحمعل','الحمل')
    t = t.replace('أدويعة','أدوية')
    t = t.replace('زجععاج','زجاج').replace('زجعاج','زجاج')
    t = t.replace('ع ع ع','').replace('ع ع','')
    
    # رير -> غير (بس لو كلمة لوحدها عشان ما نبوزش حرير / صغير)
    t = re.sub(r'\bرير\b', 'غير', t)
    t = re.sub(r'\bريرها\b', 'غيرها', t)
    t = re.sub(r'\bرير مشربة\b', 'غير مشربة', t)
    
    # إصلاحات عامة للعين الزيادة: نعين في آخر الكلمة -> نين
    # مثال: مكونعين -> مكونين (لو فات من اللي فوق)
    t = re.sub(r'(\w)عين\b', r'\1ين', t)
    # ويعة -> وية (أدويعة -> أدوية)
    t = re.sub(r'ويعة\b', 'وية', t)
    
    return ' '.join(t.split())

def find_code_and_desc(row):
    code = None
    for val in reversed(list(row)):
        s = str(val).strip()
        if s == 'nan' or s == '': continue
        if is_hs_code(s):
            if len(s.replace('.','')) == 1: continue
            code = s
            break
    longest = ""
    for val in row:
        s = str(val).strip()
        if s == 'nan' or s == '': continue
        if has_arabic(s) and len(s) > len(longest) and len(s) > 5:
            longest = s
    return code, longest

# === Load EXLL ===
print("Loading EXLL.xlsx ...", flush=True)
df = pd.read_excel("EXLL.xlsx", header=None)
items = []
for idx in range(len(df)):
    row = df.iloc[idx]
    c,d = find_code_and_desc(row)
    if c is None or d == "": continue
    if len(d) > 400: continue
    if "جمهورية مصر" in d or "وزارة المالية" in d: continue
    if c == '2007': continue
    cd = clean_desc(d)
    items.append({'code': c, 'desc': cd, 'clean': c.replace('.','').replace(' ','')})

print(f"Loaded {len(items)} EXLL items", flush=True)

# === Load Tariff ===
tariff_items = []
tariff_map = {}
try:
    tariff_file = None
    for name in ["Tariff.xlsx", "Tariff.xlsx.csv", "tariff.xlsx", "Tariff.csv"]:
        if os.path.exists(name):
            tariff_file = name
            break
    if tariff_file:
        print(f"Loading {tariff_file} ...", flush=True)
        df_t = pd.read_excel(tariff_file, dtype=str, header=None)
        start = 1 if "رمز النظام" in str(df_t.iloc[0,0]) else 0
        for i in range(start, len(df_t)):
            row = df_t.iloc[i]
            try:
                code_raw = str(row.iloc[0]).strip()
                if code_raw == 'nan' or code_raw == '': continue
                code_clean = code_raw.replace('.','').strip()
                if not code_clean.isdigit(): continue
                arabic = str(row.iloc[1]) if len(row) > 1 else ""
                duty = str(row.iloc[3]) if len(row) > 3 else ""
                proc = str(row.iloc[4]) if len(row) > 4 else ""
                if duty == 'nan': duty = ""
                if proc == 'nan': proc = ""
                if arabic == 'nan': arabic = ""
                # نضف وصف التعريفة كمان
                arabic = clean_desc(arabic)
                item = {'code': code_raw, 'clean': code_clean, 'desc': arabic, 'duty': duty, 'proc': proc}
                tariff_items.append(item)
                tariff_map[code_clean] = duty
            except:
                continue
        print(f"Loaded {len(tariff_items)} Tariff items", flush=True)
    else:
        print("No Tariff file found", flush=True)
except Exception as e:
    print(f"Tariff load error: {e}", flush=True)

def is_allowed(uid):
    return uid in allowed_users or uid == ADMIN_ID

def search_combined(query):
    query = query.strip()
    q_lower = query.lower()
    digits = "".join([ch for ch in query if ch.isdigit()])
    
    if digits and len(digits) >= 2:
        for it in items:
            if it["clean"] == digits:
                duties = [t for t in tariff_items if t["clean"].startswith(digits)]
                return [it], duties[:5]
        matched_tariff = [t for t in tariff_items if t["clean"] == digits]
        if matched_tariff:
            parent = None
            for it in items:
                if digits.startswith(it["clean"]) or it["clean"].startswith(digits[:4]):
                    parent = it
                    break
            return [parent] if parent else [], matched_tariff[:5]
        
        ex_results = [it for it in items if it["clean"].startswith(digits) or digits.startswith(it["clean"])]
        tar_results = [t for t in tariff_items if t["clean"].startswith(digits)]
        if ex_results or tar_results:
            return ex_results[:5], tar_results[:5]

    q_words = [w for w in q_lower.split() if len(w) > 1]
    if not q_words:
        return [], []
    
    scored_ex = []
    for it in items:
        dl = it["desc"].lower()
        score = 100 if q_lower in dl else 0
        for w in q_words:
            if w in dl: score += 10
        if score > 0: scored_ex.append((score, it))
    scored_ex.sort(key=lambda x: x[0], reverse=True)
    
    scored_tar = []
    for it in tariff_items:
        dl = str(it["desc"]).lower()
        score = 100 if q_lower in dl else 0
        for w in q_words:
            if w in dl: score += 10
        if score > 0: scored_tar.append((score, it))
    scored_tar.sort(key=lambda x: x[0], reverse=True)
    
    return [x[1] for x in scored_ex[:7]], [x[1] for x in scored_tar[:7]]

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
    await update.message.reply_text("اهلا بيك\nابحث ب:\n- الكود: 5506 - 0101\n- الاسم: زجاج - خشب - خيول\nالبوت هيجيب الوصف + التعريفة مع بعض")

async def handle_search(update, context):
    ok = await request_approval(update, context)
    if not ok: return
    text = update.message.text.strip()
    if len(text) < 2:
        await update.message.reply_text("اكتب كلمة اطول")
        return
    
    ex_found, tar_found = search_combined(text)
    
    if not ex_found and not tar_found:
        await update.message.reply_text("مفيش نتائج لـ: " + text)
        return
    
    msg = ""
    if ex_found:
        it = ex_found[0]
        msg += "📦 البند: " + it['code'] + "\n" + it['desc'] + "\n\n"
    
    if tar_found:
        msg += "💰 التعريفة:\n"
        for t in tar_found[:5]:
            duty = t['duty'] if t['duty'] else "غير محدد"
            msg += "- " + t['code'] + " : " + duty + " | " + str(t['desc'])[:60] + "\n"
    
    if len(ex_found) > 1:
        msg += "\n🔍 نتائج مشابهة:\n"
        for it in ex_found[1:4]:
            short = it['desc'][:50]
            msg += "- [" + it['code'] + "] " + short + "\n"
    
    if len(msg) > 4000:
        msg = msg[:4000]
    
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
        try: await context.bot.send_message(chat_id=uid, text="تمت الموافقة! ابعت /start")
        except: pass
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