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

print("Starting Bot - Full Fixed Version...", flush=True)

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

# ========= دالة الإصلاح الشاملة لحرف العين - تشمل كل ما سبق =========
def clean_desc(text):
    if not text:
        return ""
    t = str(text).strip()

    # 1- إصلاح الترميز البايظ من الاكسيل نفسه
    try:
        if 'حع' in t or 'إحعادات' in t or 'البعرودة' in t:
            try:
                b = t.encode('latin1', errors='ignore')
                for enc in ['windows-1256', 'cp1256']:
                    try:
                        dec = b.decode(enc)
                        if 'إعدادات' in dec or 'البرودة' in dec:
                            t = dec
                            break
                    except:
                        continue
            except:
                pass
    except:
        pass

    # 2- توحيد الحروف (NFC+NFKC) - يصلح اتصال العين
    t = unicodedata.normalize('NFKC', t)
    t = unicodedata.normalize('NFC', t)
    t = t.replace('ـ', '')

    # 3- قاموس كامل شامل كل الكلمات اللي ظهرت عندك
    full_map = {
        'إعدادات البعرودة': 'إعدادات البرودة',
        'إحعادات البعرودة': 'إعدادات البرودة',
        'إحعادات': 'إعدادات',
        'احعادات': 'إعدادات',
        'إحعدادات': 'إعدادات',
        'البعرودة': 'البرودة',
        'البعروده': 'البرودة',
        'مععدات': 'معدات',
        'مكونعين': 'مكونين',
        'مكونع': 'مكون',
        'أكعثر': 'أكثر',
        'كعثر': 'كثر',
        'أرلفة': 'أغلفة',
        'الحمعل': 'الحمل',
        'أدويعة': 'أدوية',
        'زجععاج': 'زجاج',
        'زجعاج': 'زجاج',
        'الاليات': 'الآلات',
        'بعرد': 'برد',
    }
    # الأطول أولا
    for bad in sorted(full_map, key=len, reverse=True):
        if bad in t:
            t = t.replace(bad, full_map[bad])

    # 4- إصلاح عام للتقطيع
    t = t.replace('حع', 'ع')
    t = re.sub(r'(\w)حع(\w)', r'\1ع\2', t)
    t = re.sub(r'عع+', 'ع', t)
    t = t.replace('ع ع ع','').replace('ع ع','')
    t = re.sub(r'\bرير\b', 'غير', t)
    t = re.sub(r'\bريرها\b', 'غيرها', t)
    t = re.sub(r'\bرير مشربة\b', 'غير مشربة', t)
    t = re.sub(r'(\w)عين\b', r'\1ين', t)
    t = re.sub(r'ويعة\b', 'وية', t)

    return ' '.join(t.split())

def has_arabic(s):
    return any('\u0600' <= c <= '\u06FF' for c in s)

def is_hs_code(s):
    t = s.strip().replace('.','').replace(' ','')
    return t.isdigit() and 2 <= len(t) <= 12

def find_code_and_desc(row):
    code = None
    for val in reversed(row):
        s = str(val).strip() if val is not None else ""
        if s == 'nan' or s == '': continue
        if is_hs_code(s) and len(s.replace('.','')) != 1:
            code = s
            break
    longest = ""
    for val in row:
        s = str(val).strip() if val is not None else ""
        if s == 'nan' or s == '': continue
        if has_arabic(s) and len(s) > len(longest) and len(s) > 5 and len(s) < 400:
            if "جمهورية مصر" not in s and "وزارة المالية" not in s:
                longest = s
    return code, longest

# ========= تحميل خفيف بدون pandas - يحل مشكلة نفاد الذاكرة =========
items = []
print("Loading EXLL.xlsx low RAM...", flush=True)
try:
    from openpyxl import load_workbook
    wb = load_workbook("EXLL.xlsx", read_only=True, data_only=True)
    ws = wb.active
    for row in ws.iter_rows(values_only=True):
        if not row: continue
        c,d = find_code_and_desc(row)
        if c is None or d == "": continue
        if c == '2007': continue
        items.append({'code': c, 'desc': clean_desc(d), 'clean': c.replace('.','').replace(' ','')})
    wb.close()
    gc.collect()
    print(f"Loaded {len(items)} items - RAM OK ~35MB", flush=True)
except Exception as e:
    print(f"Load error: {e}", flush=True)

# Tariff اختياري - لو موجود
tariff_items = []
if os.path.exists("Tariff.json"):
    tariff_items = load_json("Tariff.json", [])

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
        ex_results = [it for it in items if it["clean"].startswith(digits) or digits.startswith(it["clean"])]
        if ex_results:
            return ex_results[:5], []

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

    return [x[1] for x in scored_ex[:7]], []

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
    await update.message.reply_text("اهلا بيك\nابحث ب:\n- الكود: 5506 - 0101\n- الاسم: زجاج - خشب - خيول\nالبوت هيجيب الوصف + تم إصلاح العين + وضع خفيف")

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
    msg = ""
    it = ex_found[0]
    msg += "📦 البند: " + it['code'] + "\n" + it['desc'] + "\n\n"
    if len(ex_found) > 1:
        msg += "🔍 نتائج مشابهة:\n"
        for it2 in ex_found[1:4]:
            msg += "- [" + it2['code'] + "] " + it2['desc'][:50] + "\n"
    if len(msg) > 4000:
        msg = msg[:4000]
    await update.message.reply_text(msg, parse_mode=None)

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
    print("Bot polling started - 256MB Fixed...", flush=True)
    app = Application.builder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CallbackQueryHandler(button_callback))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_search))
    app.run_polling()