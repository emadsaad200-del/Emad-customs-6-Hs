import os
import json
import pandas as pd
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, CallbackQueryHandler, filters, ContextTypes

TOKEN = os.getenv("BOT_TOKEN") or os.getenv("TELEGRAM_TOKEN")
ADMIN_ID = 1280808746

ALLOWED_FILE = "allowed_users.json"
PENDING_FILE = "pending_users.json"

print("Starting Bot...")

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
    if len(cleaned) <2 or len(cleaned) >8:
        return False
    return True

def clean_desc(text):
    if not text:
        return ""
    t = text.strip()
    if t.startswith('– ') or t.startswith('- '):
        t = t[2:].strip()
    if t.startswith('ر '):
        t = t[2:].strip()
    reps = {
        'رير مشربة': 'غير مشربة',
        'ريرها': 'غيرها',
        'رير': 'غير',
        'زجععاج': 'زجاج',
        'متحصععل': 'متحصل',
        'عععليه': 'عليه',
        'التععويم': 'التعويم',
        'مجلعو': 'مجلو',
        'معصقول': 'مصقول',
        'الوجعه': 'الوجه',
        'مثعل': 'مثل',
        'المعوفرات': 'الموفرات',
        'بطريقعة': 'بطريقة',
        'أخعرى': 'أخرى',
        'حيعععوانععات': 'حيوانات',
        'أجزء': 'أجزاء',
    }
    for old,new in reps.items():
        t = t.replace(old, new)
    # شيل الحروف المكررة من الـ OCR
    res = ""
    prev = ""
    cnt = 0
    for ch in t:
        if ch == prev and '\u0600' <= ch <= '\u06FF':
            cnt += 1
            if cnt >= 2:
                continue
        else:
            cnt = 0
        res += ch
        prev = ch
    return ' '.join(res.split())

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

print("Loading EXLL.xlsx ...")
df = pd.read_excel("EXLL.xlsx", header=None)
items = []
for idx in range(len(df)):
    row = df.iloc[idx]
    c,d = find_code_and_desc(row)
    if c is None or d == "": continue
    if len(d) > 400: continue
    if "جمهورية مصر" in d or "وزارة المالية" in d or "بالقرار الجمهورى" in d: continue
    if c == '2007': continue
    cd = clean_desc(d)
    items.append({'code': c, 'desc': cd, 'clean': c.replace('.','')})

print(f"Loaded {len(items)} valid items - CLEANED")
print(f"Token exists: {TOKEN is not None}")

def is_allowed(uid): return uid in allowed_users or uid == ADMIN_ID

async def request_approval(update, context):
    user = update.effective_user
    if is_allowed(user.id): return True
    if str(user.id) in pending_users:
        await update.message.reply_text("⏳ طلبك قيد المراجعة")
        return False
    pending_users[str(user.id)] = {"name": user.full_name, "username": user.username}
    save_json(PENDING_FILE, pending_users)
    await update.message.reply_text("🔒 البوت خاص. تم ارسال طلبك للادمن.")
    kb = [[InlineKeyboardButton("
