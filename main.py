import os, pandas as pd
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, ContextTypes, filters

def is_code(s):
    s = s.strip()
    if "." in s:
        parts = s.split(".")
        if len(parts)==2 and parts[0].isdigit() and parts[1].isdigit():
            if 1 <= len(parts[0]) <=4 and 1 <= len(parts[1]) <=4:
                return True
    return False

def extract_digits(s):
    return "".join([c for c in s if c.isdigit()])

print("Loading EXLL.xlsx ...")
xls = pd.ExcelFile("EXLL.xlsx")
df = pd.read_excel(xls, sheet_name='Table 1', dtype=str, header=None).fillna("")

data = []
for idx in range(len(df)):
    row = df.iloc[idx].tolist()
    cells = [str(c).strip() for c in row if str(c).strip() and str(c).strip()!='nan']
    if not cells: continue
    
    for cell in cells:
        tokens = cell.replace("\n"," ").split()
        for tok in tokens:
            tok = tok.strip().strip(".,;:")
            if is_code(tok):
                # هات الوصف العربي الأطول
                arabic_cands = []
                for c in cells:
                    if c==cell: continue
                    if len(c)<15: continue
                    if "القرار" in c or "المادة" in c or "رئيس الجمهورية" in c: continue
                    cnt = sum(1 for ch in c if '\u0600' <= ch <= '\u06FF')
                    if cnt>5:
                        arabic_cands.append(c)
                if arabic_cands:
                    desc = max(arabic_cands, key=len)
                    data.append({
                        "code": tok,
                        "desc": desc,
                        "clean": extract_digits(tok),
                        "search": desc.lower()
                    })
                break
        if is_code(cell):
            arabic_cands = [c for c in cells if c!=cell and len(c)>15 and sum(1 for ch in c if '\u0600' <= ch <= '\u06FF')>5]
            if arabic_cands and "القرار" not in max(arabic_cands, key=len):
                data.append({
                    "code": cell,
                    "desc": max(arabic_cands, key=len),
                    "clean": extract_digits(cell),
                    "search": max(arabic_cands, key=len).lower()
                })

# شيل المكرر
uniq = {}
for d in data:
    key = d["code"]+"|"+d["desc"][:60]
    uniq[key]=d
data = list(uniq.values())

print(f"Loaded {len(data)} items - FINAL CLEAN")
for d in data[:3]:
    print(d)

def find(q):
    q = q.strip().lower()
    qd = extract_digits(q)
    words = [w for w in q.split() if len(w)>=2]
    scored=[]
    for x in data:
        s=0
        if qd and qd in x["clean"]:
            s+=100
            if x["clean"]==qd: s+=100
        for w in words:
            if w in x["search"]:
                s+=40
        if s>0:
            scored.append((s,x))
    scored.sort(key=lambda a:a[0], reverse=True)
    return [x for _,x in scored[:5]]

TOKEN = os.getenv("BOT_TOKEN") or os.getenv("TELEGRAM_BOT_TOKEN")

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        f"✅ Ale Emad Customs\n📦 {len(data)} بند جمركي نضيف\n\n🔍 جرب:\n• 23.09\n• 1.01\n• خيول\n• تغذية الحيوانات\n• كلاب"
    )

async def handle_msg(update: Update, context: ContextTypes.DEFAULT_TYPE):
    txt = update.message.text.strip()
    if "start" in txt.lower(): return await start(update, context)
    res=find(txt)
    if not res:
        await update.message.reply_text(f"❌ لم أجد '{txt}'\nجرب: 23.09 أو تغذية أو خيول")
        return
    for x in res:
        await update.message.reply_text(f"📋 البند: {x['code']}\n\n📝 {x['desc']}\n\n🔢 الكود: {x['clean']}")

app = ApplicationBuilder().token(TOKEN).build()
app.add_handler(CommandHandler("start", start))
app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_msg))
print("Bot polling started...")
app.run_polling()
