import os, pandas as pd
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, ContextTypes, filters

print("Starting Bot...")

# تحميل الملف - يحاول كل الأماكن
file_path = "EXLL.xlsx"
if not os.path.exists(file_path):
    file_path = "./EXLL.xlsx"

print(f"Loading {file_path} ...")
try:
    xls = pd.ExcelFile(file_path)
    df = pd.read_excel(xls, sheet_name=0, dtype=str, header=None).fillna("")
    print(f"File loaded: {len(df)} rows")
except Exception as e:
    print(f"ERROR loading file: {e}")
    df = pd.DataFrame()

# قاموس تصليح العربي
FIX = {
    "منلجات": "منتجات",
    "المسم": "القسم",
    "اليصل": "الفصل",
    "الممسم": "القسم",
    "رير ": "غير ",
    "ولاداخلة": "ولا داخلة",
    "نفايا": "نفايات",
    "الأرذية": "الأغذية",
    "أرذية": "أغذية",
    "رذائية": "غذائية",
    "الممشش": "التبن",
    "القلفعونية": "القلفونية",
    "اللعرييمة": "التعريفة",
    "الجمركيمة": "الجمركية",
    "اليئة": "الفئة",
}

def fix_text(t):
    for k,v in FIX.items():
        t = t.replace(k,v)
    # شيل حرف مقصوص في الاول زي "ر "
    if len(t)>2 and t[1]==' ' and t[0] in 'رع':
        t = t[2:]
    return t.strip()

data=[]
for i in range(len(df)):
    row = [str(c).strip() for c in df.iloc[i].tolist() if str(c).strip() and str(c).strip()!='nan']
    if not row: continue
    text = " ".join(row)
    if len(text)<15: continue
    if "القرار الجمهورى" in text and len(text)<100: continue
    if "جمهورية مصر" in text and len(text)<60: continue
    
    # دور على كود
    code=None
    code_idx=-1
    for idx,c in enumerate(row):
        cc=c.strip(".,;:()[] ")
        if "." in cc:
            parts=cc.split(".")
            if len(parts)==2 and parts[0].isdigit() and parts[1].isdigit():
                if 1<=len(parts[0])<=4 and 1<=len(parts[1])<=4:
                    code=cc
                    code_idx=idx
                    break
    
    if not code: continue
    
    # الوصف
    best=""
    best_score=0
    for idx,c in enumerate(row):
        if idx==code_idx: continue
        if len(c)<15: continue
        ar=sum(1 for ch in c if '\u0600' <= ch <= '\u06FF')
        if ar>best_score and ar>8:
            best_score=ar
            best=c
    
    if not best: continue
    best=fix_text(best)
    
    data.append({
        "code": code,
        "desc": best,
        "clean": "".join([ch for ch in code if ch.isdigit()]),
        "search": fix_text(text).lower()
    })

# مكرر
uniq={}
for d in data:
    k=d["code"]+"|"+d["desc"][:40]
    uniq[k]=d
data=list(uniq.values())

print(f"Loaded {len(data)} items - FINAL FIXED")
for d in data[:2]:
    print(d)

def find(q):
    # يفهم 73-02-10-00-00
    q = q.replace("-", ".").replace("–", ".").replace("_", ".")
    q = fix_text(q)
    qd = "".join([ch for ch in q if ch.isdigit()])
    words = [w for w in q.lower().split() if len(w)>=2]
    scored=[]
    for x in data:
        s=0
        if qd and qd in x["clean"]:
            s+=100
            if x["clean"]==qd: s+=150
            if x["clean"].startswith(qd[:4]): s+=50
        for w in words:
            if w in x["search"]:
                s+=30
        if s>0:
            scored.append((s,x))
    scored.sort(key=lambda a:a[0], reverse=True)
    return [x for _,x in scored[:5]]

TOKEN = os.getenv("BOT_TOKEN") or os.getenv("TELEGRAM_BOT_TOKEN") or os.getenv("TELEGRAM_TOKEN")
print(f"Token exists: {bool(TOKEN)}")

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(f"✅ Ale Emad Customs شغال\n📦 {len(data)} بند\n\nجرب:\n73-02-10-00-00\n1.01\nخيول\nقضبان")

async def handle_msg(update: Update, context: ContextTypes.DEFAULT_TYPE):
    txt = update.message.text.strip()
    if "/start" in txt.lower() or txt.lower()=="start":
        return await start(update, context)
    res = find(txt)
    if not res:
        await update.message.reply_text(f"❌ لم أجد '{txt}'\nجرب كود زي 730210 أو 1.01")
        return
    for x in res:
        await update.message.reply_text(f"📋 البند: {x['code']}\n\n📝 {x['desc']}\n\n🔢 الكود النظيف: {x['clean']}")

if not TOKEN:
    print("ERROR: No BOT_TOKEN in env!")
else:
    app = ApplicationBuilder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_msg))
    print("Bot polling started...")
    app.run_polling()
