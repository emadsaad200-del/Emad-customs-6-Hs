import os, pandas as pd, re
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, ContextTypes, filters

# تحميل الملف
xls = pd.ExcelFile("EXLL.xlsx")
data=[]
for sheet in xls.sheet_names:
    try:
        df = pd.read_excel(xls, sheet_name=sheet, dtype=str, header=None).fillna("")
        for _, row in df.iterrows():
            cells = [str(c).strip() for c in row if str(c).strip() not in ["", "nan", "None"]]
            if len(cells) < 2: continue
            txt = " | ".join(cells)
            if len(txt) < 10: continue
            if "قرار جمهوري" in txt and len(txt) < 60: continue
            
            code = cells[0]
            # دور على كود شبه بند
            for c in cells:
                if re.search(r'\d{2,4}[\.\-]?\d{2,4}', c):
                    code = c
                    break
            
            clean = "".join(filter(str.isdigit, code))
            if len(clean) >= 4:
                data.append({
                    "code": code.strip(),
                    "desc": txt,
                    "clean": clean,
                    "search": txt.lower() + " " + clean
                })
    except Exception as e:
        print(f"err {sheet}: {e}")

print(f"Loaded {len(data)} items - FINAL")
if data: print(data[0])

def find(query):
    query = query.strip().lower()
    qd = "".join(filter(str.isdigit, query))
    q_words = query.split()
    
    scored = []
    for x in data:
        score = 0
        # لو رقم البند موجود
        if qd and qd in x["clean"]:
            score += 100
            if x["clean"] == qd: score += 50
        # لو كلمة عربية موجودة
        for w in q_words:
            if len(w) >= 2 and w in x["search"]:
                score += 20
        
        if score > 0:
            scored.append((score, x))
    
    scored.sort(key=lambda a: a[0], reverse=True)
    return [x for s,x in scored[:5]]

TOKEN = os.getenv("BOT_TOKEN") or os.getenv("TELEGRAM_BOT_TOKEN")

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        f"✅ أهلا بيك - بوت Ale Emad Customs\n"
        f"📦 محمل {len(data)} بند جمركي من ملفك\n\n"
        f"🔍 طرق البحث:\n"
        f"1️⃣ برقم البند: 2309 - 8471 - 0101\n"
        f"2️⃣ بكلمة عربية: حيوانات - سيارات - ريره - أدوية\n"
        f"3️⃣ بجزء من البند: 1702.30\n\n"
        f"جرب دلوقتي 👇",
    )

async def handle_msg(update: Update, context: ContextTypes.DEFAULT_TYPE):
    txt = update.message.text.strip()
    print(f"Search: {txt}")
    
    if txt.lower() in ["start", "/start", "start /", "البداية", "بداية"]:
        return await start(update, context)
    
    res = find(txt)
    if not res:
        await update.message.reply_text(f"❌ لم أجد '{txt}'\nجرب كلمة تانية: حيوانات، سيارات، أدوية، أو رقم بند")
        return
    
    for x in res:
        # قص الوصف بشكل حلو
        desc = x['desc']
        if len(desc) > 700:
            desc = desc[:700] + "..."
        
        msg = f"📋 البند: {x['code']}\n\n📝 {desc}"
        await update.message.reply_text(msg)

app = ApplicationBuilder().token(TOKEN).build()
app.add_handler(CommandHandler("start", start))
app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_msg))
print("Bot polling started...")
app.run_polling()
