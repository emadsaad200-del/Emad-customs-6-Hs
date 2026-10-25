import os, pandas as pd
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, ContextTypes, filters

xls = pd.ExcelFile("EXLL.xlsx")
data=[]
for sheet in xls.sheet_names:
    try:
        df = pd.read_excel(xls, sheet_name=sheet, dtype=str, header=None).fillna("")
        for _, row in df.iterrows():
            txt = " | ".join([str(c).strip() for c in row if str(c).strip()!=""])
            if len(txt) < 20: continue
            # طلع اي رقم شبه بند جمركي
            clean = "".join(filter(str.isdigit, txt))[:10]
            if len(clean) >= 4:
                data.append({"code": txt[:100], "clean": clean, "desc": txt})
    except: pass

print(f"Loaded {len(data)} items - FINAL")
print(f"Sample: {data[0] if data else 'empty'}")

def find(q):
    qd = "".join(filter(str.isdigit, q))
    if not qd: return []
    return [x for x in data if qd in x["clean"] or q in x["desc"]][:5]

TOKEN = os.getenv("BOT_TOKEN") or os.getenv("TELEGRAM_BOT_TOKEN")

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(f"✅ بوت التعريفة الجمركية جاهز\n📦 محمل {len(data)} بند\n\nابعت رقم البند مثلا: 8471 أو كلمة من الوصف")

async def handle_msg(update: Update, context: ContextTypes.DEFAULT_TYPE):
    txt = update.message.text.strip()
    res = find(txt)
    if res:
        for x in res:
            await update.message.reply_text(f"🔍 {x['desc'][:800]}")
    else:
        await update.message.reply_text(f"لم أجد '{txt}'\nجرب رقم بند أو كلمة مفتاحية")

app = ApplicationBuilder().token(TOKEN).build()
app.add_handler(CommandHandler("start", start))
app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_msg))
app.run_polling()
