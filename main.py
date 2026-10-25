import os
import pandas as pd
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, ContextTypes, filters

print("Loading EXLL.xlsx...")
xls = pd.ExcelFile("EXLL.xlsx")
data = []
for sheet in xls.sheet_names:
    try:
        df = xls.parse(sheet, header=None)
        for _, row in df.iterrows():
            if len(row.values) < 3: continue
            duty = str(row.values[0]) if pd.notna(row.values[0]) else ""
            desc = str(row.values[1]) if pd.notna(row.values[1]) else ""
            code = str(row.values[2]) if pd.notna(row.values[2]) else ""
            digits = "".join([c for c in code if c.isdigit()])
            if len(digits)>=2 and len(digits)<=10 and len(desc)>5 and len(code)<30:
                if len(duty)<25:
                    clean = "".join([c for c in code if c.isdigit() or c=="."])
                    data.append({"code": code.strip(), "clean": clean, "duty": duty.strip(), "desc": desc.strip()})
    except: pass

print(f"Loaded {len(data)} items")

def find(q):
    q_clean = "".join([c for c in q if c.isdigit() or c=="."])
    q_digits = "".join([c for c in q_clean if c.isdigit()])
    results = []
    for item in data:
        ic = item["clean"].replace(".","")
        if q_digits and (q_digits in ic or ic.startswith(q_digits)):
            results.append(item)
    # بحث بالوصف
    if not results:
        for item in data:
            if q in item["desc"]:
                results.append(item)
    return results[:5]

TOKEN = os.getenv("BOT_TOKEN") or os.getenv("TELEGRAM_BOT_TOKEN")

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(f"أهلا يا عماد 👋\nمحمل {len(data)} بند من EXLL.xlsx\nابعت: 8471 أو 2309 أو 870321")

async def handle_msg(update: Update, context: ContextTypes.DEFAULT_TYPE):
    txt = update.message.text.strip()
    res = find(txt)
    if res:
        msg = ""
        for r in res[:3]:
            duty = r["duty"]
            if "معف" in duty or "معي" in duty or "معا" in duty: duty = "معفاة 0% Q"
            if not duty or duty.lower()=="nan": duty = "راجع"
            msg += f"✅ {r['code']}\n📝 {r['desc'][:200]}\n💰 {duty}\n\n"
        await update.message.reply_text(msg)
    else:
        await update.message.reply_text(f"🔍 {txt} مش موجود، جرب 8471")

app = ApplicationBuilder().token(TOKEN).build()
app.add_handler(CommandHandler("start", start))
app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_msg))
app.run_polling()
