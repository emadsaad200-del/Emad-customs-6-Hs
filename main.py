import os
import pandas as pd
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, ContextTypes, filters

print("Loading EXLL.xlsx...")
xls = pd.ExcelFile("EXLL.xlsx")
print(f"Sheets found: {xls.sheet_names}")

data = []
for sheet in xls.sheet_names:
    print(f"--- Checking sheet: {sheet} ---")
    try:
        df = pd.read_excel(xls, sheet_name=sheet, header=None)
        print(f"Shape: {df.shape}")
        print(df.head(3).to_string())
        
        for idx, row in df.iterrows():
            row_vals = [str(v) for v in row.values if str(v) != 'nan']
            if len(row_vals) < 2:
                continue
            
            # دور على اي حاجة شكلها كود جمركي
            for val in row_vals:
                clean = "".join([c for c in val if c.isdigit()])
                # لو فيه 4 ارقام على الاقل يبقى ده بند
                if len(clean) >= 4:
                    desc = " | ".join(row_vals)
                    if len(desc) > 10:
                        data.append({"code": val.strip(), "clean": clean, "desc": desc[:500], "duty": "شوف الوصف"})
                        break
    except Exception as e:
        print(f"Error in sheet {sheet}: {e}")

# لو الطريقة اللي فوق فشلت، جرب نقرأ الملف كله ك نص
if len(data) == 0:
    print("Trying second method...")
    for sheet in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=sheet, header=None, dtype=str)
            for _, row in df.iterrows():
                for cell in row:
                    s = str(cell)
                    if len(s) > 15: # وصف طويل
                        data.append({"code": s[:30], "clean": "".join(filter(str.isdigit, s)), "desc": s, "duty": ""})
        except: pass

print(f"Loaded {len(data)} items - FINAL")
# اطبع اول بند عشان نتأكد
if data:
    print(f"Sample: {data[0]}")

def find(q):
    qd = "".join([c for c in q if c.isdigit()])
    if not qd:
        return []
    r=[]
    for it in data:
        if qd in it["clean"] or qd in it["code"]:
            r.append(it)
    return r[:5]

TOKEN = os.getenv("BOT_TOKEN") or os.getenv("TELEGRAM_BOT_TOKEN")

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(f"✅ البوت شغال! محمل {len(data)} بند من ملفك\nجرب ابعت 8471")

async def handle_msg(update: Update, context: ContextTypes.DEFAULT_TYPE):
    txt = update.message.text.strip()
    print(f"User search: {txt}")
    res = find(txt)
    if res:
        msg=""
        for x in res:
            msg+=f"✅ {x['code']}\n{x['desc'][:300]}\n\n"
        await update.message.reply_text(msg[:4000])
    else:
        await update.message.reply_text(f"مش لاقي '{txt}' - جرب رقم تاني. عندي {len(data)} بند")

app = ApplicationBuilder().token(TOKEN).build()
app.add_handler(CommandHandler("start", start))
app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_msg))
print("Bot polling...")
app.run_polling()
