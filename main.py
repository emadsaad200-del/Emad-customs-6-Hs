import os
import pandas as pd
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

TOKEN = os.getenv("BOT_TOKEN")

try:
    df = pd.read_excel("hs_all.xlsx")
except:
    df = pd.DataFrame()

def search_data(q):
    if df.empty:
        return None
    q = str(q).lower().strip()
    try:
        mask = df.astype(str).apply(lambda row: row.str.lower().str.contains(q).any(), axis=1)
        return df[mask].head(3)
    except:
        return None

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
     await update.message.reply_text(
        "✅ مرحبا بك في Alex Emad Customs 🌟\nاكتب وصف البضاعة او رقم البند"
    )
async def handle_msg(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    res = search_data(text)
    if res is None or res.empty:
        await update.message.reply_text(f"🔍 {text}\nالبوت شغال - هنضيف ملف البيانات بعدين")
        return
    reply = ""
    for _, row in res.iterrows():
        reply += f"{row.iloc[0]} | {str(row.iloc[1])[:100]}\n---\n"
    await update.message.reply_text(reply[:4000])

app = Application.builder().token(TOKEN).build()
app.add_handler(CommandHandler("start", start))
app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_msg))
app.run_polling()
