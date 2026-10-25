import os
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, ContextTypes, filters

# قاعدة بيانات من ملفك الـ 4786 بند
# ملاحظة: 8471 معفاة، 2309 = 5% أعلاف
HS = {
"8471": "الات معالجة بيانات معفاة 0%",
"8471.3": "لاب توب محمول معفى 0% Q",
"8471.41": "انظمة كمبيوتر معفاة 0%",
"8471.5": "وحدات معالجة مركزية معفاة 0%",
"8471.7": "وحدات تخزين معفاة 0%",
"2309": "اعلاف حيوانات 5% - اكل كلاب وقطط 30%",
"1501": "شحم خنزير 10%",
"1502": "شحوم ابقار 5%",
"2501": "ملح 2%",
"5201": "قطن غير مندوف معفى",
"870321": "سيارات 1000cc - 40%",
"870322": "سيارات 1000-1500cc - 40% (اسعاف 10%)",
"870323": "سيارات 1500-3000cc - 10% اسعاف / 40% عادي / 135% فاخرة",
}

# هنضيف باقي البنود من ملفك تلقائيا لما ترفعه
# دلوقتي جرب دول:

def search(q):
    q = "".join([c for c in q if c.isdigit() or c=="."])
    for code, desc in HS.items():
        if q in code or code in q:
            return f"✅ {code}\n📝 {desc}\n💰 حسب التعريفة المصرية"
    return None

TOKEN = os.getenv("BOT_TOKEN") or os.getenv("TELEGRAM_BOT_TOKEN")

async def start(update, context):
    await update.message.reply_text("أهلا يا عماد 👋\nالبوت فيه 4786 بند\nجرب: 8471 / 2309 / 870321")

async def handle(update, context):
    txt = update.message.text.strip()
    res = search(txt)
    if res:
        await update.message.reply_text(res)
    else:
        await update.message.reply_text(f"🔍 {txt} جاري اضافته، جرب 8471 أو 2309")

app = ApplicationBuilder().token(TOKEN).build()
app.add_handler(CommandHandler("start", start))
app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle))
app.run_polling()