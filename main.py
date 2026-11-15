import os
import re
import glob
import pandas as pd
import telebot

# ============ الاعدادات ============
TOKEN = os.getenv("BOT_TOKEN", "حط_التوكن_هنا_لو_مش_في_الاعدادات")
bot = telebot.TeleBot(TOKEN)

# ============ قراءة الاكسيل ============
def load_data():
    files = glob.glob("*.xlsx")
    # نستبعد ملف الباك اب
    files = [f for f in files if "backup" not in f]
    FILE = files[0]
    print(f"الملف: {FILE}")

    df_raw = pd.read_excel(FILE, header=None)

    # ندور على صف العناوين
    header_row = None
    for i in range(min(15, len(df_raw))):
        row_text = " ".join([str(x) for x in df_raw.iloc[i].values])
        if "بند" in row_text or "فئ" in row_text or "وارد" in row_text or "HS" in row_text:
            header_row = i
            print(f"لقينا العناوين في الصف: {i}")
            break
    if header_row is None:
        header_row = 3

    df = pd.read_excel(FILE, header=header_row)
    print("الأعمدة:", list(df.columns))

    # تنظيف أسماء الأعمدة
    df.columns = [str(c).strip() for c in df.columns]
    return df

try:
    df = load_data()
    print(f"DB built tariffs: {len(df)}")
except Exception as e:
    print(f"خطأ في تحميل الاكسيل: {e}")
    df = pd.DataFrame()

# ============ توحيد الأعمدة - فئة = وارد ============
def get_col(keys):
    for c in df.columns:
        for k in keys:
            if k in str(c):
                return c
    return None

COL_CODE = get_col(["بند", "HS", "كود", "رمز"])
COL_FEA_WARED = get_col(["فئ", "وارد"])  # فئة ووارد واحد
COL_DESC = get_col(["بيان", "وصف", "سلعة", "اسم", "الصنف"])
COL_UNIT = get_col(["وحدة", "الوحدة"])
COL_GS1 = get_col(["GS1", "باركود", "GTIN"])

print(f"عمود الكود: {COL_CODE}")
print(f"عمود الفئة/الوارد: {COL_FEA_WARED}")
print(f"عمود البيان: {COL_DESC}")

# ============ دالة البحث الرئيسية ============
def search(query):
    query = str(query).strip()
    if df.empty or COL_CODE is None:
        return "قاعدة البيانات مش محملة، جرب تاني بعد شوية"

    # 1- بحث برقم البند: 0101 أو 01012100
    if query.isdigit():
        # ندور على تطابق كامل أو بداية الكود
        df_code = df[COL_CODE].astype(str).str.strip().str.split(".").str[0]
        row = df[df_code == query]
        if row.empty:
            row = df[df_code.str.startswith(query)]
        if not row.empty:
            r = row.iloc[0]
            fea = r[COL_FEA_WARED] if COL_FEA_WARED else "غير موجود"
            desc = r[COL_DESC] if COL_DESC else ""
            unit = r[COL_UNIT] if COL_UNIT else ""
            return f"✅ البند: {r[COL_CODE]}\n📦 البيان: {desc}\n💰 الفئة / الوارد: {fea}%\n📏 الوحدة: {unit}"

    # 2- بحث بالفئة أو الوارد: فئة 5% - وارد 10 - فئة 5 - وارد 5%
    if "فئ" in query or "وارد" in query:
        nums = re.findall(r"\d+", query)
        if nums:
            num = nums[0]
            # ندور على كل البنود اللي فئتها = الرقم ده
            df_fea = df[COL_FEA_WARED].astype(str)
            rows = df[df_fea.str.contains(num, na=False)]
            if not rows.empty:
                res = f"🔍 البنود اللي فئتها / واردها {num}%:\n\n"
                for _, r in rows.head(15).iterrows():
                    desc = r[COL_DESC] if COL_DESC else ""
                    res += f"{r[COL_CODE]} - {desc} - فئة/وارد {r[COL_FEA_WARED]}%\n"
                return res
            else:
                return f"مفيش بنود فئتها {num}%"

    # 3- بحث بالاسم: لحوم - بن - سيارات
    if COL_DESC:
        rows = df[df[COL_DESC].astype(str).str.contains(query, na=False)]
        if not rows.empty:
            res = f"🔍 نتائج البحث عن: {query}\n\n"
            for _, r in rows.head(15).iterrows():
                fea = r[COL_FEA_WARED] if COL_FEA_WARED else ""
                res += f"{r[COL_CODE]} - {r[COL_DESC]} - فئة/وارد {fea}%\n"
            return res

    return "❌ مش لاقي البند ده، جرب كود زي 0101 أو اسم زي بن"

# ============ أوامر البوت ============
@bot.message_handler(commands=['start', 'help'])
def start(message):
    bot.reply_to(message,
        "أهلا بيك 👋\n"
        "ابعت كود البند زي: 0101\n"
        "أو ابعت اسم الصنف زي: بن\n"
        "أو ابعت: فئة 5% أو وارد 10\n"
        "وهجيبلك الفئة / الوارد بتاعه"
    )

@bot.message_handler(func=lambda m: True)
def handle(message):
    try:
        result = search(message.text)
        bot.reply_to(message, result)
    except Exception as e:
        print(f"خطأ: {e}")
        bot.reply_to(message, "حصل خطأ، جرب تاني")

# ============ تشغيل ============
if __name__ == "__main__":
    print("Bot is running...")
    bot.infinity_polling()