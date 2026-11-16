import os
import re
import pandas as pd
import telebot

# ============ الاعدادات ============
TOKEN = os.getenv("BOT_TOKEN", "حط_التوكن_هنا_لو_مش_في_الاعدادات")
bot = telebot.TeleBot(TOKEN)

# ============ ملفاتك في GitHub ============
TARIFF_FILE = "Tariff.xlsx.csv"
GS1_FILE = "gs1 exle.xlsx"

# ============ تحميل ملف التعريفة ============
def load_tariff():
    print(f"الملف: {TARIFF_FILE}")
    try:
        df = pd.read_csv(TARIFF_FILE, encoding="utf-8-sig", low_memory=False)
    except Exception as e1:
        print(f"utf-8 فشل: {e1} - هنجرب cp1256")
        df = pd.read_csv(TARIFF_FILE, encoding="cp1256", low_memory=False)

    print("الأعمدة:", list(df.columns))
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")
    print(f"DB built tariffs: {len(df)}")
    return df

# ============ تحميل ملف GS1 (اختياري) ============
def load_gs1():
    try:
        g = pd.read_excel(GS1_FILE)
        g.columns = [str(c).strip() for c in g.columns]
        print(f"DB built gs1: {len(g)} - الأعمدة: {list(g.columns)}")
        return g
    except Exception as e:
        print(f"GS1 مش متحمل: {e}")
        return pd.DataFrame()

try:
    df = load_tariff()
except Exception as e:
    print(f"خطأ في تحميل التعريفة: {e}")
    df = pd.DataFrame()

try:
    df_gs1 = load_gs1()
except:
    df_gs1 = pd.DataFrame()

# ============ توحيد الأعمدة - فئة = وارد ============
def get_col(dataframe, keys):
    for c in dataframe.columns:
        for k in keys:
            if k in str(c):
                return c
    return None

COL_CODE = get_col(df, ["بند", "HS", "كود", "رمز"])
COL_FEA_WARED = get_col(df, ["فئ", "وارد"])  # فئة ووارد واحد
COL_DESC = get_col(df, ["بيان", "وصف", "سلعة", "اسم", "الصنف"])
COL_UNIT = get_col(df, ["وحدة", "الوحدة"])

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
        df_code = df[COL_CODE].astype(str).str.strip().str.split(".").str[0]
        # نظف الأصفار والمسافات
        df_code_clean = df_code.str.replace(r"\.0$", "", regex=True).str.strip()
        row = df[df_code_clean == query]
        if row.empty:
            row = df[df_code_clean.str.startswith(query)]
        if not row.empty:
            r = row.iloc[0]
            fea = r[COL_FEA_WARED] if COL_FEA_WARED else "غير موجود"
            desc = r[COL_DESC] if COL_DESC else ""
            unit = r[COL_UNIT] if COL_UNIT else ""
            return f"✅ البند: {r[COL_CODE]}\n📦 البيان: {desc}\n💰 الفئة / الوارد: {fea}%\n📏 الوحدة: {unit}"

    # 2- بحث بالفئة أو الوارد: فئة 5% - وارد 10 - فئة 5 - وارد 5%
    if "فئ" in query or "وارد" in query:
        nums = re.findall(r"\d+", query)
        if nums and COL_FEA_WARED:
            num = nums[0]
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

    # 3- بحث في GS1 لو الرقم طويل (باركود)
    if query.isdigit() and len(query) >= 12 and not df_gs1.empty:
        col_gs1_code = get_col(df_gs1, ["باركود", "GTIN", "GS1", "كود"])
        if col_gs1_code:
            rows = df_gs1[df_gs1[col_gs1_code].astype(str).str.contains(query, na=False)]
            if not rows.empty:
                r = rows.iloc[0]
                return f"📦 نتيجة GS1:\n{r.to_string()}"

    # 4- بحث بالاسم: لحوم - بن - سيارات
    if COL_DESC:
        rows = df[df[COL_DESC].astype(str).str.contains(query, na=False)]
        if not rows.empty:
            res = f"🔍 نتائج البحث عن: {query}\n\n"
            for _, r in rows.head(15).iterrows():
                fea = r[COL_FEA_WARED] if COL_FEA_WARED else ""
                res += f"{r[COL_CODE]} - {r[COL_DESC]} - فئة/وارد {fea}%\n"
            return res

    return "❌ مش لاقي البند ده، جرب كود زي 0101 أو اسم زي بن أو فئة 5%"

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