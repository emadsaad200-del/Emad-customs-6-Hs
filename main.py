import os
import re
import glob
import pandas as pd
import telebot

# ============ الاعدادات ============
TOKEN = os.getenv("BOT_TOKEN", "حط_التوكن_هنا_لو_مش_في_الاعدادات")
bot = telebot.TeleBot(TOKEN)

# ============ تحميل أي ملف - جديد - بيحل مشكلة سطر 233 البايظ ============
def load_file(path):
    try:
        if path.endswith(".csv"):
            # جرب كل الترميزات والفواصل لحد ما واحدة تنجح
            for enc in ["cp1256", "utf-8-sig", "windows-1256"]:
                for sep in [",", ";"]:
                    try:
                        df_try = pd.read_csv(
                            path,
                            encoding=enc,
                            sep=sep,
                            engine="python",
                            on_bad_lines="skip",
                            low_memory=False
                        )
                        # لو قرأ 3 أعمدة على الأقل يبقى الـ sep صح
                        if len(df_try.columns) >= 3:
                            print(f"اتقرأ {path} بـ {enc} و فاصل {sep} - الأعمدة: {len(df_try.columns)}")
                            return df_try
                    except:
                        continue
            print(f"فشل قراءة CSV: {path}")
            return None
        else:
            return pd.read_excel(path)
    except Exception as e:
        print(f"فشل {path}: {e}")
        return None

# ============ ندور على ملف التعريفة الصح لوحده - جديد ============
def load_tariff():
    files = glob.glob("*.csv") + glob.glob("*.xlsx")
    print(f"كل الملفات: {files}")
    # جرب Tariff الأول
    files_sorted = sorted(files, key=lambda x: 0 if "tariff" in x.lower() else 1)
    for f in files_sorted:
        if "backup" in f.lower():
            continue
        df_try = load_file(f)
        if df_try is None:
            continue
        df_try.columns = [str(c).strip() for c in df_try.columns]
        cols = " ".join(df_try.columns)
        print(f"بجرب {f} -> {list(df_try.columns)[:6]}")
        # الملف الصح لازم يكون فيه بند أو فئة أو وارد
        if "بند" in cols or "فئ" in cols or "وارد" in cols or "HS" in cols:
            print(f"✅ ملف التعريفة الصح هو: {f}")
            return df_try.dropna(how="all")
    print("❌ مفيش ملف فيه بند/فئة/وارد - كل الملفات GS1")
    return pd.DataFrame()

# ============ تحميل ملف GS1 - قديم ============
def load_gs1():
    try:
        g = pd.read_excel("gs1 exle.xlsx")
        g.columns = [str(c).strip() for c in g.columns]
        print(f"DB built gs1: {len(g)}")
        return g
    except Exception as e:
        print(f"GS1 مش متحمل: {e}")
        return pd.DataFrame()

# ============ تحميل البيانات ============
df = load_tariff()
df_gs1 = load_gs1()

print(f"DB built tariffs: {len(df)}")
print("الأعمدة:", list(df.columns))

# ============ توحيد الأعمدة - فئة = وارد - قديم + جديد ============
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

# ============ دالة البحث الرئيسية - قديم + جديد ============
def search(query):
    query = str(query).strip()
    if df.empty or COL_CODE is None:
        return "قاعدة البيانات مش محملة صح - ملف التعريفة فيه مشكلة، اتأكد ان Tariff.xlsx.csv مرفوع صح"

    # 1- بحث برقم البند: 0101 أو 01012100 - قديم
    if query.isdigit():
        df_code = df[COL_CODE].astype(str).str.strip().str.split(".").str[0]
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

    # 2- بحث بالفئة أو الوارد: فئة 5% - وارد 10 - قديم
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

    # 3- بحث في GS1 لو الرقم طويل (باركود) - قديم
    if query.isdigit() and len(query) >= 12 and not df_gs1.empty:
        col_gs1_code = get_col(df_gs1, ["باركود", "GTIN", "GS1", "كود"])
        if col_gs1_code:
            rows = df_gs1[df_gs1[col_gs1_code].astype(str).str.contains(query, na=False)]
            if not rows.empty:
                r = rows.iloc[0]
                return f"📦 نتيجة GS1:\n{r.to_string()}"

    # 4- بحث بالاسم: لحوم - بن - سيارات - قديم
    if COL_DESC:
        rows = df[df[COL_DESC].astype(str).str.contains(query, na=False)]
        if not rows.empty:
            res = f"🔍 نتائج البحث عن: {query}\n\n"
            for _, r in rows.head(15).iterrows():
                fea = r[COL_FEA_WARED] if COL_FEA_WARED else ""
                res += f"{r[COL_CODE]} - {r[COL_DESC]} - فئة/وارد {fea}%\n"
            return res

    return "❌ مش لاقي البند ده، جرب كود زي 0101 أو اسم زي بن أو فئة 5%"

# ============ أوامر البوت - قديم ============
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

# ============ تشغيل - قديم ============
if __name__ == "__main__":
    print("Bot is running...")
    bot.infinity_polling()