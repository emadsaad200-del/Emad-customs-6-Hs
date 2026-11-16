import os
import re
import glob
import pandas as pd
import telebot

TOKEN = os.getenv("BOT_TOKEN", "حط_التوكن_هنا")
bot = telebot.TeleBot(TOKEN)

SITE_URL = "https://www.customs.gov.eg"

def load_file(path):
    if path.endswith(".csv"):
        for enc in ["cp1256", "utf-8-sig", "windows-1256", "latin1"]:
            for sep in [",", ";", "\t"]:
                try:
                    df_try = pd.read_csv(path, encoding=enc, sep=sep, engine="python", on_bad_lines="skip", low_memory=False)
                    if len(df_try.columns) >= 3:
                        return df_try
                except Exception:
                    continue
        try:
            return pd.read_excel(path, engine="openpyxl")
        except Exception:
            return None
    else:
        try:
            return pd.read_excel(path)
        except Exception:
            return None

def load_tariff():
    files = glob.glob("*.csv") + glob.glob("*.xlsx")
    files_sorted = sorted(files, key=lambda x: 0 if "tariff" in x.lower() else 1)
    for f in files_sorted:
        if "backup" in f.lower() or "gs1" in f.lower():
            continue
        df_try = load_file(f)
        if df_try is None:
            continue
        df_try.columns = [str(c).strip() for c in df_try.columns]
        cols = " ".join(df_try.columns)
        if "بند" in cols or "فئ" in cols or "وارد" in cols or "HS" in cols or "رمز" in cols:
            return df_try.dropna(how="all")
    return pd.DataFrame()

def load_gs1():
    try:
        g = pd.read_excel("gs1 exle.xlsx")
        g.columns = [str(c).strip() for c in g.columns]
        return g
    except Exception as e:
        print(e)
        return pd.DataFrame()

df = load_tariff()
df_gs1 = load_gs1()

def get_col(dataframe, keys):
    for c in dataframe.columns:
        for k in keys:
            if k.lower() in str(c).lower():
                return c
    return None

COL_CODE = get_col(df, ["بند", "HS", "كود", "رمز"])
COL_FEA = get_col(df, ["فئ", "وارد"])
COL_DESC_AR = get_col(df, ["بيان", "وصف", "الصنف", "سلعة", "عربي"])
COL_DESC_EN = get_col(df, ["english", "انجليزي", "description", "item", "name"])
COL_UNIT = get_col(df, ["وحدة"])
COL_VAT = get_col(df, ["قيمة مضافة", "مضافة", "VAT", "ضريبة"])
COL_PROFIT = get_col(df, ["ارباح", "أرباح", "تجاري", "صناعي", "profit"])

if COL_DESC_AR is None and COL_DESC_EN:
    COL_DESC_AR = COL_DESC_EN
if COL_DESC_EN is None and COL_DESC_AR:
    COL_DESC_EN = COL_DESC_AR

def normalize_code(x):
    s = str(x).strip().split(".")[0].strip().replace(" ", "").replace("-", "")
    if s.isdigit() and len(s) <= 4:
        s = s.zfill(4)
    return s

def format_item(r):
    if COL_FEA:
        fea = r[COL_FEA]
    else:
        fea = "غير موجود"
    if COL_DESC_AR:
        desc_ar = r[COL_DESC_AR]
    else:
        desc_ar = ""
    desc_en = ""
    if COL_DESC_EN and COL_DESC_EN != COL_DESC_AR:
        desc_en = r[COL_DESC_EN]
    if COL_UNIT:
        unit = r[COL_UNIT]
    else:
        unit = ""
    if COL_VAT:
        vat = r[COL_VAT]
    else:
        vat = "14"
    if COL_PROFIT:
        profit = r[COL_PROFIT]
    else:
        profit = "غير موجود"
    txt = "✅ البند: " + str(r[COL_CODE]) + "\n"
    txt += "📦 البيان عربي: " + str(desc_ar) + "\n"
    if desc_en and str(desc_en) != "nan":
        txt += "📦 البيان انجليزي: " + str(desc_en) + "\n"
    txt += "💰 الفئة / الوارد: " + str(fea) + "%\n"
    txt += "🧾 القيمة المضافة: " + str(vat) + "%\n"
    txt += "🏭 الأرباح التجارية الصناعية: " + str(profit) + "%\n"
    txt += "📏 الوحدة: " + str(unit)
    return txt

def calc_cif(row, amount):
    try:
        fea = float(str(row[COL_FEA]).replace("%", "").strip()) if COL_FEA else 0
    except Exception:
        fea = 0
    try:
        vat = float(str(row[COL_VAT]).replace("%", "").strip()) if COL_VAT else 14
    except Exception:
        vat = 14
    try:
        profit = float(str(row[COL_PROFIT]).replace("%", "").strip()) if COL_PROFIT else 0
    except Exception:
        profit = 0
    cif = amount
    customs = cif * fea / 100
    vat_val = (cif + customs) * vat / 100
    profit_val = cif * profit / 100
    total = cif + customs + vat_val + profit_val
    txt = format_item(row)
    txt += "\n\n💵 مبلغ الفاتورة CIF: " + format(cif, ",.2f") + "\n"
    txt += "🛃 رسم الوارد (" + str(fea) + "%): " + format(customs, ",.2f") + "\n"
    txt += "🧾 قيمة مضافة (" + str(vat) + "%): " + format(vat_val, ",.2f") + "\n"
    txt += "🏭 أرباح تجارية صناعية (" + str(profit) + "%): " + format(profit_val, ",.2f") + "\n"
    txt += "💰 الإجمالي: " + format(total, ",.2f")
    return txt

def search_gs1(keyword):
    if df_gs1.empty:
        return "ملف GS1 مش متحمل"
    keyword = keyword.strip().lower()
    mask = pd.Series([False]*len(df_gs1))
    for c in df_gs1.columns:
        mask = mask | df_gs1[c].astype(str).str.lower().str.contains(keyword, na=False)
    rows = df_gs1[mask]
    if rows.empty:
        return "❌ مفيش نتيجة في GS1 لـ: " + keyword
    res = "🔍 نتائج GS1 لـ: " + keyword + "\n\n"
    for _, r in rows.head(10).iterrows():
        res += "---\n"
        for c in df_gs1.columns:
            v = r[c]
            if str(v) != "nan" and str(v).strip() != "":
                res += str(c) + ": " + str(v) + "\n"
    return res

def search(query):
    query = str(query).strip()
    q_low = query.lower()
    if q_low.startswith("gpc") or q_low.startswith("gs1"):
        keyword = re.sub(r"^(gpc|gs1)", "", query, flags=re.IGNORECASE).strip()
        if not keyword:
            return "ابعت كده: gpc milk"
        return search_gs1(keyword)
    parts = query.split()
    if len(parts) == 2 and parts[0].replace(".", "").isdigit() and parts[1].replace(",", "").replace(".", "").isdigit():
        code_norm = normalize_code(parts[0])
        amount = float(parts[1].replace(",", ""))
        df_codes = df[COL_CODE].astype(str).apply(normalize_code)
        row = df[df_codes == code_norm]
        if row.empty:
            row = df[df_codes.str.startswith(code_norm)]
        if not row.empty:
            return calc_cif(row.iloc[0], amount)
        else:
            return "❌ البند " + parts[0] + " مش موجود"
    if query.replace(" ", "").isdigit() and len(query.replace(" ", "")) <= 10:
        q_norm = normalize_code(query)
        df_codes = df[COL_CODE].astype(str).apply(normalize_code)
        row = df[df_codes == q_norm]
        if row.empty:
            row = df[df_codes.str.startswith(q_norm)]
        if not row.empty:
            return format_item(row.iloc[0])
    if query.isdigit() and len(query) >= 12 and not df_gs1.empty:
        col_gs1_code = get_col(df_gs1, ["باركود", "GTIN", "GS1", "كود"])
        if col_gs1_code:
            rows = df_gs1[df_gs1[col_gs1_code].astype(str).str.contains(query, na=False)]
            if not rows.empty:
                r = rows.iloc[0]
                res = "📦 نتيجة GS1:\n"
                for c in df_gs1.columns:
                    res += str(c) + ": " + str(r[c]) + "\n"
                return res
    if COL_DESC_AR or COL_DESC_EN:
        mask = pd.Series([False]*len(df))
        if COL_DESC_AR:
            mask = mask | df[COL_DESC_AR].astype(str).str.contains(query, case=False, na=False)
        if COL_DESC_EN and COL_DESC_EN != COL_DESC_AR:
            mask = mask | df[COL_DESC_EN].astype(str).str.contains(query, case=False, na=False)
        rows = df[mask]
        if not rows.empty:
            res = "🔍 نتائج البحث عن: " + query + "\n\n"
            for _, r in rows.head(15).iterrows():
                if COL_FEA:
                    fea_v = r[COL_FEA]
                else:
                    fea_v = ""
                if COL_DESC_AR:
                    desc_v = r[COL_DESC_AR]
                else:
                    desc_v = ""
                if COL_VAT:
                    vat_v = r[COL_VAT]
                else:
                    vat_v = "14"
                if COL_PROFIT:
                    profit_v = r[COL_PROFIT]
                else:
                    profit_v = ""
                res += str(r[COL_CODE]) + " - " + str(desc_v) + " - فئة/وارد " + str(fea_v) + "% - VAT " + str(vat_v) + "% - أرباح " + str(profit_v) + "%\n"
            return res
    return "❌ مش لاقي البند ده\nجرب: 0101 - بن - coffee - gpc milk - 0101 10000"

@bot.message_handler(commands=["start", "help"])
def start(message):
    bot.reply_to(message, "أهلا بيك 👋\nابعت كود البند: 0101\nابعت اسم عربي أو انجليزي: بن - coffee\nابعت GS1: gpc milk\nابعت بند + مبلغ: 0101 10000")

@bot.message_handler(func=lambda m: True)
def handle(message):
    try:
        bot.reply_to(message, search(message.text))
    except Exception as e:
        print(e)
        bot.reply_to(message, "حصل خطأ، جرب تاني")

if __name__ == "__main__":
    print("Bot is running...")
    bot.infinity_polling()