import os
import re
import pandas as pd
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

TOKEN = os.getenv("BOT_TOKEN")

def find_file(keywords):
    files = os.listdir(".")
    for f in files:
        low = f.lower().replace(" ", "").replace("_", "")
        if all(k.lower().replace(" ", "") in low for k in keywords):
            return f
    return None

def load_table_file(path):
    if not path or not os.path.exists(path):
        return pd.DataFrame()
    try:
        if path.lower().endswith(".csv") or ".csv" in path.lower():
            try:
                return pd.read_csv(path, dtype=str, low_memory=False, encoding="utf-8-sig")
            except:
                return pd.read_csv(path, dtype=str, low_memory=False, encoding="cp1256", errors="ignore")
        return pd.read_excel(path, dtype=str, engine="openpyxl")
    except:
        try:
            return pd.read_csv(path, dtype=str, low_memory=False, encoding="utf-8-sig")
        except:
            return pd.read_csv(path, dtype=str, low_memory=False, encoding="cp1256", errors="ignore")

def norm_col(c):
    return str(c).strip().replace(" ", "").replace("_", "").replace("\n", "")

def find_header_row(df_raw):
    for i in range(min(12, len(df_raw))):
        row = " ".join([str(x) for x in df_raw.iloc[i].tolist()]).lower()
        if ("bnd" in row or "code" in row or "hs" in row):
            return i
        if ("wared" in row or "wasf" in row):
            return i
    return 0

def clean_df(df):
    if df.empty:
        return df
    try:
        h = find_header_row(df)
        if h > 0:
            df.columns = df.iloc[h].astype(str)
            df = df.iloc[h+1:].reset_index(drop=True)
    except:
        pass
    df.columns = [str(c).strip() for c in df.columns]
    df = df.astype(str).fillna("")
    return df

BOOK_FILE = find_file(["book1"]) or "Book1LOGIST TATTEF.xlsx"
TARIFF_FILE = find_file(["tariff"]) or "Tariff.xlsx.csv"
GS1_FILE = find_file(["gs1"]) or "gs1 exle.xlsx"

df_book = clean_df(load_table_file(BOOK_FILE))
df_tariff = clean_df(load_table_file(TARIFF_FILE))
df_gs1 = clean_df(load_table_file(GS1_FILE))
print(f"DB built book: {len(df_book)} tariffs: {len(df_tariff)} gs1: {len(df_gs1)}")

def get_col(df, names):
    if df.empty:
        return None
    cols = {norm_col(c): c for c in df.columns}
    for n in names:
        nn = norm_col(n)
        if nn in cols:
            return cols[nn]
        for k, v in cols.items():
            if nn in k:
                return v
    return None

def clean_code(s):
    return re.sub(r"\D", "", str(s))

def clean_val(v):
    if v is None:
        return None
    s = str(v).strip().lower()
    s = s.replace("nan", "").replace("none", "").replace("%", "").strip()
    if s == "" or s == "-":
        return None
    s = s.translate(str.maketrans("0123456789", "0123456789"))
    m = re.search(r"\d+(\.\d+)?", s)
    if not m:
        return None
    try:
        return float(m.group(0))
    except:
        return None

def extract_rates(row, df):
    cat = None
    wared = None
    vat = None
    arbah = None
    for col in df.columns:
        c = str(col).lower()
        val = clean_val(row[col])
        if val is None:
            continue
        if "vat" in c:
            vat = val
        elif "category" in c:
            cat = val
        elif "arbah" in c or "profit" in c:
            arbah = val
        elif "wared" in c or "duty" in c or "custom" in c:
            if wared is None:
                wared = val
        # فحص بالعربي بسطور قصيرة
        if "vat" not in c and "category" not in c:
            if "مضاف" in c:
                vat = val
            if "ارباح" in c:
                arbah = val
            if "وارد" in c:
                if wared is None:
                    wared = val
            if "فئة" in c:
                if cat is None:
                    cat = val
    return cat, wared, vat, arbah

def search_in_df(df, q):
    if df.empty:
        return pd.DataFrame()
    q_clean = clean_code(q)
    q_low = str(q).lower().strip()
    code_col = get_col(df, ["code", "hs"])
    desc_col = get_col(df, ["description", "name"])
    mask = pd.Series([False]*len(df))
    if code_col and q_clean:
        mask = mask | df[code_col].astype(str).apply(lambda x: q_clean in clean_code(x))
    if desc_col and q_low:
        mask = mask | df[desc_col].astype(str).str.lower().str.contains(re.escape(q_low), na=False)
    return df[mask].head(5)

def get_all_rates(code):
    cat = None
    wared = None
    vat = None
    arbah = None
    desc = ""
    for df in [df_tariff, df_book, df_gs1]:
        if df.empty:
            continue
        res = search_in_df(df, code)
        if res.empty:
            continue
        row = res.iloc[0]
        if not desc:
            dc = get_col(df, ["description", "name"])
            if dc:
                v = str(row[dc]).strip()
                if v not in ["", "nan", "None"]:
                    desc = v[:120]
        c1, w1, v1, a1 = extract_rates(row, df)
        if cat is None:
            cat = c1
        if wared is None:
            wared = w1
        if vat is None:
            vat = v1
        if arbah is None:
            arbah = a1
    if wared is None:
        wared = cat if cat is not None else 0
    if cat is None:
        cat = wared
    if vat is None:
        vat = 14.0
    if arbah is None:
        arbah = 1.0
    return cat, wared, vat, arbah, desc

async def send_long(message, text):
    if not text:
        text = "مفيش نتيجة"
    for i in range(0, len(text), 4000):
        await message.reply_text(text[i:i+4000])

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await send_long(update.message, "بوت الجمارك جاهز!\n/hs 8471 - بحث\n/calc 8471 1000 - حساب كامل\n/gpc 8471 - بحث GS1")

async def hs_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = " ".join(context.args) if context.args else ""
    if not q:
        await send_long(update.message, "ابعت كده: /hs 8471")
        return
    await do_search(update, q)

async def gpc_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = " ".join(context.args) if context.args else ""
    if not q:
        await send_long(update.message, "ابعت كده: /gpc 8471")
        return
    res = search_in_df(df_gs1, q)
    if res.empty:
        await send_long(update.message, f"مفيش نتيجة في GS1 لـ: {q}")
        return
    txt = f"GS1 نتائج: {q}\n\n"
    for _, r in res.iterrows():
        vals = [str(x).strip() for x in r.tolist() if str(x).strip() not in ["", "nan", "None"]]
        txt += " | ".join(vals[:4]) + "\n---\n"
    await send_long(update.message, txt)

async def calc_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if len(context.args) < 2:
        await send_long(update.message, "ابعت كده: /calc 8471 1000")
        return
    code = clean_code(context.args[0])
    try:
        value = float(context.args[1])
    except:
        value = 0
    cat, wared, vat, arbah, desc = get_all_rates(code)
    duty = value * wared / 100
    profit_tax = value * arbah / 100
    vat_val = (value + duty) * vat / 100
    total = duty + profit_tax + vat_val
    txt = "حساب البند: " + code + "\n" + desc + "\n\n"
    txt += "القيمة: " + str(value) + " $\n"
    txt += "الفئة العامة: " + str(cat) + "%\n"
    txt += "رسم الوارد: " + str(wared) + "% = " + str(round(duty, 2)) + " $\n"
    txt += "القيمة المضافة: " + str(vat) + "% = " + str(round(vat_val, 2)) + " $\n"
    txt += "الارباح: " + str(arbah) + "% = " + str(round(profit_tax, 2)) + " $\n"
    txt += "الاجمالي رسوم: " + str(round(total, 2)) + " $\n"
    txt += "الاجمالي بالقيمة: " + str(round(value + total, 2)) + " $"
    await send_long(update.message, txt)

async def do_search(update, q):
    res = search_in_df(df_tariff, q)
    src = "التعريفة"
    if res.empty:
        res = search_in_df(df_book, q)
        src = "Book1"
    if res.empty:
        res = search_in_df(df_gs1, q)
        src = "GS1"
    if res.empty:
        await send_long(update.message, f"مفيش نتيجة لـ: {q}")
        return
    txt = f"نتائج ({src}): {q}\n\n"
    for _, r in res.iterrows():
        vals = [str(x).strip() for x in r.tolist() if str(x).strip() not in ["", "nan", "None"]]
        txt += " | ".join(vals[:4]) + "\n---\n"
    if clean_code(q):
        cat, wared, vat, arbah, desc = get_all_rates(q)
        txt += f"\nالفئة: {cat}% | وارد: {wared}% | قيمة مضافة: {vat}% | ارباح: {arbah}%"
    await send_long(update.message, txt)

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await do_search(update, update.message.text.strip())

async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await send_long(update.message, "وصلت الصورة، ابعت اسم الصنف او الكود عشان احسب صح.")

def main():
    app = Application.builder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", start))
    app.add_handler(CommandHandler("hs", hs_cmd))
    app.add_handler(CommandHandler("gpc", gpc_cmd))
    app.add_handler(CommandHandler("calc", calc_cmd))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    print("Bot is running...")
    app.run_polling()

if __name__ == "__main__":
    main()