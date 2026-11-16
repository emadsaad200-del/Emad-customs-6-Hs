import os
import re
import pandas as pd
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

TOKEN = os.getenv("BOT_TOKEN")

# ============ البحث عن الملفات بأي شكل ============
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
            try:
                return pd.read_csv(path, dtype=str, low_memory=False, encoding="cp1256", errors="ignore")
            except:
                return pd.DataFrame()

def norm_col(c):
    return str(c).strip().replace(" ", "").replace("_", "").replace("\n", "")

def find_header_row(df_raw):
    for i in range(min(12, len(df_raw))):
        row = " ".join([str(x) for x in df_raw.iloc[i].tolist()]).lower()
        if ("بند" in row or "code" in row or "hs" in row) and ("وصف" in row or "description" in row or "بيان" in row):
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

print(f"Files found: Book={BOOK_FILE} Tariff={TARIFF_FILE} GS1={GS1_FILE}")

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
            if nn in k or k in nn:
                return v
    return None

def clean_code(s):
    return re.sub(r"\D", "", str(s))

def clean_val(v):
    # بيشيل nan و فاضي و % و مسافات، ويرجع رقم
    if v is None:
        return None
    s = str(v).strip().lower().replace("nan", "").replace("none", "").replace("%", "").replace("٪", "").strip()
    if s == "" or s == "-":
        return None
    # الارقام العربي
    s = s.translate(str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789"))
    m = re.search(r"\d+(\.\d+)?", s)
    if not m:
        return None
    try:
        return float(m.group(0))
    except:
        return None

def extract_rates_from_row(row, df):
    # بيطلع ال 4 نسب من صف واحد، بيدور على اسم العمود
    rates = {"fئة": None, "wared": None, "vat": None, "arbah": None}
    for col in df.columns:
        c = str(col)
        cn = norm_col(c)
        val = clean_val(row[col])
        if val is None:
            continue
        # القيمة المضافة
        if "قيم" in c or "مضاف" in c or "vat" in c.lower():
            rates["vat"] = val
        # الارباح
        elif "ربح