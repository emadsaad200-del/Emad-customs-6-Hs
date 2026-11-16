import os
import re
import glob
import pandas as pd
import telebot

TOKEN = os.getenv("BOT_TOKEN", "حط_التوكن_هنا")
bot = telebot.TeleBot(TOKEN)

# ============ الموقع بتاع تعديل فئة البند - ثابت ومش هنغيره تاني ============
SITE_URL = "https://www.customs.gov.eg"

# ============ تحميل الملفات ============
def load_file(path):
    if path.endswith(".csv"):
        for enc in ["cp1256", "utf-8-sig", "windows-1256", "latin1"]:
            for sep in [",", ";", "\t"]:
                try:
                    df_try = pd.read_csv(path, encoding=enc, sep=sep, engine="python", on_bad_lines="skip", low_memory=False)
                    if len(df_try.columns) >= 3:
                        return df_try
                except:
                    continue
        try:
            return pd.read_excel(path, engine="openpyxl")
        except:
            try:
                return pd.read_excel(path)
            except:
                return None
    else:
        try:
            return pd.read_excel(path)
        except:
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
            print(f"✅ التعريفة: {f}")
            return df_try.dropna(how="all")
    return pd.DataFrame()

def load_gs1():
    try:
        g = pd.read_excel("gs1 exle.xlsx")
        g.columns = [str(c).strip() for c in g.columns]
        print(f"DB built gs1: {len(g)}")
        return g
    except Exception as e:
        print(f"GS1 error: {e}")
        return pd.DataFrame()

df = load_tariff()
df_gs1 = load_gs1()
print(f"DB built tariffs: {len(df)}")
print("اعمدة التعريفة:", list(df.columns))
if not df_gs1.empty:
    print("اعمدة GS1:", list(df_gs1.columns))

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
    fea = r[COL_FEA] if COL_FEA else "غير موجود"
    desc_ar = r[COL_DESC_AR] if COL_DESC_AR else ""
    desc_en = ""
    if COL_DESC_EN and COL_DESC_EN != COL_DESC_AR:
        desc_en = r[COL_DESC_EN]
    unit = r[COL_UNIT] if COL_UNIT else ""
    vat = r[COL_VAT] if COL_VAT else "14"
    profit = r[COL_PROFIT] if COL_PROFIT else "غير موجود"
    txt = f"✅ البند: {r[COL_CODE]}\n"
    txt += f"📦 البيان عربي: {desc_ar}\n"
    if desc_en and str(desc_en) != "nan":
        txt += f"📦 البيان انجليزي: {desc_en}\n"
    txt += f"💰 الفئة / الوارد: {fea}%\n"
    txt += f"🧾 القيمة المضافة: {vat}%\n"
    txt += f"🏭 الأرباح التجارية الصناعية: {profit}%\n"
    txt += f"📏 الوحدة: {unit}"
    return txt

def calc_cif(row, amount):
    try:
        fea = float(str(row[COL_FEA]).replace("%","").strip()) if COL_FEA else 0
    except:
        fea = 0
    try:
        vat = float(str(row[COL_VAT]).replace("%","").strip()) if COL_VAT else 14
    except:
        vat = 14
    try:
        profit = float(str(row[COL_PROFIT]).replace("%","").strip()) if COL_PROFIT else 0
    except:
        profit = 0
    cif = amount
    customs = cif * fea / 100
    vat_val = (cif + customs) * vat / 100
    profit_val = cif * profit / 100
    total = cif + customs + vat_val + profit_val
    txt = format_item(row)
    txt += f"\n\n💵 مبلغ الفاتورة CIF: {cif:,.2f}\n"
    txt += f"🛃 رسم الوارد ({fea}%): {customs:,.2f}\n"
    txt += f"🧾 قيمة مضافة ({vat}%): {vat_val:,.2f}\n"
    txt += f"🏭 أرباح تجارية صناعية ({profit}%): {profit_val:,.2f}\n"
    txt += f"💰 الإجمالي: {total:,.2f}"
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
        return f"❌ مفيش نتيجة في GS1 لـ: {keyword}"
    res = f"🔍 نتائج GS1 لـ: {keyword}\n\n"
    for _, r in rows.head(10).iterrows():
        res += "---\n"
        for c in df_gs1.columns:
            v = r[c]
            if str(v) != "nan" and str(v).strip() != "":
                res += f"{c}: {