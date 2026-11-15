import pandas as pd
import requests, re, glob, shutil

FILE = glob.glob("*.xlsx")[0]
print(f"الملف: {FILE}")
shutil.copy(FILE, FILE.replace(".xlsx", "_backup.xlsx"))

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

book = pd.read_excel(FILE, header=header_row)
print("الأعمدة:", list(book.columns))

# نلاقي عمود الكود
col_code = next((c for c in book.columns if any(k in str(c) for k in ["بند", "HS", "كود", "رمز"])), book.columns[0])

# نلاقي كل أعمدة الفئة / الوارد - الاتنين واحد
cols_fea_wared = [c for c in book.columns if "فئ" in str(c) or "وارد" in str(c)]
print(f"عمود الكود: {col_code} | أعمدة الفئة/الوارد: {cols_fea_wared}")

def get_rate(code):
    try:
        code = str(code).strip().split(".")[0]
        if not code.isdigit():
            return None
        url = f"https://www.gomruky.com/tariff/{code}"
        r = requests.get(url, timeout=15, headers={"User-Agent": "Mozilla/5.0"})
        m = re.search(r"(فئة|وارد).*?(\d+)\s*%", r.text)
        if m:
            return m.group(2)
        m2 = re.search(r"(\d+)\s*%", r.text)
        return m2.group(1) if m2 else None
    except:
        return None

if not cols_fea_wared:
    print("مفيش عمود فئة ولا وارد!")
else:
    for i, row in book.iterrows():
        code = row[col_code]
        new = get_rate(code)
        if new:
            for c in cols_fea_wared:
                book.at[i, c] = new
            print(f"{code} -> فئة/وارد {new}%")
        else:
            print(f"{code} فشل")

    book.to_excel(FILE, index=False)
    print("تم الحفظ - فئة ووارد اتحدثوا")