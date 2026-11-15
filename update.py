import pandas as pd
import requests, re, os, shutil, glob

FILE = glob.glob("*.xlsx")[0]
print(f"الملف: {FILE}")

# ناخد نسخة احتياطية
shutil.copy(FILE, FILE.replace(".xlsx", "_backup.xlsx"))

# نقرأ من غير ما نفترض مكان العناوين
df_raw = pd.read_excel(FILE, header=None)
print("عدد الصفوف:", len(df_raw))

# ندور على الصف اللي فيه كلمة بند او كود او HS
header_row = None
for i in range(min(15, len(df_raw))):
    row_text = " ".join([str(x) for x in df_raw.iloc[i].values])
    if "بند" in row_text or "HS" in row_text or "كود" in row_text or "رمز" in row_text:
        header_row = i
        print(f"لقينا العناوين في الصف رقم: {i}")
        break

if header_row is None:
    header_row = 3
    print("ملقيناش صف عناوين واضح، هنستخدم الصف 3")

book = pd.read_excel(FILE, header=header_row)
print("الأعمدة اللي لقيناها:", list(book.columns))

# نلاقي عمود الكود وعمود ضريبة الوارد تلقائيا
cols = [str(c) for c in book.columns]
col_code = next((c for c in book.columns if any(k in str(c) for k in ["بند", "HS", "كود", "رمز", "Code"])), book.columns[0])
col_wared = next((c for c in book.columns if "وارد" in str(c)), book.columns[1])

print(f"عمود الكود: {col_code} | عمود الوارد: {col_wared}")

def get_wared(code):
    try:
        code = str(code).strip().split(".")[0]
        if not code.isdigit():
            return None
        url = f"https://www.gomruky.com/tariff/{code}"
        r = requests.get(url, timeout=15, headers={"User-Agent": "Mozilla/5.0"})
        m = re.search(r"ضريبة الوارد.*?(\d+)", r.text)
        return m.group(1) if m else None
    except:
        return None

for i, row in book.iterrows():
    code = row[col_code]
    new = get_wared(code)
    if new:
        book.at[i, col_wared] = new
        print(f"{code} -> {new}")
    else:
        print(f"{code} فشل - مش كود صحيح او الموقع مردش")

book.to_excel(FILE, index=False)
print("تم الحفظ")