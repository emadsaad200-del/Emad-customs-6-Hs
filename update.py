import pandas as pd
import requests
import re
import shutil
import os

# ============ ملفاتك في GitHub ============
FILE = "Tariff.xlsx.csv"
BACKUP = "Tariff.xlsx.csv_backup.csv"

print(f"الملف اللي هيتحدث: {FILE}")

# ============ باك اب قبل التحديث ============
try:
    shutil.copy(FILE, BACKUP)
    print(f"اتعمل باك اب: {BACKUP}")
except Exception as e:
    print(f"مش قادر يعمل باك اب: {e}")

# ============ قراءة ملف التعريفة ============
try:
    book = pd.read_csv(FILE, encoding="utf-8-sig", low_memory=False)
except Exception as e1:
    print(f"utf-8 فشل: {e1} - هنجرب cp1256")
    book = pd.read_csv(FILE, encoding="cp1256", low_memory=False)

print("الأعمدة:", list(book.columns))
book.columns = [str(c).strip() for c in book.columns]
book = book.dropna(how="all")
print(f"عدد البنود: {len(book)}")

# ============ لاقي الأعمدة - فئة = وارد ============
col_code = next((c for c in book.columns if any(k in str(c) for k in ["بند", "HS", "كود", "رمز"])), book.columns[0])
cols_fea_wared = [c for c in book.columns if "فئ" in str(c) or "وارد" in str(c)]

print(f"عمود الكود: {col_code}")
print(f"أعمدة الفئة/الوارد: {cols_fea_wared}")

if not cols_fea_wared:
    print("مفيش عمود فئة ولا وارد! وقفنا - اتأكد من اسم العمود في Tariff.xlsx.csv")
    exit()

# ============ دالة جلب الفئة من gomruky ============
def get_rate(code):
    try:
        code = str(code).strip().split(".")[0].replace(".0", "")
        if not code.isdigit():
            return None
        url = f"https://www.gomruky.com/tariff/{code}"
        r = requests.get(url, timeout=15, headers={"User-Agent": "Mozilla/5.0"})
        # ندور على رقم جنب كلمة فئة أو وارد
        m = re.search(r"(فئة|وارد).*?(\d+)\s*%", r.text)
        if m:
            return m.group(2)
        m2 = re.search(r"(\d+)\s*%", r.text)
        return m2.group(1) if m2 else None
    except Exception as e:
        print(f"خطأ في {code}: {e}")
        return None

# ============ تحديث كل البنود ============
updated = 0
failed = 0

for i, row in book.iterrows():
    code = row[col_code]
    new = get_rate(code)
    if new:
        for c in cols_fea_wared:
            book.at[i, c] = new
        print(f"{code} -> فئة/وارد {new}%")
        updated += 1
    else:
        print(f"{code} فشل")
        failed += 1

# ============ حفظ ============
book.to_csv(FILE, index=False, encoding="utf-8-sig")
print(f"تم الحفظ - اتحدث: {updated} - فشل: {failed}")
print("تم الحفظ - فئة ووارد اتحدثوا")