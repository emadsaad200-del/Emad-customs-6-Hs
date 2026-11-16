import pandas as pd
import requests
import re
import shutil
import glob
import os

# ============ دالة قراءة أي ملف - جديد - بتحل مشكلة سطر 233 ============
def load_file(path):
    try:
        if path.endswith(".csv"):
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
def find_tariff_file():
    files = glob.glob("*.csv") + glob.glob("*.xlsx")
    print(f"كل الملفات: {files}")
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
        if "بند" in cols or "فئ" in cols or "وارد" in cols or "HS" in cols:
            print(f"✅ ملف التعريفة الصح هو: {f}")
            return f, df_try.dropna(how="all")
    print("❌ مفيش ملف فيه بند/فئة/وارد")
    return None, pd.DataFrame()

FILE, book = find_tariff_file()

if FILE is None:
    print("وقفنا - مفيش ملف تعريفة")
    exit()

print(f"الملف اللي هيتحدث: {FILE}")
print("الأعمدة:", list(book.columns))

# ============ باك اب قبل التحديث - قديم ============
BACKUP = FILE + "_backup.csv"
try:
    shutil.copy(FILE, BACKUP)
    print(f"اتعمل باك اب: {BACKUP}")
except Exception as e:
    print(f"مش قادر يعمل باك اب: {e}")

# ============ لاقي الأعمدة - فئة = وارد - قديم + جديد ============
col_code = next((c for c in book.columns if any(k in str(c) for k in ["بند", "HS", "كود", "رمز"])), book.columns[0])
cols_fea_wared = [c for c in book.columns if "فئ" in str(c) or "وارد" in str(c)]

print(f"عمود الكود: {col_code}")
print(f"أعمدة الفئة/الوارد: {cols_fea_wared}")
print(f"عدد البنود: {len(book)}")

if not cols_fea_wared:
    print("مفيش عمود فئة ولا وارد! وقفنا - اتأكد من اسم العمود في Tariff.xlsx.csv")
    exit()

# ============ دالة جلب الفئة من gomruky - قديم ============
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

# ============ تحديث كل البنود - قديم ============
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

# ============ حفظ - قديم + جديد - بنفس ترميز الملف الأصلي ============
try:
    # لو الملف كان csv هنحفظه csv بنفس الترميز اللي اتقرأ بيه
    if FILE.endswith(".csv"):
        book.to_csv(FILE, index=False, encoding="cp1256")
        print(f"اتحفظ بـ cp1256")
    else:
        book.to_excel(FILE, index=False)
        print(f"اتحفظ بـ excel")
except Exception as e:
    print(f"الحفظ بـ cp1256 فشل: {e} - هنجرب utf-8-sig")
    book.to_csv(FILE, index=False, encoding="utf-8-sig")

print(f"تم الحفظ - اتحدث: {updated} - فشل: {failed}")
print("تم الحفظ - فئة ووارد اتحدثوا")