import pandas as pd
import requests
import re
import shutil
import glob
import os

# ============ دالة قراءة أي ملف - جديد - بتحل مشكلة Tariff.xlsx.csv ============
def load_file(path):
    # 1- لو الملف .csv جرب تقراه CSV بكل الطرق
    if path.endswith(".csv"):
        for enc in ["cp1256", "utf-8-sig", "windows-1256", "latin1"]:
            for sep in [",", ";", "\t"]:
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
                        print(f"اتقرأ {path} كـ CSV بـ {enc}")
                        return df_try
                except:
                    continue
        # 2- لو فشل CSV، جربه Excel - هو ده الحل بتاع Tariff.xlsx.csv
        # الملف ده اكسيل حقيقي بس اسمه .csv بالغلط
        print(f"CSV فشل في {path} - هنجربه Excel")
        try:
            df_excel = pd.read_excel(path, engine="openpyxl")
            print(f"اتقرأ {path} كـ Excel - الأعمدة: {len(df_excel.columns)}")
            return df_excel
        except Exception as e:
            print(f"Excel بـ openpyxl فشل في {path}: {e}")
            try:
                return pd.read_excel(path)
            except:
                return None
    else:
        # 3- لو الملف .xlsx اقراه Excel على طول
        try:
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
        # الملف الصح لازم يكون فيه بند أو فئة أو وارد
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

# ============ حفظ - قديم + جديد - بنفس نوع الملف الأصلي ============
try:
    # لو الملف كان csv هنحفظه csv بنفس الترميز اللي اتقرأ بيه
    # بس لو هو كان اكسيل متسمي csv، هنحفظه Excel صح
    if FILE.endswith(".csv"):
        # جرب نحفظه Excel الأول عشان هو أصلا Excel
        try:
            book.to_excel(FILE, index=False)
            print(f"اتحفظ كـ Excel في {FILE}")
        except:
            book.to_csv(FILE, index=False, encoding="cp1256")
            print(f"اتحفظ كـ CSV بـ cp1256")
    else:
        book.to_excel(FILE, index=False)
        print(f"اتحفظ بـ Excel")
except Exception as e:
    print(f"الحفظ فشل: {e} - هنجرب utf-8-sig")
    book.to_csv(FILE, index=False, encoding="utf-8-sig")

print(f"تم الحفظ - اتحدث: {updated} - فشل: {failed}")
print("تم الحفظ - فئة ووارد اتحدثوا")