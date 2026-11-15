import pandas as pd
import requests
from bs4 import BeautifulSoup
import re
import time
import shutil
from datetime import datetime

FILE = "Book1LOGIST TATTEF.xlsx"

# نسخة احتياطية
shutil.copy(FILE, f"Book1_backup_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")
book = pd.read_excel(FILE)

print("الاعمدة اللي لقيناها:", list(book.columns))

# لو اسم العمود مختلف، هنلقطه اتوماتيك
# اول عمود فيه ارقام طويلة = الكود، والعمود اللي فيه % = الفئة
COL_CODE = None
COL_CAT = None
for c in book.columns:
    s = str(c).strip()
    if s in ["الكود", "كود", "HS", "HS Code", "hs_code", "البند", "بند"]:
        COL_CODE = c
    if "فئة" in s or "الفئة" in s or "ضريبة" in s or "وارد" in s:
        COL_CAT = c

# لو ملقيناش، افترض ان اول عمود هو الكود، وتاني عمود فيه % هو الفئة
if not COL_CODE:
    COL_CODE = book.columns[0]
if not COL_CAT:
    # دور على اي عمود فيه %
    for c in book.columns:
        if book[c].astype(str).str.contains("%").any():
            COL_CAT = c
            break

print(f"هنستخدم الكود من: {COL_CODE} والفئة من: {COL_CAT}")

session = requests.Session()
session.headers.update({"User-Agent": "Mozilla/5.0"})

def get_wared_from_gomruky(hs_code):
    hs = str(hs_code).strip().replace(".0", "").replace(" ", "")
    hs = hs.ljust(10, "0")[:10]
    url = f"https://www.gomruky.com/tariff/{hs}"
    try:
        r = session.get(url, timeout=15)
        if r.status_code != 200:
            return None
        soup = BeautifulSoup(r.text, 'html.parser')
        text = soup.get_text(" ", strip=True)
        m = re.search(r'ضريبة الوارد.*?(\d+(?:\.\d+)?)\s*%', text)
        if m:
            val = float(m.group(1))
            return f"{int(val)}%" if val.is_integer() else f"{val}%"
        return None
    except:
        return None

for i, row in book.iterrows():
    code = row[COL_CODE]
    new = get_wared_from_gomruky(code)
    if new:
        book.at[i, COL_CAT] = new
        print(f"{code} -> {new}")
    else:
        print(f"{code} فشل")
    time.sleep(0.6)

book.to_excel(FILE, index=False)
print("خلصنا!")