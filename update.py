import pandas as pd
import requests
from bs4 import BeautifulSoup
import re
import time
import shutil
from datetime import datetime

shutil.copy("Book1.xlsx", f"Book1_backup_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")
book = pd.read_excel("Book1.xlsx")
COL_CODE = "الكود"
COL_CAT = "الفئة العامة"
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
    time.sleep(0.6)

book.to_excel("Book1.xlsx", index=False)
print("خلصنا!")