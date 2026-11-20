import os
import json
import time
import random
import re
import openpyxl
from telegram import Update
from telegram import InlineKeyboardButton
from telegram import InlineKeyboardMarkup
from telegram.ext import ApplicationBuilder
from telegram.ext import CommandHandler
from telegram.ext import MessageHandler
from telegram.ext import CallbackQueryHandler
from telegram.ext import ContextTypes
from telegram.ext import filters
BOT = os.getenv("BOT_TOKEN")
ADM = os.getenv("ADMIN_ID")
SL = chr(47)
os.makedirs("docs", exist_ok=True)
def LJ(f, d):
    try:
        return json.load(open(f, encoding="utf-8"))
    except:
        return d
def SJ(f, d):
    json.dump(d, open(f, "w", encoding="utf-8"))
def CC(v):
    if v is None:
        return ""
    return str(v).strip()
def NA(s):
    s = str(s or "").strip().lower()
    s = s.replace("أ", "ا")
    s = s.replace("إ", "ا")
    s = s.replace("آ", "ا")
    s = s.replace("ة", "ه")
    s = s.replace("ى", "ي")
    return s
def B(t, d):
    return InlineKeyboardButton(t, callback_data=d)
def MK(r):
    return InlineKeyboardMarkup(r)
def MAIN_KB():
    r1 = [B("🚢 توكيلات", "m_lines")]
    r1.append(B("📦 GPC", "m_gpc"))
    r2 = [B("🧾 التعريفة", "m_tariff")]
    r2.append(B("📁 المستندات", "m_docs"))
    r3 = [B("🚢 Rotation", "m_rot")]
    r3.append(B("🏢 الشركات", "m_comp"))
    r4 = [B("🌍 Country", "m_country")]
    r4.append(B("🏙️ Cities", "m_cities"))
    r5 = [B("💵 Currency", "m_cur")]
    r5.append(B("📦 Package", "m_pack"))
    r6 = [B("📏 UOM", "m_uom")]
    r6.append(B("🕘 اخر بحث", "m_recent"))
    r7 = [B("❓ مساعدة", "m_help")]
    return MK([r1, r2, r3, r4, r5, r6, r7])
def COMP_KB():
    r1 = [B("🏢 الخاصة", "ct_pri")]
    r2 = [B("🏛️ الحكومية", "ct_gov")]
    r3 = [B("🌍 الاجنبية", "ct_for")]
    r4 = [B("🏠 القايمة", "m_main")]
    return MK([r1, r2, r3, r4])
def LCAT_KB(k):
    r1 = [B("📘 Booking", "lc_" + k + "_BK")]
    r1.append(B("🧾 Invoice", "lc_" + k + "_IN"))
    r2 = [B("📧 الايميلات", "lc_" + k + "_AL")]
    r3 = [B("🏠 القايمة", "m_main")]
    return MK([r1, r2, r3])
BI = {}
BC = {}
BA = {}
for fn in ["gs1xlsx.xlsx"]:
    if not os.path.exists(fn):
        continue
    wb = openpyxl.load_workbook(fn, read_only=True)
    ws = wb[wb.sheetnames[0]]
    for r in ws.iter_rows(values_only=True):
        if not r or len(r) < 11:
            continue
        bt = str(r[10] or "")
        if not bt:
            continue
        k = bt.strip().lower()
        if k not in BI:
            d = {}
            d["c"] = str(r[9] or "")
            d["t"] = bt
            d["s"] = r[1]
            d["f"] = r[4]
            d["cl"] = r[7]
            d["i"] = r[11] or ""
            d["e"] = r[12] if len(r) > 12 else ""
            BI[k] = d
            BC[d["c"]] = k
        if len(r) > 14 and r[13] is not None:
            ac = str(r[13])
            if k not in BA:
                BA[k] = []
            ex = None
            for a in BA[k]:
                if a["c"] == ac:
                    ex = a
            if ex is None:
                ex = {"c": ac, "t": r[14], "v": []}
                BA[k].append(ex)
            vt = str(r[17]) if len(r) > 17 and r[17] else None
            if vt and vt not in ex["v"]:
                ex["v"].append(vt)
    wb.close()
REF = {}
for kind in ["cities", "country", "currency", "package", "uom"]:
    REF[kind] = {}
if os.path.exists("cities.xlsx"):
    wb = openpyxl.load_workbook("cities.xlsx", read_only=True)
    for kind, sh in [("cities", "City List"), ("country", "Country List"), ("currency", "Currency List"), ("package", "Package Type list"), ("uom", "UOM List")]:
        if sh not in wb.sheetnames:
            continue
        for r in wb[sh].iter_rows(values_only=True):
            if not r or not r[0] or not r[1]:
                continue
            if kind == "currency":
                c = str(r[1]).strip().upper()
                n = str(r[0]).strip()
            else:
                c = str(r[0]).strip().upper()
                n = str(r[1]).strip()
            if c in REF[kind]:
                continue
            REF[kind][c] = {"code": c, "name": n}
    wb.close()
CL = []
CBR = {}
def AC(n, f, rg, en, ac, st, ad, ct):
    n = CC(n)
    rg = CC(rg).replace(" ", "")
    if not n or not rg:
        return
    if rg in CBR:
        return
    it = {}
    it["name"] = n
    it["file"] = CC(f)
    it["reg"] = rg
    it["ent"] = CC(en)
    it["act"] = CC(ac)
    it["st"] = CC(st)
    it["ad"] = CC(ad)
    it["ct"] = ct
    CL.append(it)
    CBR[rg] = it
F1 = "ارقام_تسجيل_الشركات_111111.xlsx"
if os.path.exists(F1):
    wb = openpyxl.load_workbook(F1, read_only=True)
    ws = wb[wb.sheetnames[0]]
    for r in ws.iter_rows(values_only=True):
        if not r or not r[0]:
            continue
        if "الاسم" in str(r[0]):
            continue
        a1 = CC(r[1]) if len(r) > 1 else ""
        a2 = CC(r[2]) if len(r) > 2 else ""
        a3 = CC(r[3]) if len(r) > 3 else ""
        a4 = CC(r[4]) if len(r) > 4 else ""
        a5 = CC(r[5]) if len(r) > 5 else ""
        a6 = CC(r[6]) if len(r) > 6 else ""
        AC(r[0], a1, a2, a3, a4, a5, a6, "private")
    wb.close()
F2 = "ارقام_تسجيل_الشركات_2026.xlsx"
if os.path.exists(F2):
    wb = openpyxl.load_workbook(F2, read_only=True)
    if "خاص" in wb.sheetnames:
        for r in wb["خاص"].iter_rows(values_only=True):
            if not r or not r[0] or not r[1]:
                continue
            e = "خاص - " + CC(r[3])
            AC(r[1], "", r[0], e, "", CC(r[5]), CC(r[4]), "private")
    if "حكومي" in wb.sheetnames:
        for r in wb["حكومي"].iter_rows(values_only=True):
            if not r or not r[1] or not r[2]:
                continue
            AC(r[2], "", r[1], "حكومي", "", "", CC(r[3]), "gov")
    if "فرع اجنبي" in wb.sheetnames:
        for r in wb["فرع اجنبي"].iter_rows(values_only=True):
            if not r or not r[1] or not r[2]:
                continue
            AC(r[2], "", r[1], "فرع اجنبي", "", "", CC(r[3]), "foreign")
    wb.close()
def FC(c):
    t = "🏢 " + c["name"] + "\n"
    t = t + "🔢 رقم التسجيل: " + c["reg"] + "\n"
    if c["file"]:
        t = t + "📁 رقم الملف: " + c["file"] + "\n"
    if c["ent"]:
        t = t + "🏛️ الكيان: " + c["ent"] + "\n"
    if c["act"]:
        t = t + "📦 النشاط: " + c["act"] + "\n"
    if c["ad"]:
        t = t + "📍 العنوان: " + c["ad"] + "\n"
    return t.strip()
def GC(q, ct):
    out = [c for c in CL if not ct or c["ct"] == ct]
    nq = NA(q)
    if not nq:
        return out
    o = []
    for c in out:
        if nq in NA(c["name"]):
            o.append(c)
        elif nq in NA(c["act"]):
            o.append(c)
    return o
TL = []
TB = {}
def NTC(s):
    s = str(s or "").strip().replace(" ", "")
    if re.match(r"^\d{6}$", s):
        s = s[:4] + "." + s[4:]
    if re.match(r"^\d{8}$", s):
        s = s[:4] + "." + s[4:6] + "." + s[6:]
    return s
def load_tar(fn):
    wb = openpyxl.load_workbook(fn, read_only=True)
    ws = wb[wb.sheetnames[0]]
    for r in ws.iter_rows(values_only=True):
        if not r or len(r) < 3:
            continue
        code = NTC(r[0])
        desc = CC(r[1])
        if not code or not desc:
            continue
        if code in TB:
            continue
        it = {}
        it["code"] = code
        it["desc"] = desc
        it["tax"] = CC(r[2])
        TB[code] = it
        TL.append(it)
    wb.close()
if os.path.isdir("tarifa"):
    for f in sorted(os.listdir("tarifa")):
        if f.lower().endswith(".xlsx"):
            load_tar(os.path.join("tarifa", f))
for fn in ["30285-1.xlsx", "70490-2.xlsx"]:
    if os.path.exists(fn):
        load_tar(fn)
def ST(q):
    q = (q or "").strip()
    if not q:
        return []
    cq = NTC(q)
    if cq in TB:
        return [TB[cq]]
    nq = NA(q)
    out = []
    for t in TL:
        if nq in NA(t["desc"]):
            out.append(t)
            if len(out) >= 10:
                break
    return out
def FT(t):
    x = "🧾 " + t["desc"] + "\n"
    x = x + "🔢 البند: " + t["code"] + "\n"
    x = x + "💰 الضريبة: " + t["tax"]
    return x.strip()
LC = {}
LN = {}
ROT = []
F3 = "ايميلات_وارقام_توكيلات.xlsx"
if os.path.exists(F3):
    wb = openpyxl.load_workbook(F3, read_only=True)
    if "EMAILS" in wb.sheetnames:
        cur = ""
        rows = list(wb["EMAILS"].iter_rows(values_only=True))
        for r in rows[2:]:
            if not r:
                continue
            if r[0]:
                cur = str(r[0]).split("\n")[0].strip()
            if not cur:
                continue
            k = cur.upper()
            if k not in LN:
                LN[k] = cur
            nm = CC(r[1]) if len(r) > 1 else ""
            po = CC(r[2]) if len(r) > 2 else ""
            em = CC(r[3]) if len(r) > 3 else ""
            nu = CC(r[4]) if len(r) > 4 else ""
            bk = CC(r[5]) if len(r) > 5 else ""
            inv = CC(r[9]) if len(r) > 9 else ""
            if not nm and not em and not nu:
                continue
            if k not in LC:
                LC[k] = []
            d = {}
            d["name"] = nm
            d["pos"] = po
            d["email"] = em
            d["num"] = nu
            d["book"] = bk
            d["inv"] = inv
            LC[k].append(d)
    if "SERVICE ROTATION" in wb.sheetnames:
        for r in wb["SERVICE ROTATION"].iter_rows(values_only=True):
            if not r:
                continue
            pol = CC(r[0])
            pod = CC(r[1])
            if not pol or not pod:
                continue
            d = {}
            d["pol"] = pol
            d["pod"] = pod
            d["line"] = CC(r[2])
            d["tt"] = CC(r[3])
            d["co"] = CC(r[4])
            ROT.append(d)
    wb.close()
AM = {}
AM["ميرسك"] = "MAERSK"
AM["هاباج"] = "HAPAG"
AM["كوسكو"] = "COSCO"
AM["سي ام ايه"] = "CMA"
AM["ايفرجرين"] = "EVERGREEN"
AM["ام اس سي"] = "MSC"
AM["زيم"] = "ZIM"
AM["اركاس"] = "ARKAS"
AM["وان"] = "ONE"
AM["ترانسمار"] = "TRANSMAR"
def RL(q):
    q = (q or "").strip()
    if not q:
        return None
    if q.upper() in LC:
        return q.upper()
    if q.lower() in AM:
        return AM[q.lower()]
    ql = q.lower()
    for k in LN:
        if ql in k.lower():
            return k
    return None
def FL(k, cat):
    cs = LC.get(k, [])
    dp = LN.get(k, k)
    if not cs:
        return "❌ ملقيتش: