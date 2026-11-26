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

BOT = os.getenv("BOT_TOKEN") or os.getenv("TELEGRAM_BOT_TOKEN") or os.getenv("TELEGRAM_TOKEN")
ADM = os.getenv("ADMIN_ID")
SL = chr(47)
os.makedirs("docs", exist_ok=True)

END_IMG = "end_session.jpg"
END_WORDS = ["خلصت", "شكرا", "شكرا لك", "تمام كده", "كفايه", "كفاية", "سلام", "bye", "thanks", "thank you"]
END_TIMEOUT = 600

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
def GV(r, i):
    if r is None:
        return ""
    if len(r) <= i:
        return ""
    if r[i] is None:
        return ""
    return r[i]
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
    r1 = [B("🚢 توكيلات", "m_lines"), B("📦 GPC", "m_gpc")]
    r2 = [B("📜 التعريفة", "m_tariff"), B("📄 مستندات", "m_docs")]
    r3 = [B("🔄 Rotation", "m_rot"), B("🏢 شركات", "m_comp")]
    r4 = [B("🌍 Country", "m_country"), B("🏙️ Cities", "m_cities")]
    r5 = [B("💱 Currency", "m_cur"), B("📦 Package", "m_pack")]
    r6 = [B("⚖️ UOM", "m_uom"), B("🕐 اخر بحث", "m_recent")]
    r7 = [B("🧮 حسبة الرسوم التقريبية", "m_calc")]
    r7b = [B("📦 حسبة CBM", "m_cbm")]
    r8 = [B("❓ مساعدة", "m_help")]
    return MK([r1, r2, r3, r4, r5, r6, r7, r7b, r8])
def COMP_KB():
    r1 = [B("🏢 الخاصة", "ct_pri")]
    r2 = [B("🚢 وكلاء الشحن", "ct_agent")]
    r3 = [B("🏛️ الحكومية", "ct_gov")]
    r4 = [B("🌍 الاجنبية", "ct_for")]
    r5 = [B("🏠 القايمة الرئيسية", "m_main")]
    return MK([r1, r2, r3, r4, r5])
def LCAT_KB(k):
    r1 = [B("📩 Booking", "lc_" + k + "_BK")]
    r1.append(B("🧾 Invoice", "lc_" + k + "_IN"))
    r2 = [B("📧 الايميلات", "lc_" + k + "_AL")]
    r3 = [B("🏠 القايمة الرئيسية", "m_main")]
    return MK([r1, r2, r3])

def DOCS_ADMIN_KB():
    r1 = [B("➕ اضافة مستند", "docs_add")]
    r2 = [B("📋 قائمة المستندات", "docs_list")]
    r3 = [B("🗑️ حذف كل المستندات", "docs_del_all")]
    r4 = [B("🏠 القايمة الرئيسية", "m_main")]
    return MK([r1, r2, r3, r4])

BI = {}
BC = {}
BA = {}
def LOAD_GPC_FILE(fn):
    if not os.path.exists(fn):
        print(f"GPC FILE NOT FOUND: {fn}")
        return
    try:
        sz = os.path.getsize(fn)//1024
        print(f"Loading GPC file: {fn} - {sz} KB")
        wb = openpyxl.load_workbook(fn, read_only=True)
        ws = wb[wb.sheetnames[0]]
        cnt = 0
        for r in ws.iter_rows(values_only=True):
            if not r:
                continue
            r = list(r)
            r += [""] * 15
            r = tuple(r)
            bt = str(GV(r, 10) or "")
            if not bt:
                continue
            k = bt.strip().lower()
            if k not in BI:
                d = {}
                d["c"] = str(GV(r, 9) or "")
                d["t"] = bt
                d["s"] = GV(r, 1)
                d["f"] = GV(r, 4)
                d["cl"] = GV(r, 7)
                d["i"] = GV(r, 11) or ""
                d["e"] = GV(r, 12) or ""
                BI[k] = d
                BC[d["c"]] = k
                cnt += 1
            if GV(r, 13) != "":
                ac = str(GV(r, 13))
                if k not in BA:
                    BA[k] = []
                ex = None
                for a in BA[k]:
                    if a["c"] == ac:
                        ex = a
                if ex is None:
                    ex = {"c": ac, "t": GV(r, 14), "v": []}
                    BA[k].append(ex)
        wb.close()
        print(f"Done GPC {fn} -> added {cnt}, total now {len(BI)}")
    except Exception as e:
        print(f"GPC load fail {fn}: {e}")

LOAD_GPC_FILE("GPC as of May 2023 v20230605 EN.xlsx")
LOAD_GPC_FILE("gs1xlsx.xlsx")
for f in sorted(os.listdir(".")):
    low = f.lower()
    if "gpc" in low and low.endswith(".xlsx") and f not in ["GPC as of May 2023 v20230605 EN.xlsx", "gs1xlsx.xlsx"]:
        LOAD_GPC_FILE(f)
print(f"GPC loaded: {len(BI)} items")

REF = {}
for kind in ["cities", "country", "currency", "package", "uom"]:
    REF[kind] = {}

def ADD_REF(kind, code, name):
    code = str(code or "").strip().upper()
    name = str(name or "").strip()
    if not code or not name:
        return
    if code in ["CODE", "CITYCODE", "CURRENCY", "NAME"]:
        return
    if code in REF[kind]:
        return
    REF[kind][code] = {"code": code, "name": name}

if os.path.exists("cities.xlsx"):
    wb = openpyxl.load_workbook("cities.xlsx", read_only=True)
    for kind, sh in [("cities", "City List"), ("country", "Country List"), ("currency", "Currency List"), ("package", "Package Type list"), ("uom", "UOM List")]:
        if sh not in wb.sheetnames:
            continue
        for r in wb[sh].iter_rows(values_only=True):
            if not r or not GV(r, 0) or not GV(r, 1):
                continue
            if kind == "currency":
                ADD_REF(kind, GV(r, 1), GV(r, 0))
            else:
                ADD_REF(kind, GV(r, 0), GV(r, 1))
    wb.close()
    print(f"REF OLD loaded cities.xlsx -> { {k: len(v) for k, v in REF.items()} }")
else:
    print("REF FILE MISSING: cities.xlsx")

def LOAD_NEW_REF_FILE(fn, kind, code_col, name_col):
    if not os.path.exists(fn):
        print(f"REF NEW MISSING: {fn}")
        return 0
    try:
        wb = openpyxl.load_workbook(fn, read_only=True)
        ws = wb[wb.sheetnames[0]]
        before = len(REF[kind])
        for r in ws.iter_rows(values_only=True):
            if not r:
                continue
            ADD_REF(kind, GV(r, code_col), GV(r, name_col))
        wb.close()
        added = len(REF[kind]) - before
        print(f"REF NEW loaded {fn} -> +{added} total {kind}: {len(REF[kind])}")
        return added
    except Exception as e:
        print(f"REF NEW fail {fn}: {e}")
        return 0

LOAD_NEW_REF_FILE("Country_List.xlsx", "country", 0, 1)
LOAD_NEW_REF_FILE("Cities.xlsx", "cities", 0, 1)
LOAD_NEW_REF_FILE("Currency_List.xlsx", "currency", 1, 0)
LOAD_NEW_REF_FILE("Package_Type_list.xlsx", "package", 0, 1)
LOAD_NEW_REF_FILE("Unit_of_measure_list.xlsx", "uom", 0, 1)
print(f"REF MERGED TOTAL: country={len(REF['country'])} cities={len(REF['cities'])} currency={len(REF['currency'])} package={len(REF['package'])} uom={len(REF['uom'])}")

REF_AR = {
    "مصر": "EGYPT", "امريكا": "UNITED STATES", "الصين": "CHINA", "السعوديه": "SAUDI ARABIA", "السعودية": "SAUDI ARABIA",
    "الامارات": "UNITED ARAB EMIRATES", "المانيا": "GERMANY", "فرنسا": "FRANCE", "ايطاليا": "ITALY", "اسبانيا": "SPAIN",
    "انجلترا": "UNITED KINGDOM", "بريطانيا": "UNITED KINGDOM", "تركيا": "TURKEY", "الهند": "INDIA",
    "اليابان": "JAPAN", "كوريا": "KOREA", "روسيا": "RUSSIAN FEDERATION", "البرازيل": "BRAZIL", "كندا": "CANADA",
    "استراليا": "AUSTRALIA", "هولندا": "NETHERLANDS", "بلجيكا": "BELGIUM", "سويسرا": "SWITZERLAND",
    "القاهره": "CAIRO", "القاهرة": "CAIRO", "الاسكندريه": "ALEXANDRIA", "الاسكندرية": "ALEXANDRIA",
    "بورسعيد": "PORT SAID", "دمياط": "DAMIETTA", "السويس": "SUEZ", "السخنه": "SOKHNA", "السخنة": "SOKHNA",
    "دبي": "DUBAI", "الرياض": "RIYADH", "جده": "JEDDAH", "جدة": "JEDDAH", "الدوحه": "DOHA", "الدوحة": "DOHA",
    "دولار": "USD", "الدولار": "USD", "يورو": "EUR", "اليورو": "EUR", "جنيه": "EGP", "الجنيه": "EGP",
    "جنيه مصري": "EGP", "ريال": "SAR", "ريال سعودي": "SAR", "درهم": "AED", "دينار": "KWD", "ين": "JPY",
    "صندوق": "BOX", "كرتونه": "CARTON", "كرتونة": "CARTON", "بالته": "PALLET", "بالتة": "PALLET",
    "برميل": "DRUM", "شوال": "BAG", "كيس": "BAG", "حاويه": "CONTAINER", "حاوية": "CONTAINER",
    "كيلو": "KG", "كيلوجرام": "KG", "طن": "TON", "متر": "METER", "قطعه": "PIECE", "قطعة": "PIECE",
    "لتر": "LITRE", "علبه": "BOX", "حبه": "PIECE",
}
REF_EN_TO_AR = {}
for _ar, _en in REF_AR.items():
    if _en not in REF_EN_TO_AR:
        REF_EN_TO_AR[_en] = _ar

CL = []
CBR = {}
def AC(n, f, rg, en, ac, st, ad, ct, alt=""):
    n = CC(n)
    rg = CC(rg).replace(" ", "").replace("-", "")
    if not n or not rg:
        return
    if "قطاع" in n and "حكومي" in n:
        return
    if "الاسم" in n:
        return
    if rg in CBR:
        ex = CBR[rg]
        if alt and alt != ex.get("name", "") and alt not in ex.get("alt", ""):
            ex["alt"] = (ex.get("alt", "") + " / " + alt).strip(" /")
        if en and en not in ex.get("ent", ""):
            ex["ent"] = (ex.get("ent", "") + " | " + en).strip(" |")
        if ct == "agent" and ex.get("ct") == "private":
            ex["ct"] = "agent"
        if ad and not ex.get("ad"):
            ex["ad"] = ad
        if ac and not ex.get("act"):
            ex["act"] = ac
        if f and not ex.get("file"):
            ex["file"] = f
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
    it["alt"] = CC(alt)
    CL.append(it)
    CBR[rg] = it

F2_OLD = "ارقام_تسجيل_الشركات_2026.xlsx"
F2_NEW = "ارقام_تسجيل_الشركات_ووكلاء_الشحن_محدث_نهائي.xlsx"
if os.path.exists(F2_OLD):
    try:
        wb = openpyxl.load_workbook(F2_OLD, read_only=True)
        if "خاص" in wb.sheetnames:
            for r in wb["خاص"].iter_rows(values_only=True):
                if not r or not GV(r, 0) or not GV(r, 1):
                    continue
                e = "خاص - " + CC(GV(r, 3))
                AC(GV(r, 1), "", GV(r, 0), e, "", CC(GV(r, 5)), CC(GV(r, 4)), "private")
        if "حكومي" in wb.sheetnames:
            for r in wb["حكومي"].iter_rows(values_only=True):
                if not r or not GV(r, 1) or not GV(r, 2):
                    continue
                AC(GV(r, 2), "", GV(r, 1), "حكومي", "", "", CC(GV(r, 3)), "gov")
        if "فرع اجنبي" in wb.sheetnames:
            for r in wb["فرع اجنبي"].iter_rows(values_only=True):
                if not r or not GV(r, 1) or not GV(r, 2):
                    continue
                AC(GV(r, 2), "", GV(r, 1), "فرع اجنبي", "", "", CC(GV(r, 3)), "foreign")
        wb.close()
        print(f"Companies OLD loaded: {len(CL)}")
    except Exception as e:
        print(f"Companies OLD fail: {e}")
else:
    print(f"COMPANIES OLD MISSING: {F2_OLD}")
if os.path.exists(F2_NEW):
    try:
        wb = openpyxl.load_workbook(F2_NEW, read_only=True)
        sh = "القائمة الموحدة النهائية"
        if sh in wb.sheetnames:
            rows = list(wb[sh].iter_rows(values_only=True))
            for r in rows[4:]:
                if not r or not GV(r, 1) or not GV(r, 2):
                    continue
                nm = CC(GV(r, 1))
                rg = CC(GV(r, 2))
                alt = CC(GV(r, 3))
                cat = CC(GV(r, 4))
                src = CC(GV(r, 5))
                if "وكيل" in cat or "توكيل" in cat:
                    ct = "agent"
                else:
                    ct = "private"
                ent = cat
                if src:
                    ent += " | " + src
                AC(nm, "", rg, ent, "", "", "", ct, alt)
            print(f"Companies NEW merged, total now: {len(CL)}")
        else:
            print(f"Companies NEW sheet not found: {sh} in {wb.sheetnames}")
        wb.close()
    except Exception as e:
        print(f"Companies NEW fail: {e}")
else:
    print(f"COMPANIES NEW MISSING: {F2_NEW}")
print(f"Companies loaded TOTAL (merged): {len(CL)}")

def FC(c):
    t = "🏢 " + c["name"] + "\n"
    t += "رقم التسجيل: " + c["reg"] + "\n"
    if c.get("alt"):
        t += "اسم بديل / تجاري: " + c["alt"] + "\n"
    if c["file"]:
        t += "رقم الملف: " + c["file"] + "\n"
    if c["ent"]:
        t += "الكيان: " + c["ent"] + "\n"
    if c["act"]:
        t += "النشاط: " + c["act"] + "\n"
    if c["ad"]:
        t += "العنوان: " + c["ad"] + "\n"
    if "وكيل" in c.get("ent", ""):
        t += "التصنيف: 🚢 وكيل شحن / توكيل ملاحي\n"
    return t.strip()

def GC(q, ct):
    if ct == "agent":
        out = [c for c in CL if c["ct"] == "agent"]
    elif ct == "private":
        out = [c for c in CL if c["ct"] == "private"]
    elif ct:
        out = [c for c in CL if c["ct"] == ct]
    else:
        out = list(CL)
    nq = NA(q)
    if not nq:
        return out
    words = [w for w in nq.split() if len(w) >= 2 and w not in ["شركه", "شركة", "الشركه"]]
    o = []
    for c in out:
        nn = NA(c["name"])
        na = NA(c.get("act", ""))
        ne = NA(c.get("ent", ""))
        nalt = NA(c.get("alt", ""))
        reg = str(c["reg"]).replace(" ", "")
        if nq.replace(" ", "") in reg or reg in nq.replace(" ", ""):
            o.append(c)
            continue
        if words:
            if any(w in nn or w in na or w in ne or w in nalt for w in words):
                o.append(c)
        else:
            if nq in nn or nq in na or nq in ne or nq in nalt:
                o.append(c)
    return o

TL = []
TB = {}
def NTC(s):
    s = str(s or "").strip().replace(" ", "").replace("-", ".")
    s = s.strip(".")
    if re.match(r"^\d{4}$", s):
        s = s[:2] + "." + s[2:]
    elif re.match(r"^\d{6}$", s):
        s = s[:4] + "." + s[4:]
    elif re.match(r"^\d{8}$", s):
        s = s[:4] + "." + s[4:6] + "." + s[6:]
    elif re.match(r"^\d{10}$", s):
        s = s[:4] + "." + s[4:6] + "." + s[6:8] + "." + s[8:]
    return s

def load_tar(fn):
    try:
        wb = openpyxl.load_workbook(fn, read_only=True)
    except Exception as e:
        print(f"Tariff load fail {fn}: {e}")
        return
    ws = wb[wb.sheetnames[0]]
    before = len(TL)
    for r in ws.iter_rows(values_only=True):
        if not r:
            continue
        code = NTC(GV(r, 0))
        desc = CC(GV(r, 1))
        if not code or not desc:
            continue
        if code in TB:
            if not TB[code].get("tax") and CC(GV(r, 2)):
                TB[code]["tax"] = CC(GV(r, 2))
            continue
        it = {}
        it["code"] = code
        it["desc"] = desc
        it["tax"] = CC(GV(r, 2))
        TB[code] = it
        TL.append(it)
    wb.close()
    print(f"Loaded {fn} -> +{len(TL)-before} total {len(TL)}")

if os.path.isdir("tarifa"):
    for f in sorted(os.listdir("tarifa")):
        if f.lower().endswith(".xlsx"):
            load_tar(os.path.join("tarifa", f))
for i in range(1, 31):
    fn1 = f"tarifa_{i:02d}.xlsx.xlsx"
    fn2 = f"tarifa_{i:02d}.xlsx"
    if os.path.exists(fn1):
        load_tar(fn1)
    elif os.path.exists(fn2):
        load_tar(fn2)
for f in sorted(os.listdir(".")):
    low = f.lower()
    if low.startswith("tarifa_") and low.endswith(".xlsx"):
        already = False
        for i in range(1, 31):
            if f == f"tarifa_{i:02d}.xlsx.xlsx" or f == f"tarifa_{i:02d}.xlsx":
                already = True
        if not already:
            load_tar(f)
for fn in ["30285-1.xlsx", "70490-2.xlsx"]:
    if os.path.exists(fn):
        load_tar(fn)
print(f"Tariff loaded: {len(TL)} items")

def ST(q):
    q = (q or "").strip()
    if not q:
        return []
    cq = NTC(q)
    if cq in TB:
        return [TB[cq]]
    qd = q.replace(".", "").replace(" ", "").replace("-", "")
    if qd.isdigit():
        out = []
        for code, it in TB.items():
            cd = code.replace(".", "")
            if cd == qd or cd.startswith(qd) or qd.startswith(cd[:4]):
                out.append(it)
                if len(out) >= 50:
                    break
        if out:
            return out
    nq = NA(q)
    out = []
    for t in TL:
        if nq in NA(t["desc"]) or nq in NA(t["code"]) or q in t["code"]:
            out.append(t)
            if len(out) >= 50:
                break
    return out

def FT(t):
    x = "📜 " + t["desc"] + "\n"
    x += "البند: " + t["code"] + "\n"
    x += "الضريبة: " + t["tax"]
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
            if GV(r, 0):
                cur = str(GV(r, 0)).split("\n")[0].strip()
            if not cur:
                continue
            k = cur.upper()
            if k not in LN:
                LN[k] = cur
            nm = CC(GV(r, 1)); po = CC(GV(r, 2)); em = CC(GV(r, 3)); nu = CC(GV(r, 4)); bk = CC(GV(r, 5)); inv = CC(GV(r, 9))
            if not nm and not em and not nu:
                continue
            if k not in LC:
                LC[k] = []
            d = {}
            d["name"] = nm; d["pos"] = po; d["email"] = em; d["num"] = nu; d["book"] = bk; d["inv"] = inv
            LC[k].append(d)
    if "SERVICE ROTATION" in wb.sheetnames:
        for r in wb["SERVICE ROTATION"].iter_rows(values_only=True):
            if not r:
                continue
            pol = CC(GV(r, 0)); pod = CC(GV(r, 1))
            if not pol or not pod:
                continue
            d = {}
            d["pol"] = pol; d["pod"] = pod; d["line"] = CC(GV(r, 2)); d["tt"] = CC(GV(r, 3)); d["co"] = CC(GV(r, 4))
            ROT.append(d)
    wb.close()
    print(f"Lines loaded: {len(LN)}")
else:
    print(f"LINES FILE MISSING: {F3}")

AM = {}
AM["ميرسك"] = "MAERSK"; AM["هاباج"] = "HAPAG"; AM["كوسكو"] = "COSCO"; AM["ايفرجرين"] = "EVERGREEN"; AM["ام اس سي"] = "MSC"; AM["زيم"] = "ZIM"; AM["اركاس"] = "ARKAS"; AM["وان"] = "ONE"
def RL(q):
    q = (q or "").strip()
    if not q:
        return None
    if q.upper() in LC:
        return q.upper()
    if q.lower() in AM:
        return AM[q.lower()]
    ql = NA(q)
    for k in LN:
        if ql in NA(LN[k]) or NA(LN[k]) in ql:
            return k
    return None
def FL(k, cat):
    cs = LC.get(k, [])
    dp = LN.get(k, k)
    if not cs:
        return "🚢 ملقيتش: " + dp
    t = "🚢 " + dp + "\n"
    if cat == "BK":
        t += "Booking:\n"
        for x in cs:
            t += x["name"] + " - " + x["email"] + "\n" + x["book"] + "\n"
        return t.strip()
    if cat == "IN":
        t += "Invoice:\n"
        for x in cs:
            if x["inv"]:
                t += x["name"] + " - " + x["email"] + "\n" + x["inv"] + "\n"
        return t.strip()
    t += "العدد: " + str(len(cs)) + "\n\n"
    for c in cs[:10]:
        if c["name"]: t += c["name"] + "\n"
        if c["pos"]: t += c["pos"] + "\n"
        if c["email"]: t += c["email"] + "\n"
        if c["num"]: t += c["num"] + "\n"
        if c["book"]: t += "Booking: " + c["book"] + "\n"
        if c["inv"]: t += "Invoice: " + c["inv"] + "\n"
        t += "\n"
    return t.strip()
def SROT(pol, pod):
    out = []
    for r in ROT:
        if NA(pol) in NA(r["pol"]) or NA(r["pol"]) in NA(pol):
            if NA(pod) in NA(r["pod"]) or NA(r["pod"]) in NA(pod):
                out.append(r)
    return out[:10]
def FROT(rs, pol, pod):
    if not rs:
        return "🔄 ملقيتش خط من " + pol + " لـ " + pod
    t = "🔄 من " + pol + " لـ " + pod + ":\n\n"
    for r in rs:
        t += r["line"] + " | " + r["pol"] + " -> " + r["pod"]
        if r["tt"]: t += " - " + r["tt"]
        t += " | " + r["co"] + "\n"
    return t[:3500]
    ROT2 = []
ROT2_SEEN = set()
PORT_AR_EN = {
    "الاسكندريه": "alexandria", "اسكندريه": "alexandria", "الاسكندرية": "alexandria", "اسكندرية": "alexandria", "alex": "alexandria", "alexandria": "الاسكندرية",
    "بورسعيد": "port said", "بور سعيد": "port said", "port said": "بورسعيد", "portsaid": "port said",
    "دمياط": "damietta", "damietta": "دمياط",
    "السخنه": "sokhna", "السخنة": "sokhna", "سخنه": "sokhna", "سخنة": "sokhna", "sokhna": "السخنة",
    "الدخيله": "dekheila", "الدخيلة": "dekheila", "دخيله": "dekheila", "دخيلة": "dekheila",
    "شنغهاي": "shanghai", "شانغهاي": "shanghai", "shanghai": "شنغهاي",
    "بوسان": "busan", "busan": "بوسان",
    "نينجبو": "ningbo", "ningbo": "نينجبو",
    "سنغافوره": "singapore", "سنغافورة": "singapore", "singapore": "سنغافورة",
    "ازميت": "izmit", "ازمير": "izmir", "izmir": "ازمير", "izmit": "ازميت",
    "اسطنبول": "istanbul", "استانبول": "istanbul", "istanbul": "اسطنبول",
    "جده": "jeddah", "جدة": "jeddah", "jeddah": "جدة",
    "العقبه": "aqaba", "العقبة": "aqaba", "aqaba": "العقبة",
    "بيروت": "beirut", "beirut": "بيروت",
    "بيريوس": "piraeus", "piraeus": "بيريوس",
    "روتردام": "rotterdam", "rotterdam": "روتردام",
    "هامبورج": "hamburg", "هامبورغ": "hamburg", "hamburg": "هامبورج",
    "انتويرب": "antwerp", "antwerp": "انتويرب",
    "فالنسيا": "valencia", "valencia": "فالنسيا",
    "برشلونه": "barcelona", "برشلونة": "barcelona", "barcelona": "برشلونة",
    "جنوه": "genoa", "جنوة": "genoa", "genoa": "جنوه",
    "مالطا": "malta", "malta": "مالطا",
    "جيبوتي": "djibouti", "djibouti": "جيبوتي",
    "كوبير": "koper", "koper": "كوبير",
    "تريستا": "trieste", "trieste": "تريستا",
    "طنجة": "tangier", "طنجه": "tangier", "tangier": "طنجة",
    "نيويورك": "new york", "نيو يورك": "new york", "new york": "نيويورك", "newyork": "new york",
    "نورفولك": "norfolk", "norfolk": "نورفولك",
    "سافانا": "savannah", "savannah": "سافانا",
    "ميامي": "miami", "miami": "ميامي",
    "هيوستن": "houston", "houston": "هيوستن",
    "نيو اورليانز": "new orleans", "new orleans": "نيو اورليانز",
    "لوس انجلوس": "los angeles", "لوس أنجلوس": "los angeles", "los angeles": "لوس انجلوس",
    "لونج بيتش": "long beach", "long beach": "لونج بيتش",
    "اوكلاند": "oakland", "oakland": "اوكلاند",
    "سياتل": "seattle", "seattle": "سياتل",
    "هاليفاكس": "halifax", "halifax": "هاليفاكس",
    "سانتوس": "santos", "santos": "سانتوس",
    "ريو دي جانيرو": "rio de janeiro", "rio de janeiro": "ريو دي جانيرو",
    "بوينس ايرس": "buenos aires", "بوينس آيرس": "buenos aires", "buenos aires": "بوينس ايرس",
    "كارتاجينا": "cartagena", "cartagena": "كارتاجينا",
    "كالاو": "callao", "callao": "كالاو",
    "سان انطونيو": "san antonio", "san antonio": "سان انطونيو",
    "فيروكروز": "veracruz", "veracruz": "فيروكروز",
    "مونباسا": "mombasa", "mombasa": "مونباسا",
    "لاجوس": "lagos", "lagos": "لاجوس",
    "داكار": "dakar", "dakar": "داكار",
    "الدار البيضاء": "casablanca", "casablanca": "الدار البيضاء",
    "فيليكستو": "felixstowe", "felixstowe": "فيليكستو",
    "ليفربول": "liverpool", "liverpool": "ليفربول",
    "دبلن": "dublin", "dublin": "دبلن",
    "عدن": "aden", "aden": "عدن",
    "نوفوروسيسك": "novorossiysk", "novorossiysk": "نوفوروسيسك",
}
LINE_AR_EN = {
    "ميرسك": "maersk", "maersk": "ميرسك",
    "ام اس سي": "msc", "إم إس سي": "msc", "msc": "ام اس سي",
    "سي ام ايه": "cma cgm", "سي ام ايه سي جي ام": "cma cgm", "cma": "سي ام ايه", "cma cgm": "سي ام ايه",
    "هاباج": "hapag", "هاباج لويد": "hapag", "hapag": "هاباج", "hapag-lloyd": "هاباج",
    "كوسكو": "cosco", "cosco": "كوسكو",
    "ايفرجرين": "evergreen", "evergreen": "ايفرجرين",
    "زيم": "zim", "zim": "زيم",
    "اركاس": "arkas", "arkas": "اركاس",
    "وان": "one", "one": "وان",
    "اوشن": "ocean", "ocean": "اوشن",
    "ميرسك لاين": "maersk",
    "اتش ام ام": "hmm", "hmm": "اتش ام ام",
    "او او سي ال": "oocl", "oocl": "او او سي ال",
    "بي اي ال": "pil", "pil": "بي اي ال",
    "وان هاي": "wan hai", "wan hai": "وان هاي",
    "يانج مينج": "yang ming", "yang ming": "يانج مينج",
    "ymx": "yang ming", "ym": "yang ming", "yml": "yang ming",
    "توركون": "turkon", "turkon": "توركون",
    "تاروس": "tarros", "tarros": "تاروس",
    "جريمالدي": "grimaldi", "grimaldi": "جريمالدي",
    "فولك": "folk", "folk": "فولك",
    "ملاحه": "milaha", "milaha": "ملاحه",
    "ادميرال": "admiral", "admiral": "ادميرال",
}
def NORM_ROT_Q(q):
    s = NA(q)
    s = s.replace("-", " ").replace("_", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s
def EXPAND_ROT_Q(q):
    nq = NORM_ROT_Q(q)
    outs = [nq]
    for k, v in list(PORT_AR_EN.items()) + list(LINE_AR_EN.items()):
        if NA(k) == nq or NA(v) == nq:
            outs.append(NA(k))
            outs.append(NA(v))
    parts = nq.split()
    if len(parts) > 1:
        for p in parts:
            outs.extend(EXPAND_ROT_Q(p))
    return list(set([o for o in outs if o]))
def ROT_MATCH(text, q):
    if not q or not text:
        return False
    nt = NORM_ROT_Q(text)
    for cand in EXPAND_ROT_Q(q):
        if cand and (cand in nt or nt in cand):
            return True
    return False
def LOAD_ONE_ROT2_FILE(fn):
    global ROT2
    if not os.path.exists(fn):
        return 0
    try:
        wb = openpyxl.load_workbook(fn, read_only=True)
        ws = wb[wb.sheetnames[0]]
        added = 0
        started = False
        for r in ws.iter_rows(values_only=True):
            if not r:
                continue
            c0 = CC(GV(r, 0))
            if not started:
                if "Shipping Line" in c0 and "التوكيل" in c0:
                    started = True
                continue
            if "Filtered" in c0 or "Total Services" in c0:
                break
            if not c0:
                continue
            line = CC(GV(r, 0))
            agent = CC(GV(r, 1))
            service = CC(GV(r, 2))
            egy_port = CC(GV(r, 3))
            lane = CC(GV(r, 4))
            region = CC(GV(r, 5))
            rotation = CC(GV(r, 6))
            avg = CC(GV(r, 7))
            rng = CC(GV(r, 8))
            freq = CC(GV(r, 9))
            if not line and not service and not rotation:
                continue
            key = (line.lower() + "|" + service.lower() + "|" + rotation.lower()[:80])
            if key in ROT2_SEEN:
                continue
            ROT2_SEEN.add(key)
            d = {}
            d["pol"] = egy_port or line
            d["pod"] = rotation
            d["line"] = (line + " - " + service).strip(" -")
            d["tt"] = (str(avg) + " يوم") if avg else rng
            d["co"] = freq or lane
            d["rot2"] = True
            d["agent"] = agent
            d["service"] = service
            d["egy_port"] = egy_port
            d["lane"] = lane
            d["region"] = region
            d["rotation"] = rotation
            d["avg"] = avg
            d["rng"] = rng
            d["freq"] = freq
            d["src"] = fn
            ROT2.append(d)
            added += 1
        wb.close()
        print(f"ROT2 loaded {fn} -> +{added} total {len(ROT2)}")
        return added
    except Exception as e:
        print(f"ROT2 load fail {fn}: {e}")
        return 0
def LOAD_ROT2_FILE():
    cands = []
    fixed = ["egypt_shipping_lines_unified_single_sheet1111111111.xlsx", "egypt_shipping_lines_unified_single_sheet1111111111.xlsx.xlsx", "egypt_shipping_lines-v4.xlsx", "egypt_shipping_lines_v4.xlsx"]
    for fn in fixed:
        if fn not in cands:
            cands.append(fn)
    for f in sorted(os.listdir(".")):
        low = f.lower()
        if low.endswith(".xlsx") and ("egypt_shipping" in low or "shipping_lines" in low):
            if f not in cands:
                cands.append(f)
    total = 0
    for fn in cands:
        total += LOAD_ONE_ROT2_FILE(fn)
    if total == 0:
        print("ROT2 FILE MISSING: no egypt_shipping file found")
    else:
        print(f"ROT2 MERGED TOTAL: {len(ROT2)} from {len(cands)} files")
LOAD_ROT2_FILE()
def SROT2(q):
    q = (q or "").strip()
    if not q or not ROT2:
        return []
    words = [w for w in NORM_ROT_Q(q).split() if len(w) >= 2]
    if len(words) >= 2:
        out = []
        for r in ROT2:
            hay = " ".join([r.get("line",""), r.get("agent",""), r.get("service",""), r.get("egy_port",""), r.get("lane",""), r.get("region",""), r.get("rotation","")])
            if all(ROT_MATCH(hay, w) for w in words):
                out.append(r)
        return out
    out = []
    for r in ROT2:
        hay = " ".join([r.get("line",""), r.get("agent",""), r.get("service",""), r.get("egy_port",""), r.get("lane",""), r.get("region",""), r.get("rotation","")])
        if ROT_MATCH(hay, q):
            out.append(r)
    return out
def FROT2(rs, q=""):
    if not rs:
        return ["🔄 ملقيتش روتيشن لـ " + (q or "")]
    chunks = []
    for start in range(0, len(rs), 3):
        part = rs[start:start+3]
        t = "🔄 الروتيشن" + (" لـ " + q if q else "") + f" - جزء {start//3 + 1} من {(len(rs)+2)//3} - العدد الكلي: {len(rs)}:\n\n"
        for r in part:
            t += "🚢 " + r.get("line","") + "\n"
            if r.get("agent"): t += "الوكيل: " + r["agent"] + "\n"
            if r.get("egy_port"): t += "ميناء التحميل في مصر: " + r["egy_port"] + "\n"
            if r.get("lane"): t += "خط التجارة: " + r["lane"] + "\n"
            if r.get("region"): t += "المنطقة: " + r["region"] + "\n"
            if r.get("rotation"):
                rot = r["rotation"]
                if len(rot) > 250:
                    rot = rot[:250] + " ..."
                t += "خط السير: " + rot + "\n"
            if r.get("avg"): t += "متوسط الرحلة: " + str(r["avg"]) + " يوم\n"
            if r.get("rng"): t += "النطاق: " + r["rng"] + "\n"
            if r.get("freq"): t += "معدل الإبحار: " + r["freq"] + "\n"
            t += "\n"
        chunks.append(t.strip())
    return chunks
async def SEND_ALL_ROT2(update, rs, q=""):
    for chunk in FROT2(rs, q):
        await update.message.reply_text(chunk)
def FROT_MIXED(r):
    if r.get("rot2"):
        ch = FROT2([r], r.get("line",""))
        return ch[0] if ch else ""
    return FROT([r], r.get("pol",""), r.get("pod",""))

WEL = "Welcome to Captain Emad Customs Bot"
HELP = """❓ المساعدة - كل أيقونة بتعمل ايه:

🚢 توكيلات: اكتب MAERSK او ميرسك
📦 GPC: اكتب GPC Airbrushes او GPC + الاسم او الكود
📜 التعريفة: اكتب 8420 او 84 او اسم الصنف - مثال: مروحة - او 4802 / 48.02
🏢 شركات: اكتب النور او رقم التسجيل - وكيل شحن اكتب ARKAS او HAPAG او رقم التسجيل 242755704
🔄 Rotation: اكتب Alex Izmir او الاسكندرية ازمير او MSC او ميرسك او Sokhna او السخنة او New York او نيويورك او Houston او Santos او Aden او عدن او Novorossiysk
🌍 Country: اكتب country مصر او country Egypt او country EG
🏙️ Cities: اكتب city القاهرة او city Cairo
💱 Currency: اكتب currency دولار او currency USD او currency جنيه
📦 Package: اكتب package صندوق او package Box
⚖️ UOM: اكتب uom كيلو او uom KG او uom طن
📄 مستندات: اكتب تزيل 1
🧮 حسبة الرسوم التقريبية: دوس عليها - البوت هيسألك مبلغ الفاتورة - نوع العملة - التعاقد - فئة البند - سعر التحويل
📦 حسبة CBM: دوس عليها - البوت هيسألك الطول والعرض والارتفاع والوزن وعدد الكراتين
🕐 اخر بحث: بيجيب اخر حاجة بحثت عنها
القايمة: اكتب menu
"""

CALC_CONTRACTS = ["CIF", "C&F", "FCA", "FOB 12.5%", "FOB نولون", "EXW 12.5%", "EXW نولون", "FOB Free"]
CALC_CURS = ["USD", "EUR", "GBP", "SAR", "AED", "CNY", "JPY", "EGP"]
def CALC_KB_CONTRACT():
    rows = []
    for c in CALC_CONTRACTS:
        rows.append([B(c, "calc_ct_" + c)])
    rows.append([B("❌ الغاء", "m_main")])
    return MK(rows)
def CALC_KB_CUR():
    rows = []
    for i in range(0, len(CALC_CURS), 4):
        rows.append([B(c, "calc_cur_" + c) for c in CALC_CURS[i:i+4]])
    rows.append([B("❌ الغاء", "m_main")])
    return MK(rows)
def CALC_RESET(ctx):
    ctx.user_data["calc"] = {"step": "inv", "inv": 0, "cur": "", "contract": "", "duty": 0, "rate": 0, "other": 0}
    ctx.user_data["calc_on"] = True
def CALC_DO(d):
    inv_egp = d["inv"] * d["rate"]
    customs = inv_egp * (d["duty"] / 100)
    vat = (inv_egp + customs) * 0.14
    profit = inv_egp * 0.01
    total = customs + vat + profit + d["other"]
    return {"inv_egp": inv_egp, "cif": inv_egp, "customs": customs, "vat": vat, "profit": profit, "total": total}
def CALC_TXT(d, r):
    t = "🧮 حسبة الرسوم التقريبية\n"
    t += "━━━━━━━━━━━━━━\n"
    t += f"💰 مبلغ الفاتورة: {d['inv']:,.2f} {d['cur']}\n"
    t += f"💱 سعر التحويل: {d['rate']}\n"
    t += f"🇪🇬 قيمة الفاتورة بالمصري: {r['inv_egp']:,.2f} جنيه\n"
    t += f"📦 التعاقد: {d['contract']}\n"
    t += f"📜 فئة البند: {d['duty']}%\n"
    t += "━━━━━━━━━━━━━━\n"
    t += f"🏛️ الجمارك ({d['duty']}%): {r['customs']:,.2f} جنيه\n"
    t += f"🧾 قيمة مضافة 14%: {r['vat']:,.2f} جنيه\n"
    t += f"💼 أرباح تجارية 1%: {r['profit']:,.2f} جنيه\n"
    if d['other']:
        t += f"📦 مصاريف أخرى: {d['other']:,.2f} جنيه\n"
    t += "━━━━━━━━━━━━━━\n"
    t += f"✅ إجمالي الرسوم التقريبية: {r['total']:,.2f} جنيه\n"
    t += "\n⚠️ حسبة تقريبية"
    return t

# ===== NEW: حسبة CBM - اضافة فقط =====
CBM_CONTAINERS = [("20FT", 590, 235, 239), ("40FT", 1203, 235, 239), ("40FT HC", 1204, 235, 269)]
def CBM_KB_UNIT():
    return MK([[B("📐 بالمتري - سم / كجم", "cbm_unit_cm"), B("📏 بالقدم - بوصة / رطل", "cbm_unit_in")], [B("❌ الغاء", "m_main")]])
def CBM_KB_DIV():
    return MK([[B("🚢 بحري / بري - قسمة على 5000", "cbm_div_5000"), B("✈️ جوي - قسمة على 6000", "cbm_div_6000")], [B("❌ الغاء", "m_main")]])
def CBM_RESET(ctx):
    ctx.user_data["cbm"] = {"step": "unit", "unit": "cm", "div": 5000, "l": 0, "w": 0, "h": 0, "wt": 0, "qty": 0}
    ctx.user_data["cbm_on"] = True
def CBM_DO(d):
    if d["unit"] == "cm":
        vol_one = d["l"] * d["w"] * d["h"] / 1000000
        vol_w_one = d["l"] * d["w"] * d["h"] / d["div"]
    else:
        vol_one = d["l"] * d["w"] * d["h"] / 1728
        vol_w_one = d["l"] * d["w"] * d["h"] / d["div"] * 2.20462
    total_vol = vol_one * d["qty"]
    total_w = d["wt"] * d["qty"]
    total_vol_w = vol_w_one * d["qty"]
    chargeable = max(total_w, total_vol_w)
    cont_txt = []
    for name, cl, cw, ch in CBM_CONTAINERS:
        if d["unit"] == "cm":
            cont_vol = cl * cw * ch / 1000000
        else:
            cont_vol = (cl/2.54) * (cw/2.54) * (ch/2.54) / 1728
        if vol_one > 0:
            approx = int(cont_vol / vol_one * 0.85)
            approx_max = int(cont_vol / vol_one * 0.95)
            cont_txt.append(f"{name}: تقريباً {approx} - {approx_max} كرتونة")
    courier = []
    if d["unit"] == "cm":
        max_side = max(d["l"], d["w"], d["h"])
        for comp, max_w, max_len in [("DHL", 70, 120), ("FedEx", 68, 119), ("UPS", 70, 108)]:
            ow = "❌ Over Weight" if d["wt"] > max_w else "✅ وزن مقبول"
            osz = "❌ Over Size" if max_side > max_len else "✅ مقاس مقبول"
            courier.append(f"{comp}: {ow} | {osz}")
    else:
        max_side = max(d["l"], d["w"], d["h"])
        for comp, max_w, max_len in [("DHL", 154, 47), ("FedEx", 150, 47), ("UPS", 154, 42)]:
            ow = "❌ Over Weight" if d["wt"] > max_w else "✅ وزن مقبول"
            osz = "❌ Over Size" if max_side > max_len else "✅ مقاس مقبول"
            courier.append(f"{comp}: {ow} | {osz}")
    return {"vol_one": vol_one, "total_vol": total_vol, "total_w": total_w, "total_vol_w": total_vol_w, "chargeable": chargeable, "cont": cont_txt, "courier": courier}
def CBM_TXT(d, r):
    unit_vol = "CBM" if d["unit"] == "cm" else "CFT"
    unit_w = "كجم" if d["unit"] == "cm" else "رطل"
    unit_l = "سم" if d["unit"] == "cm" else "بوصة"
    t = "📦 حسبة CBM\n"
    t += "━━━━━━━━━━━━━━\n"
    t += f"📐 المقاس: {d['l']} × {d['w']} × {d['h']} {unit_l}\n"
    t += f"⚖️ وزن الكرتونة: {d['wt']} {unit_w}\n"
    t += f"📦 عدد الكراتين: {d['qty']}\n"
    t += f"📏 القسمة الحجمية: / {d['div']}\n"
    t += "━━━━━━━━━━━━━━\n"
    t += f"📦 حجم الكرتونة: {r['vol_one']:.4f} {unit_vol}\n"
    t += f"📦 الحجم الكلي: {r['total_vol']:.2f} {unit_vol}\n"
    t += f"⚖️ الوزن الكلي: {r['total_w']:,.2f} {unit_w}\n"
    t += f"⚖️ الوزن الحجمي الكلي: {r['total_vol_w']:,.2f} {unit_w}\n"
    t += f"✅ الوزن المحاسب عليه: {r['chargeable']:,.2f} {unit_w} (الأعلى)\n"
    t += "━━━━━━━━━━━━━━\n"
    t += "🚢 الكونتينرات:\n"
    for c in r["cont"]:
        t += c + "\n"
    t += "━━━━━━━━━━━━━━\n"
    t += "✈️ شحن سريع:\n"
    for c in r["courier"]:
        t += c + "\n"
    return t

def LU():
    return LJ("users.json", {})
def SU(u):
    SJ("users.json", u)
def OK(uid):
    if str(uid) == str(ADM):
        return True
    u = LU()
    if str(uid) not in u:
        return False
    return u[str(uid)].get("status") == "accepted"

def FB(b):
    t = "📦 " + b["t"] + "\n"
    t += "Brick: " + b["c"] + "\n"
    t += "Segment: " + str(b["s"]) + "\n"
    t += "Family: " + str(b["f"]) + "\n"
    t += "Class: " + str(b["cl"]) + "\n"
    if b["i"]: t += "Includes: " + str(b["i"]) + "\n"
    if b["e"]: t += "Excludes: " + str(b["e"]) + "\n"
    at = BA.get(b["t"].strip().lower(), [])
    if at:
        t += "Attributes:\n"
        for a in at:
            t += str(a["t"]) + " (" + str(a["c"]) + ")\n"
            if a["v"]: t += ", ".join(a["v"][:10]) + "\n"
    return t

def SGPC(q):
    qq = re.sub(r"\s+", " ", (q or "").strip().lower())
    if not qq:
        return None
    dg = re.sub(r"\D", "", qq)
    if dg and dg in BC:
        return BI[BC[dg]]
    if qq in BI:
        return BI[qq]
    m = [BI[k] for k in BI if qq in k]
    if len(m) == 1:
        return m[0]
    if len(m) > 1:
        return ("m", m[:10])
    words = [w for w in qq.split() if len(w) >= 3]
    if words:
        m2 = [BI[k] for k in BI if any(w in k for w in words)]
        if len(m2) == 1:
            return m2[0]
        if len(m2) > 1:
            return ("m", m2[:10])
    return None

def PG(items, page, pre, lab):
    per = 10
    st = page * per
    tp = max(1, (len(items) + per - 1) // per)
    rows = []
    for i in range(st, min(st + per, len(items))):
        cb = pre + "sel_" + str(i)
        rows.append([B(lab(items[i]), cb)])
    nav = []
    if page > 0:
        nav.append(B("⬅️ السابق", pre + "page_" + str(page - 1)))
    if page < tp - 1:
        nav.append(B("التالي ➡️", pre + "page_" + str(page + 1)))
    if nav:
        rows.append(nav)
    rows.append([B("🏠 القايمة الرئيسية", "m_main")])
    txt = "اختار صفحة " + str(page + 1) + " من " + str(tp) + " - العدد: " + str(len(items))
    return MK(rows), txt

def SREF(kind, q):
    items = sorted(REF[kind].values(), key=lambda x: x["code"])
    if not q:
        return items
    nq = NA(q)
    q_en = REF_AR.get(nq, "")
    if not q_en:
        q_en = REF_AR.get(nq.replace("ال", "", 1), "")
    q_ar = REF_EN_TO_AR.get(q.upper(), "")
    out = []
    for x in items:
        name_n = NA(x["name"])
        code_n = NA(x["code"])
        if nq in name_n or nq in code_n or q.upper() in x["code"] or nq in x["code"].lower():
            out.append(x)
            continue
        if q_en and (NA(q_en) in name_n or q_en in x["code"]):
            out.append(x)
            continue
        if q_ar and NA(q_ar) in name_n:
            out.append(x)
            continue
    return out

def FREF(kind, it):
    ar = REF_EN_TO_AR.get(it["name"].upper(), "")
    if ar:
        return f"{it['name']}\n{ar}\nالكود: {it['code']}"
    return f"{it['name']}\nالكود: {it['code']}"

def DEL_ONE_DOC(num):
    ds = LJ("docs_index.json", [])
    new_ds = []
    found = False
    for x in ds:
        if x["num"] == num:
            found = True
            try:
                if os.path.exists(x["path"]):
                    os.remove(x["path"])
            except:
                pass
        else:
            new_ds.append(x)
    if found:
        SJ("docs_index.json", new_ds)
    return found
def DEL_ALL_DOCS():
    ds = LJ("docs_index.json", [])
    for x in ds:
        try:
            if os.path.exists(x["path"]):
                os.remove(x["path"])
        except:
            pass
    SJ("docs_index.json", [])
    return len(ds)
def LIST_DOCS_TXT():
    ds = LJ("docs_index.json", [])
    if not ds:
        return "📄 مفيش مستندات لسه"
    t = "📄 قائمة المستندات:\n\n"
    for x in sorted(ds, key=lambda y: y["num"]):
        t += f"{x['num']} - {x['name']}\n"
    t += "\nاكتب: تزيل 1 عشان تنزل المستند"
    return t

SUB_INFO_TXT = """👋 أهلا بيك في Captain Emad Customs Bot
البوت ده خدمة مدفوعة باشتراك رمزي 💰
الاشتراك: 150 جنيه في الشهر فقط
بيوفرلك:
🚢 توكيلات وايميلات
📜 التعريفة الجمركية
📦 GPC
🏢 بيانات الشركات
🔄 Rotation
🧮 حسبة الرسوم الجمركية
لو حابب تشترك، ابعتلنا هنا وهنبعتلك طريقة الدفع ✅
ولو مش حابب، مفيش مشكلة نورتنا ❤️
"""

async def SEND_END_IMAGE(ctx, chat_id):
    try:
        ud = None
        try:
            ud = ctx.user_data
        except:
            ud = None
        if ud is not None and ud.get("end_sent"):
            return
        if ud is not None:
            try:
                ud["end_sent"] = True
            except:
                pass
    except Exception:
        pass
    if os.path.exists(END_IMG):
        try:
            await ctx.bot.send_photo(chat_id=chat_id, photo=open(END_IMG, "rb"), caption="لو حابب تكمل أو تشترك، دي بيانات التواصل بتاعتنا ❤️")
            return
        except:
            pass
    try:
        await ctx.bot.send_message(chat_id=chat_id, text="لو حابب تكمل أو تشترك:\nTelegram: https://t.me/tarrifabot\nWhatsApp: wa.me/+201094038730\nPhone: 01094038730")
    except Exception as e:
        print(f"SEND_END_IMAGE fail: {e}")

async def END_JOB(ctx):
    try:
        chat_id = ctx.job.chat_id
    except:
        return
    try:
        await SEND_END_IMAGE(ctx, chat_id)
    except Exception as e:
        print(f"END_JOB fail: {e}")

def RESET_END_TIMER(ctx):
    try:
        if ctx.job_queue is None:
            return
        chat_id = None
        try:
            chat_id = ctx._chat_id
        except:
            pass
        if chat_id is None:
            try:
                chat_id = ctx.job.chat_id
            except:
                pass
        if chat_id is None:
            return
        for j in list(ctx.job_queue.jobs()):
            try:
                if j.name and j.name.startswith("end_timer_") and j.chat_id == chat_id:
                    j.schedule_removal()
            except:
                pass
        ctx.job_queue.run_once(END_JOB, END_TIMEOUT, chat_id=chat_id, name="end_timer_" + str(chat_id))
    except Exception as e:
        print(f"RESET_END_TIMER fail: {e}")

def SAVE_RECENT(ctx, text):
    try:
        ud = ctx.user_data
    except:
        return
    if ud is None:
        return
    lst = ud.get("recent", [])
    lst.append(text)
    ud["recent"] = lst[-10:]

async def SEND(up, ctx, kind, page, q):
    if up.message:
        uid = up.message.from_user.id
    else:
        uid = up.callback_query.from_user.id
    ctx.user_data["fq"] = q
    ctx.user_data["fk"] = kind
    if kind == "comp":
        ct = ctx.user_data.get("ct", "")
        items = GC(q, ct)
        kb, txt = PG(items, page, "comp_", lambda c: "🏢 " + c["name"][:35] + " - " + c["reg"])
    elif kind == "tariff":
        items = TL
        if q:
            items = ST(q)
        kb, txt = PG(items, page, "tariff_", lambda t: "📜 " + t["code"] + " - " + t["desc"][:35])
    elif kind == "line":
        items = sorted(LN.keys(), key=lambda x: LN[x])
        if q:
            items = [k for k in items if NA(q) in NA(LN[k])]
        kb, txt = PG(items, page, "line_", lambda k: "🚢 " + LN[k][:40])
    elif kind == "gpc":
        items = sorted(BI.values(), key=lambda x: x["t"])
        if q:
            qq = q.lower()
            items = [x for x in items if qq in x["t"].lower() or qq in x["c"].lower()]
        kb, txt = PG(items, page, "gpc_", lambda x: "📦 " + x["t"][:40] + " - " + x["c"])
    elif kind == "rot":
        items = list(ROT) + list(ROT2)
        if q:
            ps = q.split()
            old_hits = []
            if len(ps) >= 2:
                old_hits = SROT(ps[0], ps[1])
            new_hits = SROT2(q)
            if old_hits or new_hits:
                items = old_hits + new_hits
            else:
                items = []
        kb, txt = PG(items, page, "rot_", lambda r: "🔄 " + str(r.get("pol",""))[:30] + " -> " + str(r.get("pod",""))[:25] + " | " + str(r.get("line",""))[:25])
    elif kind == "docs":
        ds = LJ("docs_index.json", [])
        if q:
            ds = [x for x in ds if NA(q) in NA(x["name"]) or str(q) in str(x["num"])]
        kb, txt = PG(ds, page, "docs_", lambda x: "📄 " + x["name"][:35] + " - " + str(x["num"]))
    else:
        items = SREF(kind, q)
        icons = {"cities": "🏙️ ", "country": "🌍 ", "currency": "💱 ", "package": "📦 ", "uom": "⚖️ "}
        ic = icons.get(kind, "")
        kb, txt = PG(items, page, "ref_" + kind + "_", lambda x: ic + x["name"][:35] + " - " + x["code"])
    ctx.user_data["lst"] = items
    if up.message:
        await up.message.reply_text(txt, reply_markup=kb)
    else:
        try:
            await up.callback_query.message.edit_text(txt, reply_markup=kb)
        except:
            await up.callback_query.message.reply_text(txt, reply_markup=kb)

async def SEND_ADMIN_REQUEST(ctx, us, code):
    if not ADM:
        return
    username = "@" + us.username if us.username else "مفيش يوزرنيم"
    full_name = (us.first_name or "") + " " + (us.last_name or "")
    full_name = full_name.strip() or "بدون اسم"
    txt = "🆕 طلب انضمام جديد\n\n"
    txt += f"👤 الاسم: {full_name}\n"
    txt += f"🔗 اليوزرنيم: {username}\n"
    txt += f"🆔 ID: {us.id}\n"
    txt += f"🎟️ الكود: {code}\n"
    txt += f"📊 الحالة: pending\n"
    txt += f"👆 اضغط للمراسلة: tg://user?id={us.id}\n"
    kb = MK([[B("✅ قبول", "accept_" + str(us.id)), B("❌ رفض", "reject_" + str(us.id))], [B("💰 ابعت تفاصيل الاشتراك", "subinfo_" + str(us.id))]])
    try:
        await ctx.bot.send_message(chat_id=int(ADM), text=txt, reply_markup=kb)
    except Exception as e:
        print(f"Admin notify fail: {e}")

async def START(update, ctx):
    try:
        if ctx.user_data is not None:
            ctx.user_data["end_sent"] = False
    except:
        pass
    us = update.message.from_user
    u = LU()
    uid = str(us.id)
    now = time.time()
    if uid == str(ADM):
        await update.message.reply_text(WEL + "\n\nانت الادمن 👑\nاكتب /users عشان تشوف اليوزرز", reply_markup=MAIN_KB())
        return
    if uid in u and u[uid].get("status") == "accepted":
        await update.message.reply_text(WEL, reply_markup=MAIN_KB())
        return
    if uid in u and u[uid].get("status") == "rejected":
        await update.message.reply_text("❌ تم رفض طلبك من قبل الادمن.\nتواصل مع الادمن لو شايف ان ده غلط.")
        return
    if uid in u and u[uid].get("status") == "pending":
        code = u[uid].get("code", "")
        await update.message.reply_text(f"⏳ طلبك لسه معلق عند الادمن\nالكود بتاعك: {code}\nاستنى الادمن يقبلك.")
        return
    code = "USR-" + "".join(random.choices("0123456789", k=4))
    d = {}
    d["code"] = code
    d["name"] = (us.first_name or "") + " " + (us.last_name or "")
    d["name"] = d["name"].strip()
    d["username"] = "@" + us.username if us.username else ""
    d["status"] = "pending"
    d["requested_at"] = now
    d["notified"] = True
    u[uid] = d
    SU(u)
    msg = WEL + "\n\n🎟️ الكود بتاعك: " + code + "\n⏳ استنى الادمن يقبلك الاول"
    await update.message.reply_text(msg)
    await SEND_ADMIN_REQUEST(ctx, us, code)

async def USERS(update, ctx):
    if str(update.message.from_user.id) != str(ADM):
        await update.message.reply_text("❌ الامر ده للادمن بس")
        return
    u = LU()
    if not u:
        await update.message.reply_text("مفيش يوزرز خالص لسه")
        return
    pending = {k: v for k, v in u.items() if v.get("status") == "pending"}
    others = {k: v for k, v in u.items() if v.get("status") != "pending"}
    if pending:
        await update.message.reply_text(f"⏳ عندك {len(pending)} طلب معلق:")
        for uid, d in pending.items():
            txt = "🆕 طلب انضمام\n\n"
            txt += f"👤 الاسم: {d.get('name','بدون اسم')}\n"
            txt += f"🔗 اليوزرنيم: {d.get('username','-')}\n"
            txt += f"🆔 ID: {uid}\n"
            txt += f"🎟️ الكود: {d.get('code','-')}\n"
            txt += f"📊 الحالة: pending\n"
            txt += f"👆 اضغط للمراسلة: tg://user?id={uid}\n"
            kb = MK([[B("✅ قبول", "accept_" + uid), B("❌ رفض", "reject_" + uid)], [B("💰 ابعت تفاصيل الاشتراك", "subinfo_" + uid)]])
            await update.message.reply_text(txt, reply_markup=kb)
    if others:
        await update.message.reply_text(f"👥 باقي اليوزرز ({len(others)}):")
        for uid, d in others.items():
            emoji = "✅" if d.get("status") == "accepted" else "❌"
            txt = f"{emoji} {d.get('name','بدون اسم')} | {d.get('username','-')}\n"
            txt += f"ID: {uid} | كود: {d.get('code','-')} | الحالة: {d.get('status')}"
            if d.get("status") == "accepted":
                kb = MK([[B("❌ رفض", "reject_" + uid)]])
                await update.message.reply_text(txt, reply_markup=kb)
            else:
                kb = MK([[B("✅ قبول", "accept_" + uid)]])
                await update.message.reply_text(txt, reply_markup=kb)
    if not pending and not others:
        await update.message.reply_text("مفيش يوزرز")

async def BTN(update, ctx):
    q = update.callback_query
    try:
        await q.answer()
    except:
        pass
    d = q.data
    uid = q.from_user.id
    RESET_END_TIMER(ctx)
    if d.startswith("subinfo_"):
        if str(uid) != str(ADM):
            await q.message.reply_text("انت مش الادمن")
            return
        uid2 = d.split("_")[1]
        try:
            await ctx.bot.send_message(chat_id=int(uid2), text=SUB_INFO_TXT)
            await q.message.reply_text("✅ اتبعتله رسالة الاشتراك (150 جنيه في الشهر)، استنى رده قبل ما تقبله")
        except:
            await q.message.reply_text("❌ مش عارف ابعتله، هو لسه مبعتش /start")
        return
    if d.startswith("accept_"):
        if str(uid) != str(ADM):
            await q.message.reply_text("انت مش الادمن")
            return
        uid2 = d.split("_")[1]
        u = LU()
        if uid2 in u:
            u[uid2]["status"] = "accepted"
            SU(u)
            try:
                await ctx.bot.send_message(chat_id=int(uid2), text=WEL, reply_markup=MAIN_KB())
            except:
                pass
            await q.message.reply_text(f"✅ تم قبول {u[uid2].get('name','')} - {uid2}")
            try:
                await q.message.edit_text(q.message.text + "\n\n✅ تم القبول")
            except:
                pass
        return
    if d.startswith("reject_"):
        if str(uid) != str(ADM):
            await q.message.reply_text("انت مش الادمن")
            return
        uid2 = d.split("_")[1]
        u = LU()
        if uid2 in u:
            u[uid2]["status"] = "rejected"
            SU(u)
            try:
                await ctx.bot.send_message(chat_id=int(uid2), text="❌ تم رفض طلبك من الادمن")
            except:
                pass
            await q.message.reply_text(f"❌ تم رفض {u[uid2].get('name','')} - {uid2}")
            try:
                await q.message.edit_text(q.message.text + "\n\n❌ تم الرفض")
            except:
                pass
        return
    if not OK(uid):
        await q.message.reply_text("⏳ استنى الادمن يقبلك الاول، ابعت /start")
        return
    if d == "docs_add":
        if str(uid) != str(ADM):
            await q.message.reply_text("❌ اضافة المستندات للادمن بس")
            return
        await q.message.reply_text("📄 ابعت الصورة أو الـ PDF واكتب عليها اسم المستند في الكابشن\nمثال: منشور 28")
        return
    if d == "docs_list":
        await q.message.reply_text(LIST_DOCS_TXT())
        return
    if d == "docs_del_all":
        if str(uid) != str(ADM):
            await q.message.reply_text("❌ للادمن بس")
            return
        n = DEL_ALL_DOCS()
        await q.message.reply_text(f"✅ اتمسح كل المستندات ({n} مستند)")
        return
    if d.startswith("docs_del_"):
        if str(uid) != str(ADM):
            await q.message.reply_text("❌ للادمن بس")
            return
        try:
            num = int(d.split("_")[2])
        except:
            return
        if DEL_ONE_DOC(num):
            await q.message.reply_text(f"✅ اتمسح المستند رقم {num}")
        else:
            await q.message.reply_text("مفيش مستند بالرقم ده")
        return
    if d == "m_main":
        ctx.user_data["calc_on"] = False
        ctx.user_data["cbm_on"] = False
        await q.message.reply_text("القايمة:", reply_markup=MAIN_KB())
        return
    if d == "m_calc":
        CALC_RESET(ctx)
        await q.message.reply_text("🧮 حسبة الرسوم التقريبية\n\nاكتب مبلغ الفاتورة بالعملة الخارجية:\nمثال: 67500", reply_markup=MK([[B("❌ الغاء", "m_main")]]))
        return
    if d.startswith("calc_ct_"):
        c = d.replace("calc_ct_", "")
        ctx.user_data["calc"]["contract"] = c
        ctx.user_data["calc"]["step"] = "duty"
        await q.message.reply_text(f"📦 التعاقد: {c}\n\nاكتب فئة البند / الجمارك %:\nمثال: 2  يعني 2%\nلو معفي اكتب 0")
        return
    if d.startswith("calc_cur_"):
        c = d.replace("calc_cur_", "")
        ctx.user_data["calc"]["cur"] = c
        ctx.user_data["calc"]["step"] = "contract"
        await q.message.reply_text(f"💱 العملة: {c}\n\nاختار التعاقد:", reply_markup=CALC_KB_CONTRACT())
        return
    if d == "m_cbm":
        CBM_RESET(ctx)
        await q.message.reply_text("📦 حسبة CBM\n\nاختار الوحدة:", reply_markup=CBM_KB_UNIT())
        return
    if d.startswith("cbm_unit_"):
        u = "cm" if d.endswith("_cm") else "in"
        ctx.user_data["cbm"]["unit"] = u
        ctx.user_data["cbm"]["step"] = "div"
        await q.message.reply_text("اختار نوع الشحن عشان الوزن الحجمي:", reply_markup=CBM_KB_DIV())
        return
    if d.startswith("cbm_div_"):
        div = int(d.replace("cbm_div_", ""))
        ctx.user_data["cbm"]["div"] = div
        ctx.user_data["cbm"]["step"] = "l"
        unit_l = "سم" if ctx.user_data["cbm"]["unit"] == "cm" else "بوصة"
        await q.message.reply_text(f"📐 دخل الطول بالـ {unit_l}:\nمثال: 60")
        return
    if d == "m_lines":
        await SEND(q, ctx, "line", 0, "")
        return
    if d == "m_comp":
        kb = COMP_KB()
        await q.message.reply_text("🏢 الشركات:", reply_markup=kb)
        return
    if d == "ct_pri":
        ctx.user_data["ct"] = "private"
        await SEND(q, ctx, "comp", 0, "")
        return
    if d == "ct_agent":
        ctx.user_data["ct"] = "agent"
        await SEND(q, ctx, "comp", 0, "")
        return
    if d == "ct_gov":
        ctx.user_data["ct"] = "gov"
        await SEND(q, ctx, "comp", 0, "")
        return
    if d == "ct_for":
        ctx.user_data["ct"] = "foreign"
        await SEND(q, ctx, "comp", 0, "")
        return
    if d == "m_tariff":
        await SEND(q, ctx, "tariff", 0, "")
        return
    if d == "m_gpc":
        await SEND(q, ctx, "gpc", 0, "")
        return
    if d == "m_rot":
        await SEND(q, ctx, "rot", 0, "")
        return
    if d == "m_country":
        await SEND(q, ctx, "country", 0, "")
        return
    if d == "m_cities":
        await SEND(q, ctx, "cities", 0, "")
        return
    if d == "m_cur":
        await SEND(q, ctx, "currency", 0, "")
        return
    if d == "m_pack":
        await SEND(q, ctx, "package", 0, "")
        return
    if d == "m_uom":
        await SEND(q, ctx, "uom", 0, "")
        return
    if d == "m_help":
        await q.message.reply_text(HELP, reply_markup=MAIN_KB())
        return
    if d == "m_recent":
        lst = ctx.user_data.get("recent", [])
        if not lst:
            await q.message.reply_text("لسه مفيش بحث")
            return
        t = "اخر بحث:\n"
        for x in lst:
            t += x + "\n"
        await q.message.reply_text(t)
        return
    if d == "m_docs":
        await SEND(q, ctx, "docs", 0, "")
        return
    if d.startswith("pick_"):
        kind = d.split("_")[1]
        qq = ctx.user_data.get("multi_q", "")
        ctx.user_data["multi_q"] = ""
        if kind == "comp":
            await SEND(q, ctx, "comp", 0, qq)
        elif kind == "tariff":
            await SEND(q, ctx, "tariff", 0, qq)
        elif kind == "line":
            await SEND(q, ctx, "line", 0, qq)
        elif kind == "gpc":
            await SEND(q, ctx, "gpc", 0, qq)
        elif kind == "rot":
            await SEND(q, ctx, "rot", 0, qq)
        return
    if d.startswith("comp_page_"):
        p = int(d.split("_")[2])
        await SEND(q, ctx, "comp", p, ctx.user_data.get("fq", ""))
        return
    if d.startswith("comp_sel_"):
        i = int(d.split("_")[2])
        lst = ctx.user_data.get("lst", [])
        if 0 <= i < len(lst):
            await q.message.reply_text(FC(lst[i]))
        return
    if d.startswith("tariff_page_"):
        p = int(d.split("_")[2])
        await SEND(q, ctx, "tariff", p, ctx.user_data.get("fq", ""))
        return
    if d.startswith("tariff_sel_"):
        i = int(d.split("_")[2])
        lst = ctx.user_data.get("lst", [])
        if 0 <= i < len(lst):
            await q.message.reply_text(FT(lst[i]))
        return
    if d.startswith("line_page_"):
        p = int(d.split("_")[2])
        await SEND(q, ctx, "line", p, ctx.user_data.get("fq", ""))
        return
    if d.startswith("line_sel_"):
        i = int(d.split("_")[2])
        lst = ctx.user_data.get("lst", [])
        if 0 <= i < len(lst):
            k = lst[i]
            await q.message.reply_text(FL(k, "ALL"), reply_markup=LCAT_KB(k))
        return
    if d.startswith("gpc_page_"):
        p = int(d.split("_")[2])
        await SEND(q, ctx, "gpc", p, ctx.user_data.get("fq", ""))
        return
    if d.startswith("gpc_sel_"):
        i = int(d.split("_")[2])
        lst = ctx.user_data.get("lst", [])
        if 0 <= i < len(lst):
            await q.message.reply_text(FB(lst[i]))
        return
    if d.startswith("rot_page_"):
        p = int(d.split("_")[2])
        await SEND(q, ctx, "rot", p, ctx.user_data.get("fq", ""))
        return
    if d.startswith("rot_sel_"):
        i = int(d.split("_")[2])
        lst = ctx.user_data.get("lst", [])
        if 0 <= i < len(lst):
            r = lst[i]
            await q.message.reply_text(FROT_MIXED(r))
        return
    if d.startswith("docs_page_"):
        p = int(d.split("_")[2])
        await SEND(q, ctx, "docs", p, ctx.user_data.get("fq", ""))
        return
    if d.startswith("docs_sel_"):
        i = int(d.split("_")[2])
        lst = ctx.user_data.get("lst", [])
        if 0 <= i < len(lst):
            x = lst[i]
            try:
                if str(uid) == str(ADM):
                    kb = MK([[B("🗑️ حذف المستند ده", "docs_del_" + str(x["num"]))], [B("🏠 القايمة الرئيسية", "m_main")]])
                    await q.message.reply_document(document=open(x["path"], "rb"), caption=x["name"], reply_markup=kb)
                else:
                    await q.message.reply_document(document=open(x["path"], "rb"), caption=x["name"])
            except:
                await q.message.reply_text("مفيش مستند")
        return
    if d.startswith("ref_"):
        ps = d.split("_")
        kind = ps[1]
        if ps[2] == "page":
            p = int(ps[3])
            await SEND(q, ctx, kind, p, ctx.user_data.get("fq", ""))
            return
        if ps[2] == "sel":
            i = int(ps[3])
            lst = ctx.user_data.get("lst", [])
            if 0 <= i < len(lst):
                it = lst[i]
                t = FREF(kind, it)
                await q.message.reply_text(t)
            return
    if d.startswith("lc_"):
        ps = d.split("_")
        k = ps[1]
        cat = ps[2]
        await q.message.reply_text(FL(k, cat), reply_markup=LCAT_KB(k))
        return

async def SAVE_DOC(update, ctx):
    if str(update.message.from_user.id) != str(ADM):
        await update.message.reply_text("❌ اضافة المستندات للادمن بس")
        return
    cap = (update.message.caption or "").strip()
    if cap in ["صورة النهاية", "end", "end_image", "تغيير صورة النهاية"]:
        f2 = None
        if update.message.photo:
            f2 = await update.message.photo[-1].get_file()
        if f2:
            await f2.download_to_drive(END_IMG)
            await update.message.reply_text("✅ اتغيرت صورة النهاية بنجاح")
        return
    if not cap:
        cap = "بدون اسم"
    f = None
    ext = "jpg"
    if update.message.photo:
        f = await update.message.photo[-1].get_file()
        ext = "jpg"
    elif update.message.document:
        f = await update.message.document.get_file()
        fname = update.message.document.file_name or ""
        if "." in fname:
            ext = fname.split(".")[-1].lower()
        else:
            ext = "pdf"
    else:
        return
    ds = LJ("docs_index.json", [])
    num = len(ds) + 1
    used = [x["num"] for x in ds]
    for i in range(1, len(ds) + 2):
        if i not in used:
            num = i
            break
    fn = "docs/" + str(num) + "." + ext
    await f.download_to_drive(fn)
    d = {}
    d["num"] = num
    d["name"] = cap
    d["path"] = fn
    ds.append(d)
    ds = sorted(ds, key=lambda x: x["num"])
    SJ("docs_index.json", ds)
    await update.message.reply_text(f"✅ اتحفظ المستند رقم {num}\n📄 الاسم: {cap}\nاكتب: تزيل {num} عشان تنزله")

async def HANDLE(update, ctx):
    us = update.message.from_user
    uid = str(us.id)
    text = update.message.text or ""
    low = text.strip().lower()
    bare = text.strip().lstrip(SL).strip().lower()
    RESET_END_TIMER(ctx)
    if not OK(uid):
        await update.message.reply_text("⏳ استنى الادمن يقبلك الاول، ابعت /start")
        return
    if ctx.user_data.get("cbm_on"):
        cbm = ctx.user_data.get("cbm", {})
        step = cbm.get("step", "unit")
        txt2 = text.strip().replace(",", "").replace("جنيه", "").replace("ج", "")
        if NA(text) in ["الغاء", "cancel", "menu", "القايمه", "القايمة"]:
            ctx.user_data["cbm_on"] = False
            await update.message.reply_text("اتلغت حسبة CBM", reply_markup=MAIN_KB())
            return
        if step in ["unit", "div"]:
            await update.message.reply_text("اختار من الزراير اللي فوق 👆", reply_markup=CBM_KB_UNIT() if step == "unit" else CBM_KB_DIV())
            return
        try:
            val = float(txt2)
        except:
            val = None
        unit_l = "سم" if cbm.get("unit") == "cm" else "بوصة"
        unit_w = "كجم" if cbm.get("unit") == "cm" else "رطل"
        if step == "l":
            if val is None or val <= 0:
                await update.message.reply_text(f"اكتب الطول صح بالـ {unit_l}، مثال: 60")
                return
            cbm["l"] = val; cbm["step"] = "w"
            await update.message.reply_text(f"📐 الطول: {val}\n\nدخل العرض بالـ {unit_l}:")
            return
        if step == "w":
            if val is None or val <= 0:
                await update.message.reply_text(f"اكتب العرض صح بالـ {unit_l}")
                return
            cbm["w"] = val; cbm["step"] = "h"
            await update.message.reply_text(f"📐 العرض: {val}\n\nدخل الارتفاع بالـ {unit_l}:")
            return
        if step == "h":
            if val is None or val <= 0:
                await update.message.reply_text(f"اكتب الارتفاع صح بالـ {unit_l}")
                return
            cbm["h"] = val; cbm["step"] = "wt"
            await update.message.reply_text(f"📐 الارتفاع: {val}\n\nدخل وزن الكرتونة الواحدة بالـ {unit_w}:")
            return
        if step == "wt":
            if val is None or val <= 0:
                await update.message.reply_text(f"اكتب وزن الكرتونة صح بالـ {unit_w}")
                return
            cbm["wt"] = val; cbm["step"] = "qty"
            await update.message.reply_text(f"⚖️ الوزن: {val} {unit_w}\n\nدخل عدد الكراتين:")
            return
        if step == "qty":
            if val is None or val <= 0:
                await update.message.reply_text("اكتب عدد الكراتين صح، مثال: 100")
                return
            cbm["qty"] = int(val)
            ctx.user_data["cbm_on"] = False
            r = CBM_DO(cbm)
            await update.message.reply_text(CBM_TXT(cbm, r), reply_markup=MAIN_KB())
            return
    if ctx.user_data.get("calc_on"):
        calc = ctx.user_data.get("calc", {})
        step = calc.get("step", "inv")
        txt = text.strip().replace(",", "").replace("جنيه", "").replace("ج", "")
        if NA(text) in ["الغاء", "cancel", "menu", "القايمه", "القايمة"]:
            ctx.user_data["calc_on"] = False
            await update.message.reply_text("اتلغت الحسبة", reply_markup=MAIN_KB())
            return
        try:
            val = float(txt)
        except:
            val = None
        if step == "inv":
            if val is None or val <= 0:
                await update.message.reply_text("اكتب مبلغ الفاتورة صح، مثال: 67500")
                return
            calc["inv"] = val; calc["step"] = "cur"
            await update.message.reply_text(f"💰 الفاتورة: {val:,.2f}\n\nاختار نوع العملة:", reply_markup=CALC_KB_CUR())
            return
        if step == "duty":
            if val is None or val < 0 or val > 100:
                await update.message.reply_text("اكتب فئة البند % صح، مثال: 2 او 0 او 10")
                return
            calc["duty"] = val; calc["step"] = "rate"
            await update.message.reply_text(f"📜 فئة البند: {val}%\n\nاكتب سعر التحويل:\nمثال: 47.38")
            return
        if step == "rate":
            if val is None or val <= 0:
                await update.message.reply_text("اكتب سعر التحويل صح، مثال: 47.38")
                return
            calc["rate"] = val; calc["step"] = "other"
            await update.message.reply_text(f"💱 سعر التحويل: {val}\n\nاكتب مصاريف أخرى بالمصري:\nلو مفيش اكتب 0")
            return
        if step == "other":
            if val is None or val < 0:
                await update.message.reply_text("اكتب مصاريف أخرى صح، او 0")
                return
            calc["other"] = val
            ctx.user_data["calc_on"] = False
            r = CALC_DO(calc)
            await update.message.reply_text(CALC_TXT(calc, r), reply_markup=MAIN_KB())
            return
    if NA(text) in [NA(w) for w in END_WORDS]:
        await update.message.reply_text("العفو، نورتنا ❤️")
        await SEND_END_IMAGE(ctx, update.message.chat_id)
        return
    if bare in ["اضافة مستند", "اضف مستند", "إضافة مستند", "اضافة مستندات"]:
        if str(uid) != str(ADM):
            await update.message.reply_text("❌ اضافة المستندات للادمن بس")
            return
        await update.message.reply_text("📄 ابعت الصورة أو الـ PDF واكتب عليها اسم المستند في الكابشن\nمثال: منشور 28", reply_markup=DOCS_ADMIN_KB())
        return
    if bare in ["ادارة المستندات", "إدارة المستندات", "مستندات الادمن"]:
        if str(uid) != str(ADM):
            await update.message.reply_text("❌ للادمن بس")
            return
        await update.message.reply_text("📄 ادارة المستندات:", reply_markup=DOCS_ADMIN_KB())
        return
    if bare in ["قائمة المستندات", "قايمة المستندات", "عرض المستندات", "list docs"]:
        await update.message.reply_text(LIST_DOCS_TXT())
        return
    if text.strip().startswith("حذف مستند") or text.strip().startswith("حذف تزيل") or text.strip().startswith("مسح مستند") or text.strip().startswith("حذف المستند"):
        if str(uid) != str(ADM):
            await update.message.reply_text("❌ حذف المستندات للادمن بس")
            return
        m = re.search(r"(\d+)", text)
        if not m:
            await update.message.reply_text("اكتب: حذف مستند 1")
            return
        num = int(m.group(1))
        if DEL_ONE_DOC(num):
            await update.message.reply_text(f"✅ اتمسح المستند رقم {num}")
        else:
            await update.message.reply_text("مفيش مستند بالرقم ده")
        return
    if text.strip() in ["حذف كل المستندات", "مسح كل المستندات", "حذف جميع المستندات"]:
        if str(uid) != str(ADM):
            await update.message.reply_text("❌ للادمن بس")
            return
        n = DEL_ALL_DOCS()
        await update.message.reply_text(f"✅ اتمسح كل المستندات ({n} مستند) - تقدر تبدأ ترفع من جديد")
        return
    if "تزيل" in text or "download" in low:
        SAVE_RECENT(ctx, text)
        m = re.search(r"(\d+)", text)
        if not m:
            await update.message.reply_text("اكتب: تزيل 1 او تزيل اسم المستند 1")
            return
        num = int(m.group(1))
        name_part = re.sub(r"تزيل|download|\d+", "", text, flags=re.I).strip()
        for x in LJ("docs_index.json", []):
            if x["num"] == num:
                if name_part and NA(name_part) not in NA(x["name"]):
                    continue
                cap = x["name"]
                await update.message.reply_document(document=open(x["path"], "rb"), caption=cap)
                return
        await update.message.reply_text("مفيش مستند")
        return
    if bare == "شركات":
        await update.message.reply_text("🏢 الشركات:", reply_markup=COMP_KB())
        return
    if bare == "خاصة" or bare == "private":
        ctx.user_data["ct"] = "private"
        await SEND(update, ctx, "comp", 0, "")
        return
    if bare in ["وكلاء الشحن", "وكلاء", "وكيل شحن", "agent", "agents"]:
        ctx.user_data["ct"] = "agent"
        await SEND(update, ctx, "comp", 0, "")
        return
    if bare == "حكومية" or bare == "gov":
        ctx.user_data["ct"] = "gov"
        await SEND(update, ctx, "comp", 0, "")
        return
    if bare == "اجنبية" or bare == "foreign":
        ctx.user_data["ct"] = "foreign"
        await SEND(update, ctx, "comp", 0, "")
        return
    if bare == "menu" or bare == "القايمة":
        await update.message.reply_text("القايمة:", reply_markup=MAIN_KB())
        return
    if bare == "تعريفة" or bare == "tariff":
        await SEND(update, ctx, "tariff", 0, "")
        return
    if bare == "توكيلات":
        await SEND(update, ctx, "line", 0, "")
        return
    if bare in ["روتيشن", "rotation", "rot", "الروتيشن", "الروتيشنات"]:
        await SEND(update, ctx, "rot", 0, "")
        return
    if low.startswith("rotation ") or low.startswith("rot ") or text.strip().startswith("روتيشن "):
        SAVE_RECENT(ctx, text)
        qq = text.strip().split(" ", 1)[1] if " " in text.strip() else ""
        if not qq:
            await SEND(update, ctx, "rot", 0, "")
            return
        await SEND(update, ctx, "rot", 0, qq)
        return
    if low.startswith("gpc ") or low.startswith("gpc-") or bare == "gpc":
        SAVE_RECENT(ctx, text)
        qq = text.strip().split(" ", 1)[1] if " " in text.strip() else ""
        if not qq:
            await SEND(update, ctx, "gpc", 0, "")
            return
        b = SGPC(qq)
        if b is None:
            await update.message.reply_text("📦 ملقيتش صنف")
            return
        if isinstance(b, tuple):
            ctx.user_data["lst"] = b[1]
            rows = []
            for i, x in enumerate(b[1]):
                rows.append([B("📦 " + x["t"][:50], "gpc_sel_" + str(i))])
            kb = MK(rows)
            await update.message.reply_text("اختار صنف:", reply_markup=kb)
            return
        await update.message.reply_text(FB(b))
        return
    if low.startswith("package ") or low.startswith("pack ") or bare == "package":
        SAVE_RECENT(ctx, text)
        qq = text.strip().split(" ", 1)[1] if " " in text.strip() else ""
        items = SREF("package", qq)
        if not items:
            await update.message.reply_text("📦 ملقيتش عبوة")
            return
        if len(items) == 1:
            await update.message.reply_text(FREF("package", items[0]))
            return
        await SEND(update, ctx, "package", 0, qq)
        return
    if low.startswith("currency ") or low.startswith("curr ") or bare == "currency":
        SAVE_RECENT(ctx, text)
        qq = text.strip().split(" ", 1)[1] if " " in text.strip() else ""
        items = SREF("currency", qq)
        if not items:
            await update.message.reply_text("💱 ملقيتش عملة")
            return
        if len(items) == 1:
            await update.message.reply_text(FREF("currency", items[0]))
            return
        await SEND(update, ctx, "currency", 0, qq)
        return
    if low.startswith("uom ") or bare == "uom":
        SAVE_RECENT(ctx, text)
        qq = text.strip().split(" ", 1)[1] if " " in text.strip() else ""
        items = SREF("uom", qq)
        if not items:
            await update.message.reply_text("⚖️ ملقيتش وحدة")
            return
        if len(items) == 1:
            await update.message.reply_text(FREF("uom", items[0]))
            return
        await SEND(update, ctx, "uom", 0, qq)
        return
    if low.startswith("country ") or bare == "country":
        SAVE_RECENT(ctx, text)
        qq = text.strip().split(" ", 1)[1] if " " in text.strip() else ""
        items = SREF("country", qq)
        if not items:
            await update.message.reply_text("🌍 ملقيتش دولة")
            return
        if len(items) == 1:
            await update.message.reply_text(FREF("country", items[0]))
            return
        await SEND(update, ctx, "country", 0, qq)
        return
    if low.startswith("city ") or low.startswith("cities ") or low.startswith("citys ") or bare in ["city", "cities"]:
        SAVE_RECENT(ctx, text)
        qq = text.strip().split(" ", 1)[1] if " " in text.strip() else ""
        items = SREF("cities", qq)
        if not items:
            await update.message.reply_text("🏙️ ملقيتش مدينة")
            return
        if len(items) == 1:
            await update.message.reply_text(FREF("cities", items[0]))
            return
        await SEND(update, ctx, "cities", 0, qq)
        return
    if re.match(r"^[\d\.\s\-]+$", text.strip()):
        SAVE_RECENT(ctx, text)
        qd = text.strip().replace(" ", "").replace("-", "").replace(".", "")
        if qd in CBR:
            await update.message.reply_text(FC(CBR[qd]))
            return
        cand = [c for c in CL if qd in c["reg"] or c["reg"] in qd]
        if cand:
            if len(cand) == 1:
                await update.message.reply_text(FC(cand[0]))
                return
            await SEND(update, ctx, "comp", 0, text.strip())
            return
        rs = ST(text.strip())
        if rs:
            if len(rs) == 1:
                await update.message.reply_text(FT(rs[0]))
                return
            await SEND(update, ctx, "tariff", 0, text.strip())
            return
        await update.message.reply_text("📜 ملقيتش بند / شركة")
        return
    ps = text.strip().split()
    if len(ps) == 2:
        rs = SROT(ps[0], ps[1])
        rs2 = SROT2(text.strip())
        if rs or rs2:
            SAVE_RECENT(ctx, text)
            if rs:
                await update.message.reply_text(FROT(rs, ps[0], ps[1]))
            if rs2:
                await SEND_ALL_ROT2(update, rs2, text.strip())
            return
    k_line = RL(text.strip())
    comp_hits = GC(text.strip(), "")
    tariff_hits = ST(text.strip())
    gpc_hit = SGPC(text)
    rot2_hits = SROT2(text.strip())
    found = []
    if k_line: found.append("line")
    if comp_hits: found.append("comp")
    if tariff_hits: found.append("tariff")
    if gpc_hit is not None: found.append("gpc")
    if rot2_hits: found.append("rot")
    if len(found) >= 2:
        SAVE_RECENT(ctx, text)
        ctx.user_data["multi_q"] = text.strip()
        names = {"line": "🚢 التوكيلات", "comp": "🏢 الشركات", "tariff": "📜 التعريفة", "gpc": "📦 GPC", "rot": "🔄 Rotation"}
        t = "بحثك `" + text.strip() + "` وجد في اكتر من ايقونة:\n\n"
        rows = []
        for f in found:
            t += "- " + names[f] + "\n"
            rows.append([B(names[f], "pick_" + f)])
        t += "\nماذا تختار منهم؟ دوس على اللي انت عايزه"
        await update.message.reply_text(t, reply_markup=MK(rows))
        return
    if k_line:
        SAVE_RECENT(ctx, text)
        await update.message.reply_text(FL(k_line, "ALL"), reply_markup=LCAT_KB(k_line))
        return
    if rot2_hits:
        SAVE_RECENT(ctx, text)
        await SEND_ALL_ROT2(update, rot2_hits, text.strip())
        return
    if text.strip().isdigit() and len(text.strip()) >= 4:
        if text.strip() in CBR:
            SAVE_RECENT(ctx, text)
            await update.message.reply_text(FC(CBR[text.strip()]))
            return
        cand = [c for c in CL if text.strip() in c["reg"]]
        if len(cand) == 1:
            SAVE_RECENT(ctx, text)
            await update.message.reply_text(FC(cand[0]))
            return
        if cand:
            SAVE_RECENT(ctx, text)
            await SEND(update, ctx, "comp", 0, text.strip())
            return
    if comp_hits:
        SAVE_RECENT(ctx, text)
        await SEND(update, ctx, "comp", 0, text.strip())
        return
    if tariff_hits:
        SAVE_RECENT(ctx, text)
        if len(tariff_hits) == 1:
            await update.message.reply_text(FT(tariff_hits[0]))
            return
        await SEND(update, ctx, "tariff", 0, text.strip())
        return
    if gpc_hit is not None:
        SAVE_RECENT(ctx, text)
        if isinstance(gpc_hit, tuple):
            ctx.user_data["lst"] = gpc_hit[1]
            rows = []
            for i, x in enumerate(gpc_hit[1]):
                rows.append([B("📦 " + x["t"][:50], "gpc_sel_" + str(i))])
            kb = MK(rows)
            await update.message.reply_text("اختار صنف:", reply_markup=kb)
            return
        await update.message.reply_text(FB(gpc_hit))
        return
    await update.message.reply_text("وصلت", reply_markup=MAIN_KB())

def MAIN():
    print("Bot starting...")
    if not BOT:
        print("ERROR: No BOT_TOKEN! حط BOT_TOKEN في Environment Variables")
        return
    print(f"GPC: {len(BI)} | Tariff: {len(TL)} | Companies: {len(CL)} | Lines: {len(LN)} | ROT old: {len(ROT)} | ROT2 new: {len(ROT2)}")
    app = ApplicationBuilder().token(BOT).build()
    app.add_handler(CommandHandler("start", START))
    app.add_handler(CommandHandler(["users", "user", "USERS"], USERS))
    app.add_handler(CallbackQueryHandler(BTN))
    app.add_handler(MessageHandler(filters.PHOTO | filters.Document.ALL, SAVE_DOC))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, HANDLE))
    app.run_polling()
    print("Bot stopped")
MAIN()
