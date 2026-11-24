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

# ===== التوكن =====
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
    r7 = [B("❓ مساعدة", "m_help")]
    return MK([r1, r2, r3, r4, r5, r6, r7])
def COMP_KB():
    r1 = [B("🏢 الخاصة", "ct_pri")]
    r2 = [B("🏛️ الحكومية", "ct_gov")]
    r3 = [B("🌍 الاجنبية", "ct_for")]
    r4 = [B("🏠 القايمة الرئيسية", "m_main")]
    return MK([r1, r2, r3, r4])
def LCAT_KB(k):
    r1 = [B("📩 Booking", "lc_" + k + "_BK")]
    r1.append(B("🧾 Invoice", "lc_" + k + "_IN"))
    r2 = [B("📧 الايميلات", "lc_" + k + "_AL")]
    r3 = [B("🏠 القايمة الرئيسية", "m_main")]
    return MK([r1, r2, r3])

# ============================================================
# GPC
# ============================================================
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

# ============================================================
# المراجع - cities.xlsx
# ============================================================
REF = {}
for kind in ["cities", "country", "currency", "package", "uom"]:
    REF[kind] = {}
if os.path.exists("cities.xlsx"):
    wb = openpyxl.load_workbook("cities.xlsx", read_only=True)
    for kind, sh in [("cities", "City List"), ("country", "Country List"), ("currency", "Currency List"), ("package", "Package Type list"), ("uom", "UOM List")]:
        if sh not in wb.sheetnames:
            continue
        for r in wb[sh].iter_rows(values_only=True):
            if not r or not GV(r, 0) or not GV(r, 1):
                continue
            if kind == "currency":
                c = str(GV(r, 1)).strip().upper()
                n = str(GV(r, 0)).strip()
            else:
                c = str(GV(r, 0)).strip().upper()
                n = str(GV(r, 1)).strip()
            if c in REF[kind]:
                continue
            REF[kind][c] = {"code": c, "name": n}
    wb.close()
    print(f"REF loaded cities.xlsx")
else:
    print("REF FILE MISSING: cities.xlsx")

# ============================================================
# الشركات
# ============================================================
CL = []
CBR = {}
def AC(n, f, rg, en, ac, st, ad, ct):
    n = CC(n)
    rg = CC(rg).replace(" ", "")
    if not n or not rg:
        return
    if rg in CBR:
        return
    if "قطاع" in n and "حكومي" in n:
        return
    if "الاسم" in n:
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

F2 = "ارقام_تسجيل_الشركات_2026.xlsx"
if os.path.exists(F2):
    wb = openpyxl.load_workbook(F2, read_only=True)
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
    print(f"Companies loaded: {len(CL)}")
else:
    print(f"COMPANIES FILE MISSING: {F2}")

def FC(c):
    t = "🏢 " + c["name"] + "\n"
    t += "رقم التسجيل: " + c["reg"] + "\n"
    if c["file"]:
        t += "رقم الملف: " + c["file"] + "\n"
    if c["ent"]:
        t += "الكيان: " + c["ent"] + "\n"
    if c["act"]:
        t += "النشاط: " + c["act"] + "\n"
    if c["ad"]:
        t += "العنوان: " + c["ad"] + "\n"
    return t.strip()

def GC(q, ct):
    out = [c for c in CL if not ct or c["ct"] == ct]
    nq = NA(q)
    if not nq:
        return out
    words = [w for w in nq.split() if len(w) >= 2 and w not in ["شركه", "شركة", "الشركه"]]
    o = []
    for c in out:
        nn = NA(c["name"])
        na = NA(c["act"])
        reg = str(c["reg"]).replace(" ", "")
        if nq.replace(" ", "") in reg or reg in nq.replace(" ", ""):
            o.append(c)
            continue
        if words:
            if any(w in nn or w in na for w in words):
                o.append(c)
        else:
            if nq in nn or nq in na:
                o.append(c)
    return o

# ============================================================
# التعريفة
# ============================================================
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

# ============================================================
# التوكيلات + Rotation - القديم زي ما هو بالظبط
# ============================================================
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
    return t

# ============================================================
# ===== اضافة جديدة فقط: الروتيشن الجديد - عربي + انجليزي =====
# الملف الجديد: egypt_shipping_lines_unified_single_sheet1111111111.xlsx
# ده اضافة، القديم اللي فوق شغال زي ما هو، مفيش حاجة اتمسحت
# ============================================================
ROT2 = []
# خريطة عربي <-> انجليزي للمواني والتوكيلات - للبحث باللغتين
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
    "فلسطين": "palestine",
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
}
def NORM_ROT_Q(q):
    # توحيد البحث: عربي وانجليزي، وازالة الهمزات والتشكيل
    s = NA(q)
    s = s.replace("-", " ").replace("_", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s
def EXPAND_ROT_Q(q):
    # يرجع كل الصيغ الممكنة للكلمة: عربي + انجليزي
    nq = NORM_ROT_Q(q)
    outs = [nq]
    for k, v in list(PORT_AR_EN.items()) + list(LINE_AR_EN.items()):
        if NA(k) == nq or NA(v) == nq:
            outs.append(NA(k))
            outs.append(NA(v))
    # لو الكلمة فيها اجزاء، وسع كل جزء
    parts = nq.split()
    if len(parts) > 1:
        for p in parts:
            outs.extend(EXPAND_ROT_Q(p))
    return list(set([o for o in outs if o]))
def ROT_MATCH(text, q):
    # هل النص فيه الكلمة بأي لغة؟
    if not q or not text:
        return False
    nt = NORM_ROT_Q(text)
    for cand in EXPAND_ROT_Q(q):
        if cand and (cand in nt or nt in cand):
            return True
    return False
def LOAD_ROT2_FILE():
    # بيدور على ملف الروتيشن الجديد بأي اسم قريب منه
    cands = ["egypt_shipping_lines_unified_single_sheet1111111111.xlsx"]
    for f in sorted(os.listdir(".")):
        low = f.lower()
        if "egypt_shipping" in low and low.endswith(".xlsx") and f not in cands:
            cands.append(f)
        if "shipping_lines" in low and low.endswith(".xlsx") and f not in cands:
            cands.append(f)
    for fn in cands:
        if not os.path.exists(fn):
            continue
        try:
            wb = openpyxl.load_workbook(fn, read_only=True)
            ws = wb[wb.sheetnames[0]]
            cnt = 0
            for r in ws.iter_rows(values_only=True):
                if not r:
                    continue
                # الصفوف الاولى عناوين وملخصات - بنتخطاها
                c0 = CC(GV(r, 0))
                if not c0 or "التوكيل" in c0 or "القاعدة" in c0 or "جميع التوكيلات" in c0 or "عدد الخطوط" in c0 or "Shipping Line" in c0:
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
                d = {}
                # نفس مفاتيح القديم pol/pod/line/tt/co عشان SEND والازرار تشتغل من غير تغيير
                d["pol"] = egy_port or line
                d["pod"] = rotation
                d["line"] = (line + " - " + service).strip(" -")
                d["tt"] = (str(avg) + " يوم") if avg else rng
                d["co"] = freq or lane
                # بيانات اضافية للعرض الكامل
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
                ROT2.append(d)
                cnt += 1
            wb.close()
            print(f"ROT2 loaded {fn} -> {cnt}, total {len(ROT2)}")
            return
        except Exception as e:
            print(f"ROT2 load fail {fn}: {e}")
    print("ROT2 FILE MISSING: egypt_shipping_lines_unified_single_sheet1111111111.xlsx")
LOAD_ROT2_FILE()
def SROT2(q):
    # بحث في الملف الجديد بكلمة واحدة او كلمتين: توكيل / ميناء / خدمة / منطقة / روتيشن
    q = (q or "").strip()
    if not q or not ROT2:
        return []
    words = [w for w in NORM_ROT_Q(q).split() if len(w) >= 2]
    # كلمتين زي Alex Izmir او الاسكندرية ازمير: لازم الكلمتين يظهروا في نفس السطر
    if len(words) >= 2:
        out = []
        for r in ROT2:
            hay = " ".join([r.get("line",""), r.get("agent",""), r.get("service",""), r.get("egy_port",""), r.get("lane",""), r.get("region",""), r.get("rotation","")])
            if all(ROT_MATCH(hay, w) for w in words):
                out.append(r)
        return out[:20]
    # كلمة واحدة
    out = []
    for r in ROT2:
        hay = " ".join([r.get("line",""), r.get("agent",""), r.get("service",""), r.get("egy_port",""), r.get("lane",""), r.get("region",""), r.get("rotation","")])
        if ROT_MATCH(hay, q):
            out.append(r)
    return out[:20]
def FROT2(rs, q=""):
    if not rs:
        return "🔄 ملقيتش روتيشن لـ " + (q or "")
    t = "🔄 الروتيشن" + (" لـ " + q if q else "") + ":\n\n"
    for r in rs[:10]:
        t += "🚢 " + r.get("line","") + "\n"
        if r.get("agent"): t += "الوكيل: " + r["agent"] + "\n"
        if r.get("egy_port"): t += "ميناء التحميل في مصر: " + r["egy_port"] + "\n"
        if r.get("lane"): t += "خط التجارة: " + r["lane"] + "\n"
        if r.get("rotation"): t += "خط السير: " + r["rotation"] + "\n"
        if r.get("avg"): t += "متوسط الرحلة: " + str(r["avg"]) + " يوم\n"
        if r.get("rng"): t += "النطاق: " + r["rng"] + "\n"
        if r.get("freq"): t += "معدل الإبحار: " + r["freq"] + "\n"
        t += "\n"
    return t.strip()
def FROT_MIXED(r):
    # عرض سطر واحد سواء كان قديم او جديد - عشان rot_sel_ يشتغل للاتنين
    if r.get("rot2"):
        return FROT2([r], r.get("line",""))
    return FROT([r], r.get("pol",""), r.get("pod",""))

WEL = "Welcome to Captain Emad Customs Bot"
HELP = """❓ المساعدة - كل أيقونة بتعمل ايه:

🚢 توكيلات: اكتب MAERSK او ميرسك
📦 GPC: اكتب GPC Airbrushes او GPC + الاسم او الكود
📜 التعريفة: اكتب 8420 او 84 او اسم الصنف - مثال: مروحة - او 4802 / 48.02
🏢 شركات: اكتب النور او رقم التسجيل
🔄 Rotation: اكتب Alex Izmir او الاسكندرية ازمير او MSC او ميرسك او Sokhna او السخنة
🌍 Country: اكتب country Egypt
🏙️ Cities: اكتب city Cairo
💱 Currency: اكتب currency USD
📦 Package: اكتب package Box
⚖️ UOM: اكتب uom KG
📄 مستندات: اكتب تزيل 1
🕐 اخر بحث: بيجيب اخر حاجة بحثت عنها
القايمة: اكتب menu
"""

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
    out = []
    for x in items:
        if nq in NA(x["name"]) or nq in NA(x["code"]) or q.upper() in x["code"] or nq in x["code"].lower():
            out.append(x)
    return out

async def SEND_END_IMAGE(ctx, chat_id):
    if ctx.user_data.get("end_sent"):
        return
    ctx.user_data["end_sent"] = True
    if os.path.exists(END_IMG):
        try:
            await ctx.bot.send_photo(chat_id=chat_id, photo=open(END_IMG, "rb"), caption="لو حابب تكمل أو تشترك، دي بيانات التواصل بتاعتنا ❤️")
            return
        except:
            pass
    await ctx.bot.send_message(chat_id=chat_id, text="لو حابب تكمل أو تشترك:\nTelegram: https://t.me/tarrifabot\nWhatsApp: wa.me/+201094038730\nPhone: 01094038730")

async def END_JOB(ctx):
    chat_id = ctx.job.chat_id
    await SEND_END_IMAGE(ctx, chat_id)

def RESET_END_TIMER(ctx):
    try:
        if ctx.job_queue is None:
            return
        for j in ctx.job_queue.jobs():
            if j.name and j.name.startswith("end_timer_"):
                try:
                    if j.chat_id == ctx._chat_id:
                        j.schedule_removal()
                except:
                    pass
    except:
        pass
    try:
        chat_id = None
        try:
            chat_id = ctx._chat_id
        except:
            pass
        if chat_id is None:
            return
        ctx.job_queue.run_once(END_JOB, END_TIMEOUT, chat_id=chat_id, name="end_timer_" + str(chat_id))
    except:
        pass

def SAVE_RECENT(ctx, text):
    lst = ctx.user_data.get("recent", [])
    lst.append(text)
    ctx.user_data["recent"] = lst[-10:]

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
        # قديم + جديد مع بعض - القديم الاول زي ما كان، وبعدين الجديد
        items = list(ROT) + list(ROT2)
        if q:
            ps = q.split()
            old_hits = []
            if len(ps) >= 2:
                old_hits = SROT(ps[0], ps[1])
            new_hits = SROT2(q)
            # لو البحث كلمتين والقديم جاب نتيجة، بنجمع الاتنين. لو لأ، الجديد بس
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

# ============================================================
# ===== نظام اليوزرز - جديد =====
# ============================================================
async def SEND_ADMIN_REQUEST(ctx, us, code):
    # دي الرسالة اللي بتوصلك انت كادمن - فيها بيانات اليوزر + زرارين قبول / رفض
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
    txt += f"📊 الحالة: pending"
    kb = MK([[B("✅ قبول", "accept_" + str(us.id)), B("❌ رفض", "reject_" + str(us.id))]])
    try:
        await ctx.bot.send_message(chat_id=int(ADM), text=txt, reply_markup=kb)
    except Exception as e:
        print(f"Admin notify fail: {e}")

async def START(update, ctx):
    ctx.user_data["end_sent"] = False
    us = update.message.from_user
    u = LU()
    uid = str(us.id)
    now = time.time()

    # الادمن نفسه - يدخل على طول
    if uid == str(ADM):
        await update.message.reply_text(WEL + "\n\nانت الادمن 👑\nاكتب /users عشان تشوف اليوزرز", reply_markup=MAIN_KB())
        return

    # مقبول قبل كده - يدخل على طول
    if uid in u and u[uid].get("status") == "accepted":
        await update.message.reply_text(WEL, reply_markup=MAIN_KB())
        return

    # مرفوض - ميقدرش يطلب تاني الا لما الادمن يقبله من /users
    if uid in u and u[uid].get("status") == "rejected":
        await update.message.reply_text("❌ تم رفض طلبك من قبل الادمن.\nتواصل مع الادمن لو شايف ان ده غلط.")
        return

    # معلق - موجود قبل كده - قوله استنى ومتزعجش الادمن تاني
    if uid in u and u[uid].get("status") == "pending":
        code = u[uid].get("code", "")
        await update.message.reply_text(f"⏳ طلبك لسه معلق عند الادمن\nالكود بتاعك: {code}\nاستنى الادمن يقبلك.")
        return

    # جديد خالص - سجله و ابعت للادمن
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
    # امر الادمن فقط - /users او /user
    if str(update.message.from_user.id) != str(ADM):
        await update.message.reply_text("❌ الامر ده للادمن بس")
        return
    u = LU()
    if not u:
        await update.message.reply_text("مفيش يوزرز خالص لسه")
        return

    # الاول المعلقين
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
            txt += f"📊 الحالة: pending"
            kb = MK([[B("✅ قبول", "accept_" + uid), B("❌ رفض", "reject_" + uid)]])
            await update.message.reply_text(txt, reply_markup=kb)

    if others:
        await update.message.reply_text(f"👥 باقي اليوزرز ({len(others)}):")
        for uid, d in others.items():
            emoji = "✅" if d.get("status") == "accepted" else "❌"
            txt = f"{emoji} {d.get('name','بدون اسم')} | {d.get('username','-')}\n"
            txt += f"ID: {uid} | كود: {d.get('code','-')} | الحالة: {d.get('status')}"
            # حتى المقبولين تقدر ترفضهم من هنا لو حبيت
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

    # القبول والرفض - للادمن بس - لازم يعدي قبل فحص OK
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
            # عدل رسالة الطلب الاصلية عشان تعرف انك قبلته
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

    # اي زرار تاني - لازم يكون مقبول
    if not OK(uid):
        await q.message.reply_text("⏳ استنى الادمن يقبلك الاول، ابعت /start")
        return

    if d == "m_main":
        await q.message.reply_text("القايمة:", reply_markup=MAIN_KB())
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
                t = it["name"] + "\n" + it["code"]
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
        return
    cap = (update.message.caption or "").strip()
    if not cap:
        cap = "بدون اسم"
    f = None
    ext = "jpg"
    if update.message.photo:
        f = await update.message.photo[-1].get_file()
    elif update.message.document:
        f = await update.message.document.get_file()
        ext = "pdf"
    else:
        return
    ds = LJ("docs_index.json", [])
    num = len(ds) + 1
    fn = "docs/" + str(num) + "." + ext
    await f.download_to_drive(fn)
    d = {}
    d["num"] = num
    d["name"] = cap
    d["path"] = fn
    ds.append(d)
    SJ("docs_index.json", ds)
    await update.message.reply_text("اتحفظ " + str(num))

async def HANDLE(update, ctx):
    us = update.message.from_user
    uid = str(us.id)
    text = update.message.text or ""
    low = text.strip().lower()
    bare = text.strip().lstrip(SL).strip().lower()
    RESET_END_TIMER(ctx)

    # ===== حماية اليوزرز - لازم تكون اول حاجة =====
    # الادمن يعدي - اي حد تاني لازم يكون accepted
    # /start و /users مش بيدخلوا هنا اصلا لانهم COMMAND
    if not OK(uid):
        await update.message.reply_text("⏳ استنى الادمن يقبلك الاول، ابعت /start")
        return

    if NA(text) in [NA(w) for w in END_WORDS]:
        await update.message.reply_text("العفو، نورتنا ❤️")
        await SEND_END_IMAGE(ctx, update.message.chat_id)
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
    # ===== الروتيشن الجديد - كلمة روتيشن / rotation تفتح القايمة زي القديم =====
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
            await update.message.reply_text(items[0]["name"] + "\n" + items[0]["code"])
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
            await update.message.reply_text(items[0]["name"] + "\n" + items[0]["code"])
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
            await update.message.reply_text(items[0]["name"] + "\n" + items[0]["code"])
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
            await update.message.reply_text(items[0]["name"] + "\n" + items[0]["code"])
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
            await update.message.reply_text(items[0]["name"] + "\n" + items[0]["code"])
            return
        await SEND(update, ctx, "cities", 0, qq)
        return
    if re.match(r"^[\d\.\s\-]+$", text.strip()):
        SAVE_RECENT(ctx, text)
        rs = ST(text.strip())
        if rs:
            if len(rs) == 1:
                await update.message.reply_text(FT(rs[0]))
                return
            await SEND(update, ctx, "tariff", 0, text.strip())
            return
        await update.message.reply_text("📜 ملقيتش بند")
        return
    ps = text.strip().split()
    if len(ps) == 2:
        rs = SROT(ps[0], ps[1])
        rs2 = SROT2(text.strip())
        if rs or rs2:
            SAVE_RECENT(ctx, text)
            # القديم الاول لو موجود، وبعدين الجديد - من غير ما نلغي القديم
            if rs:
                await update.message.reply_text(FROT(rs, ps[0], ps[1]))
            if rs2:
                await update.message.reply_text(FROT2(rs2, text.strip()))
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
        await update.message.reply_text(FROT2(rot2_hits, text.strip()))
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
