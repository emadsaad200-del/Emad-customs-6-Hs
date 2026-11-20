import os, json, time, random, string, re, difflib
import openpyxl
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, CallbackQueryHandler, ContextTypes, filters
BOT = os.getenv("BOT_TOKEN")
ADM = os.getenv("ADMIN_ID")
SL = chr(47)
os.makedirs("docs", exist_ok=True)
def JLOAD(f, d):
    try:
        with open(f, encoding="utf-8") as x:
            return json.load(x)
    except:
        return d
def JSAVE(f, d):
    with open(f, "w", encoding="utf-8") as x:
        json.dump(d, x, ensure_ascii=False, indent=2)
def CC(v):
    if v is None:
        return ""
    return str(v).strip()
def CINV(s):
    out = []
    for ch in str(s or ""):
        n = ord(ch)
        if n < 32:
            continue
        if n in (160, 1600, 8204, 8205, 8206, 8207, 1648):
            continue
        if 8234 <= n <= 8238:
            continue
        if 1611 <= n <= 1626:
            continue
        out.append(ch)
    return "".join(out)
def NA(s):
    s = CINV(s).strip().lower()
    s = s.replace("أ", "ا")
    s = s.replace("إ", "ا")
    s = s.replace("آ", "ا")
    s = s.replace("ة", "ه")
    s = s.replace("ى", "ي")
    s = s.replace("ئ", "ي")
    s = s.replace("ؤ", "و")
    s = s.translate(str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789"))
    s = re.sub(r"\s+", " ", s)
    return s.strip()
def CW(t):
    stop = ["ال", "شركه", "شركة", "لل", "ل", "شرك"]
    w = NA(t).split()
    o = []
    for x in w:
        if x not in stop and len(x) >= 2:
            o.append(x)
    if o:
        return o
    return w
BI = {}
BC = {}
BA = {}
for fn in ["gs1xlsx.xlsx", "GPC as of May 2023 v20230605 EN.xlsx"]:
    if not os.path.exists(fn):
        continue
    try:
        wb = openpyxl.load_workbook(fn, read_only=True, data_only=True)
        if "Schema" in wb.sheetnames:
            ws = wb["Schema"]
        else:
            ws = wb[wb.sheetnames[0]]
        for r in ws.iter_rows(values_only=True):
            try:
                if not r or len(r) < 11:
                    continue
                if str(r[0]).lower().startswith("segment"):
                    continue
                bc = str(r[9] or "")
                bt = str(r[10] or "")
                if not bt:
                    continue
                k = bt.strip().lower()
                if k not in BI:
                    d = {}
                    d["c"] = bc
                    d["t"] = bt
                    d["s"] = r[1]
                    d["sc"] = str(r[0])
                    d["f"] = r[4]
                    d["fc"] = str(r[3])
                    d["cl"] = r[7]
                    d["cc"] = str(r[6])
                    d["i"] = r[11] or ""
                    d["e"] = r[12] if len(r) > 12 else ""
                    BI[k] = d
                if bc:
                    BC[bc] = k
                if len(r) > 14 and r[13] is not None:
                    if k not in BA:
                        BA[k] = []
                    ex = None
                    for a in BA[k]:
                        if a["c"] == str(r[13]):
                            ex = a
                    if not ex:
                        ex = {"c": str(r[13]), "t": r[14], "v": []}
                        BA[k].append(ex)
                    vt = None
                    if len(r) > 17 and r[17] is not None:
                        vt = str(r[17])
                    if vt and vt not in ex["v"]:
                        ex["v"].append(vt)
            except:
                continue
        wb.close()
    except:
        continue
CI = {}
CB = {}
CO = {}
COB = {}
CU = {}
CUB = {}
PA = {}
PAB = {}
UO = {}
UOB = {}
if os.path.exists("cities.xlsx"):
    try:
        wb = openpyxl.load_workbook("cities.xlsx", read_only=True, data_only=True)
        def put(sheet, by, ix):
            if sheet not in wb.sheetnames:
                return
            ws = wb[sheet]
            for r in ws.iter_rows(values_only=True):
                try:
                    if not r or not r[0] or not r[1]:
                        continue
                    if str(r[0]).lower() in ("citycode", "code", "uomcode"):
                        continue
                    if "currency" in str(r[0]).lower():
                        continue
                    c = str(r[0]).strip().upper()
                    n = str(r[1]).strip()
                    ct = ""
                    if len(r) > 2 and r[2]:
                        ct = str(r[2]).strip()
                    if c in by:
                        continue
                    it = {"code": c, "name": n, "country": ct}
                    by[c] = it
                    k = n.lower()
                    if k not in ix:
                        ix[k] = []
                    ix[k].append(it)
                except:
                    continue
        put("City List", CB, CI)
        put("Cities", CB, CI)
        put("Country List", COB, CO)
        put("Currency List", CUB, CU)
        put("Package Type list", PAB, PA)
        put("UOM List", UOB, UO)
        wb.close()
    except:
        pass
CL = []
CBR = {}
CBF = {}
CN = {}
def AC(n, f, rg, en, ac2, st, ad, ct):
    n = CC(n)
    rg = CC(rg).replace(" ", "").replace(",", "")
    if not n or not rg:
        return
    if rg in CBR:
        return
    f = CC(f).replace(" ", "").replace(",", "")
    it = {}
    it["name"] = n
    it["file"] = f
    it["reg"] = rg
    it["entity"] = CC(en)
    it["act"] = CC(ac2)
    it["status"] = CC(st)
    it["addr"] = CC(ad)
    it["ctype"] = ct
    CL.append(it)
    CBR[rg] = it
    if f:
        CBF[f] = it
    idx = len(CL) - 1
    CN[idx] = {"n": NA(n), "a": NA(ac2), "d": NA(ad), "e": NA(en)}
if os.path.exists("ارقام_تسجيل_الشركات_111111.xlsx"):
    try:
        wb = openpyxl.load_workbook("ارقام_تسجيل_الشركات_111111.xlsx", read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]
        for r in ws.iter_rows(values_only=True):
            try:
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
            except:
                continue
        wb.close()
    except:
        pass
if os.path.exists("ارقام_تسجيل_الشركات_2026.xlsx"):
    try:
        wb = openpyxl.load_workbook("ارقام_تسجيل_الشركات_2026.xlsx", read_only=True, data_only=True)
        if "خاص" in wb.sheetnames:
            ws = wb["خاص"]
            for r in ws.iter_rows(values_only=True):
                try:
                    if not r or not r[0] or not r[1]:
                        continue
                    if str(r[0]).lower() == "registrationnumber":
                        continue
                    ent = "خاص - " + CC(r[3])
                    ad = "تاريخ التسجيل: " + CC(r[4])
                    AC(r[1], "", r[0], ent, "", CC(r[5]), ad, "private")
                except:
                    continue
        if "حكومي" in wb.sheetnames:
            ws = wb["حكومي"]
            for r in ws.iter_rows(values_only=True):
                try:
                    if not r or not r[1] or not r[2]:
                        continue
                    if str(r[1]).lower() == "رقم الجهة":
                        continue
                    if str(r[0]) == "قطاع حكومي":
                        continue
                    AC(r[2], "", r[1], "حكومي", "", "", CC(r[3]), "gov")
                except:
                    continue
        if "فرع اجنبي" in wb.sheetnames:
            ws = wb["فرع اجنبي"]
            for r in ws.iter_rows(values_only=True):
                try:
                    if not r or not r[1] or not r[2]:
                        continue
                    if str(r[1]).lower() == "رقم التسجيل":
                        continue
                    AC(r[2], "", r[1], "فرع اجنبي", "", "", CC(r[3]), "foreign")
                except:
                    continue
        wb.close()
    except:
        pass
def FC(c):
    t = "🏢 " + c["name"] + "\n"
    if c["reg"]:
        t = t + "🔢 رقم التسجيل: " + c["reg"] + "\n"
    if c["file"]:
        t = t + "📁 رقم الملف: " + c["file"] + "\n"
    if c["entity"]:
        t = t + "🏛️ الكيان: " + c["entity"] + "\n"
    if c["act"]:
        t = t + "📦 النشاط: " + c["act"] + "\n"
    if c["status"]:
        t = t + "📄 حاله الملف: " + c["status"] + "\n"
    if c["addr"]:
        t = t + "📍 العنوان: " + c["addr"] + "\n"
    return t.strip()
def GF(q="", ct=""):
    base = []
    for c in CL:
        if not ct or c["ctype"] == ct:
            base.append(c)
    n = NA(q)
    if not n:
        return base
    qd = q.strip().translate(str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789"))
    qd = qd.replace(" ", "")
    if qd.isdigit() and len(qd) >= 3:
        out = []
        for c in base:
            if qd in c["reg"] or qd in c["file"]:
                out.append(c)
        return out
    w = CW(q)
    out = []
    for c in base:
        nn = NA(c["name"])
        na2 = NA(c["act"])
        nd = NA(c["addr"])
        ne = NA(c["entity"])
        if w and all(x in nn for x in w):
            out.append(c)
        elif n in nn or n in na2 or n in nd or n in ne:
            out.append(c)
    return out
def SC(q, ct=""):
    q = (q or "").strip()
    if not q:
        return []
    if q in CBR:
        c = CBR[q]
        if not ct or c["ctype"] == ct:
            return [c]
        return []
    if q in CBF:
        c = CBF[q]
        if not ct or c["ctype"] == ct:
            return [c]
        return []
    return GF(q, ct)[:10]
TL = []
TB = {}
def NTC(s):
    s = str(s or "").strip().replace(" ", "").replace(",", "")
    if re.match(r"^\d{6}$", s):
        s = s[:4] + "." + s[4:]
    if re.match(r"^\d{8}$", s):
        s = s[:4] + "." + s[4:6] + "." + s[6:]
    return s
def CTT(v):
    s = str(v or "").strip()
    if s in ("", "-", "None"):
        return ""
    if s in ("معفاه", "معفاة"):
        return "معفاه"
    t = s.replace("%", "").strip()
    if t:
        return t + "%"
    return ""
def build_tariff():
    files = []
    if os.path.isdir("tarifa"):
        for f in sorted(os.listdir("tarifa")):
            if f.lower().endswith(".xlsx"):
                files.append(os.path.join("tarifa", f))
    for f in ["30285-1.xlsx", "70490-2.xlsx"]:
        if os.path.exists(f) and f not in files:
            files.append(f)
    for fn in files:
        try:
            wb = openpyxl.load_workbook(fn, read_only=True, data_only=True)
            ws = wb[wb.sheetnames[0]]
            for r in ws.iter_rows(values_only=True):
                try:
                    if not r:
                        continue
                    has = False
                    for x in r:
                        if x is not None and str(x).strip() != "":
                            has = True
                    if not has:
                        continue
                    if "رمز" in str(r[0]) or "الفصل" in str(r[0]):
                        continue
                    code = ""
                    desc = ""
                    tax = ""
                    ch = ""
                    mb = ""
                    if len(r) >= 5 and r[2] is not None and "رمز" not in str(r[2]):
                        code = NTC(r[2])
                        desc = CC(r[3])
                        tax = CTT(r[4])
                        ch = CC(r[0])
                        mb = CC(r[1])
                    elif len(r) >= 3 and r[0] is not None:
                        code = NTC(r[0])
                        desc = CC(r[1])
                        tax = CTT(r[2])
                        if "." in code:
                            ch = code.split(".")[0]
                        else:
                            ch = code[:2]
                        mb = code
                    else:
                        continue
                    if not code or code == "-":
                        continue
                    if not desc:
                        continue
                    if code in TB:
                        continue
                    it = {"code": code, "desc": desc, "tax": tax, "ch": ch, "mb": mb}
                    TB[code] = it
                    TL.append(it)
                except:
                    continue
            wb.close()
        except:
            continue
build_tariff()
def STF(q):
    q = (q or "").strip()
    if not q:
        return []
    cq = NTC(q)
    if cq in TB:
        return [TB[cq]]
    d = re.sub(r"\D", "", q)
    if d and len(d) >= 2:
        o = []
        for t in TL:
            if d in re.sub(r"\D", "", t["code"]):
                o.append(t)
        if o:
            return o[:10]
    n = NA(q)
    if len(n) < 2:
        return []
    out = []
    for t in TL:
        if n in NA(t["desc"]):
            out.append(t)
        if len(out) >= 10:
            break
    return out
def FT(t):
    x = "🧾 " + t["desc"] + "\n"
    x = x + "🔢 البند: " + t["code"] + "\n"
    if t["ch"]:
        x = x + "📂 الفصل: " + t["ch"] + "\n"
    if t["mb"] != t["code"]:
        x = x + "📁 البند الرئيسي: " + t["mb"] + "\n"
    tx = t["tax"] if t["tax"] else "-"
    x = x + "💰 الضريبة: " + tx
    return x.strip()
def GTF(q=""):
    n = NA(q)
    if not n:
        return TL
    d = re.sub(r"\D", "", q or "")
    out = []
    for t in TL:
        if n in NA(t["desc"]):
            out.append(t)
        elif d and len(d) >= 2 and d in re.sub(r"\D", "", t["code"]):
            out.append(t)
    return out
LC = {}
LN = {}
ROT = []
if os.path.exists("ايميلات_وارقام_توكيلات.xlsx"):
    try:
        wb = openpyxl.load_workbook("ايميلات_وارقام_توكيلات.xlsx", read_only=True, data_only=True)
        if "EMAILS" in wb.sheetnames:
            cur = ""
            rows = list(wb["EMAILS"].iter_rows(values_only=True))
            for r in rows[2:]:
                try:
                    if not r:
                        continue
                    if r[0]:
                        cur = str(r[0]).split("\n")[0].strip()
                    if not cur:
                        continue
                    k = cur.upper()
                    if k not in LN:
                        LN[k] = cur
                    v = []
                    for i in range(1, 11):
                        if len(r) > i:
                            v.append(CC(r[i]))
                        else:
                            v.append("")
                    if not any(v):
                        continue
                    it = {}
                    it["line"] = cur
                    it["name"] = v[0]
                    it["pos"] = v[1]
                    it["email"] = v[2]
                    it["num"] = v[3]
                    it["book"] = v[4]
                    it["inv"] = v[8]
                    if k not in LC:
                        LC[k] = []
                    LC[k].append(it)
                except:
                    continue
        if "SERVICE ROTATION" in wb.sheetnames:
            lp = ""
            ws = wb["SERVICE ROTATION"]
            for r in ws.iter_rows(values_only=True):
                try:
                    if not r:
                        continue
                    if str(r[0]).upper() == "POL":
                        continue
                    pol = CC(r[0]) or lp
                    if pol:
                        lp = pol
                    pod = CC(r[1])
                    if not pol or not pod:
                        continue
                    it = {}
                    it["pol"] = pol
                    it["pod"] = pod
                    it["line"] = CC(r[2])
                    it["tt"] = CC(r[3])
                    it["co"] = CC(r[4])
                    ROT.append(it)
                except:
                    continue
        wb.close()
        if os.path.exists("lines_override.json"):
            ov = JLOAD("lines_override.json", {})
            for k, lst in ov.items():
                if k not in LN:
                    LN[k] = k
                if k not in LC:
                    LC[k] = []
                LC[k].extend(lst)
    except:
        pass
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
AM["بي اي ال"] = "PIL"
AM["ترانسمار"] = "TRANSMAR"
AM["هيونداي"] = "HYUNDAI"
AM["دلتا"] = "DELTA"
SM = {"م": "MAERSK", "ه": "HAPAG", "ك": "COSCO", "س": "CMA", "ا": "EVERGREEN"}
def RL(q):
    q = (q or "").strip()
    if not q:
        return None
    if q.upper() in LC:
        return q.upper()
    if q.lower() in AM:
        return AM[q.lower()]
    if q.strip() in SM:
        return SM[q.strip()]
    ql = q.lower()
    for k in LN:
        if ql in k.lower() or ql in LN[k].lower():
            return k
    return None
def SG(q):
    m = difflib.get_close_matches((q or "").strip().upper(), list(LN.keys()), n=1, cutoff=0.55)
    if m:
        return m[0]
    m2 = difflib.get_close_matches((q or "").strip().lower(), list(AM.keys()), n=1, cutoff=0.55)
    if m2:
        return AM[m2[0]]
    return None
def FLC(k, cat="ALL"):
    cs = LC.get(k, [])
    d = LN.get(k, k)
    if not cs:
        return "❌ ملقيتش توكيل اسمه: " + d
    t = "🚢 " + d + "\n"
    if cat == "BOOKING":
        t = t + "📘 Booking:\n\n"
        for x in cs:
            if x["book"] or x["email"]:
                nm = x["name"] or x["pos"] or ""
                t = t + "• " + nm + " - " + x["email"] + " " + x["num"] + "\n"
                t = t + "Booking: " + x["book"] + "\n\n"
        return t.strip()
    if cat == "INVOICE":
        t = t + "🧾 Invoice:\n\n"
        for x in cs:
            if x["inv"]:
                nm = x["name"] or ""
                t = t + "• " + nm + " - " + x["email"] + "\n"
                t = t + "Invoice: " + x["inv"] + "\n\n"
        return t.strip()
    if cat == "PHONE":
        t = t + "📞 الارقام:\n\n"
        for x in cs:
            if x["num"]:
                nm = x["name"] or ""
                t = t + "• " + nm + ": " + x["num"] + "\n"
        return t.strip()
    t = t + "📞 العدد: " + str(len(cs)) + "\n\n"
    for c in cs[:12]:
        if c["name"]:
            t = t + "👤 " + c["name"] + "\n"
        if c["pos"]:
            t = t + "💼 " + c["pos"] + "\n"
        if c["email"]:
            t = t + "📧 " + c["email"] + "\n"
        if c["num"]:
            t = t + "📞 " + c["num"] + "\n"
        if c["book"]:
            t = t + "📘 Booking: " + c["book"] + "\n"
        if c["inv"]:
            t = t + "🧾 Invoice: " + c["inv"] + "\n"
        t = t + "\n"
    return t.strip()
def SROT(pol, pod):
    out = []
    for r in ROT:
        if pol.lower() in r["pol"].lower() and pod.lower() in r["pod"].lower():
            out.append(r)
    return out[:10]
def FROT(rs, pol, pod):
    if not rs:
        return "❌ ملقيتش خط من " + pol + " لـ " + pod
    t = "🚢 من " + pol + " لـ " + pod + " - " + str(len(rs)) + ":\n\n"
    for r in rs:
        line = "• " + r["line"] + " | " + r["pol"] + " -> " + r["pod"]
        if r["tt"]:
            line = line + " - T.T: " + r["tt"]
        line = line + " | " + r["co"] + "\n"
        t = t + line
    return t
def MAIN_KB():
    kb = []
    row1 = []
    row1.append(InlineKeyboardButton("🚢 توكيلات", callback_data="menu_lines"))
    row1.append(InlineKeyboardButton("📦 GPC", callback_data="menu_gpc"))
    kb.append(row1)
    row2 = []
    row2.append(InlineKeyboardButton("🧾 التعريفة", callback_data="menu_tariff"))
    row2.append(InlineKeyboardButton("📁 المستندات", callback_data="menu_docs"))
    kb.append(row2)
    row3 = []
    row3.append(InlineKeyboardButton("🚢 Rotation", callback_data="menu_rotation"))
    row3.append(InlineKeyboardButton("🏢 الشركات", callback_data="menu_companies"))
    kb.append(row3)
    row4 = []
    row4.append(InlineKeyboardButton("🌍 Country", callback_data="menu_country"))
    row4.append(InlineKeyboardButton("🏙️ Cities", callback_data="menu_cities"))
    kb.append(row4)
    row5 = []
    row5.append(InlineKeyboardButton("💵 Currency", callback_data="menu_currency"))
    row5.append(InlineKeyboardButton("📦 Package", callback_data="menu_package"))
    kb.append(row5)
    row6 = []
    row6.append(InlineKeyboardButton("📏 UOM", callback_data="menu_uom"))
    row6.append(InlineKeyboardButton("🕘 اخر بحث", callback_data="menu_recent"))
    kb.append(row6)
    row7 = []
    row7.append(InlineKeyboardButton("❓ مساعدة", callback_data="menu_help"))
    kb.append(row7)
    return InlineKeyboardMarkup(kb)
def COMP_KB():
    kb = []
    kb.append([InlineKeyboardButton("🏢 الشركات الخاصة", callback_data="comp_type_private")])
    kb.append([InlineKeyboardButton("🏛️ الشركات الحكومية", callback_data="comp_type_gov")])
    kb.append([InlineKeyboardButton("🌍 الشركات الاجنبية", callback_data="comp_type_foreign")])
    kb.append([InlineKeyboardButton("🏠 القايمة الرئيسية", callback_data="menu_main")])
    return InlineKeyboardMarkup(kb)
def LCAT_KB(k):
    kb = []
    r1 = []
    r1.append(InlineKeyboardButton("📘 Booking", callback_data="linecat_" + k + "_BOOKING"))
    r1.append(InlineKeyboardButton("