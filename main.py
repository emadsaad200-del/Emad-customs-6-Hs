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

# صورة النهاية - لازم تكون جنب الكود بنفس الاسم
END_IMG = "end_session.jpg"
END_WORDS = ["خلصت", "شكرا", "شكرا لك", "تمام كده", "كفايه", "كفاية", "سلام", "bye", "thanks", "thank you"]
# لو عايز تغير مدة السكوت غير الرقم ده - 600 ثانية = 10 دقايق
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

# ===== GPC - من ملفه بس =====
BI = {}
BC = {}
BA = {}
for fn in ["gs1xlsx.xlsx"]:
    if not os.path.exists(fn):
        continue
    wb = openpyxl.load_workbook(fn, read_only=True)
    ws = wb[wb.sheetnames[0]]
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

# ===== المراجع - من ملف cities.xlsx بس =====
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

# ===== الشركات - من 2026 بس، 111111 اتلغى من الكود =====
CL = []
CBR = {}
def AC(n, f, rg, en, ac, st, ad, ct):
    n = CC(n)
    rg = CC(rg).replace(" ", "")
    if not n or not rg:
        return
    if rg in CBR:
        return
    # تخطي الصف البايظ في شيت حكومي
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

# ملحوظة: ملف 111111 موجود عندك بس البوت مش بيقراه تاني عشان كان بيبوظ النتايج
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

# ===== التعريفة - كل ملفات tarifa + توحيد الأعمدة =====
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
    try:
        wb = openpyxl.load_workbook(fn, read_only=True)
    except:
        return
    ws = wb[wb.sheetnames[0]]
    for r in ws.iter_rows(values_only=True):
        if not r:
            continue
        code = NTC(GV(r, 0))
        desc = CC(GV(r, 1))
        if not code or not desc:
            continue
        if code in TB:
            # لو البند موجود قبل كده، كمل بياناته لو ناقصة - مش تمسح القديم
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
    qd = q.replace(".", "").replace(" ", "")
    if qd.isdigit():
        for code, it in TB.items():
            if code.replace(".", "") == qd or code.replace(".", "").startswith(qd):
                return [it]
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

# ===== التوكيلات + Rotation - من ملف التوكيلات بس =====
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

WEL = "Welcome to Captain Emad Customs Bot"
HELP = """❓ المساعدة - كل أيقونة بتعمل ايه:

🚢 توكيلات: اكتب MAERSK او ميرسك - هيطلعلك الايميلات والارقام والـ Booking والـ Invoice
📦 GPC: اكتب GPC Airbrushes او GPC + الاسم او الكود
📜 التعريفة: اكتب 8420 او 84 او اسم الصنف او جزء منه او كلمة دالة - مثال: مروحة
    - دوس على البند هيطلعلك بيانات التعريفة كاملة
    - استخدم التالي / السابق عشان تقلب في النتايج
🏢 شركات: اكتب النور او رقم التسجيل 100271510 - اختار الخاصة / الحكومية / الاجنبية من القايمة
🔄 Rotation: اكتب Alex Izmir - مدينة البداية ومدينة الوصول
🌍 Country: اكتب country Egypt او country مص
🏙️ Cities: اكتب city Cairo او city القاهرة
💱 Currency: اكتب currency USD
📦 Package: اكتب package Box
⚖️ UOM: اكتب uom KG
📄 مستندات: اكتب تزيل 1 او تزيل اسم المستند 1 - او دوس على زرار مستندات
🕐 اخر بحث: بيجيب اخر حاجة بحثت عنها
القايمة: اكتب menu

لو بحثك لقى نتيجة في اكتر من ايقونة، البوت هيخيرك تختار انهي واحدة.
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

# ===== صورة النهاية =====
async def SEND_END_IMAGE(ctx, chat_id):
    # مرة واحدة بس في الجلسة
    if ctx.user_data.get("end_sent"):
        return
    ctx.user_data["end_sent"] = True
    if os.path.exists(END_IMG):
        try:
            await ctx.bot.send_photo(chat_id=chat_id, photo=open(END_IMG, "rb"), caption="لو حابب تكمل أو تشترك، دي بيانات التواصل بتاعتنا ❤️")
            return
        except:
            pass
    # لو الصورة مش موجودة، ابعت نص بديل
    await ctx.bot.send_message(chat_id=chat_id, text="لو حابب تكمل أو تشترك:\nTelegram: https://t.me/tarrifabot\nWhatsApp: wa.me/+201094038730\nPhone: 01094038730")

async def END_JOB(ctx):
    # شغل بعد 10 دقايق سكوت
    chat_id = ctx.job.chat_id
    await SEND_END_IMAGE(ctx, chat_id)

def RESET_END_TIMER(ctx):
    # الغي المؤقت القديم وابدأ واحد جديد - عشان الـ 10 دقايق تتحسب من اخر رسالة
    try:
        jobs = ctx.job_queue.get_jobs_by_name("end_timer_" + str(ctx._chat_id if hasattr(ctx, "_chat_id") else ""))
        for j in jobs:
            j.schedule_removal()
    except:
        pass
    try:
        ctx.job_queue.run_once(END_JOB, END_TIMEOUT, chat_id=ctx._chat_id if hasattr(ctx, "_chat_id") else None, name="end_timer")
    except:
        pass

def SAVE_RECENT(ctx, text):
    lst = ctx.user_data.get("recent", [])
    lst.append(text)
    # خلي اخر 10 بس
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
            if not items:
                items = [t for t in TL if NA(q) in NA(t["desc"]) or q in t["code"]]
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
        items = ROT
        if q:
            ps = q.split()
            if len(ps) >= 2:
                items = SROT(ps[0], ps[1])
        kb, txt = PG(items, page, "rot_", lambda r: "🔄 " + r["pol"] + " -> " + r["pod"] + " | " + r["line"])
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

async def START(update, ctx):
    ctx.user_data["end_sent"] = False
    us = update.message.from_user
    u = LU()
    uid = str(us.id)
    now = time.time()
    if uid in u and u[uid].get("status") == "accepted":
        await update.message.reply_text(WEL, reply_markup=MAIN_KB())
        return
    if uid not in u:
        code = "USR-" + "".join(random.choices("0123456789", k=4))
        d = {}
        d["code"] = code
        d["name"] = us.first_name or ""
        d["status"] = "pending"
        d["requested_at"] = now
        u[uid] = d
        SU(u)
    else:
        code = u[uid]["code"]
        u[uid]["status"] = "pending"
        SU(u)
    msg = WEL + "\nCode: " + code
    await update.message.reply_text(msg, reply_markup=MAIN_KB())
    if ADM:
        a = B("قبول", "accept_" + uid)
        b = B("رفض", "reject_" + uid)
        kb = MK([[a, b]])
        try:
            await ctx.bot.send_message(chat_id=int(ADM), text="New: " + code, reply_markup=kb)
        except:
            pass

async def BTN(update, ctx):
    q = update.callback_query
    try:
        await q.answer()
    except:
        pass
    d = q.data
    uid = q.from_user.id
    # اي ضغطة زرار بتلغي مؤقت النهاية وتبدأ واحد جديد
    try:
        ctx.job_queue.run_once(END_JOB, END_TIMEOUT, chat_id=q.message.chat_id, name="end_timer")
    except:
        pass
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
    # اختيار من البحث المتعدد
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
            await q.message.reply_text(FROT([r], r["pol"], r["pod"]))
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
    if d.startswith("accept_"):
        uid2 = d.split("_")[1]
        u = LU()
        if uid2 in u:
            u[uid2]["status"] = "accepted"
            SU(u)
            await ctx.bot.send_message(chat_id=int(uid2), text=WEL, reply_markup=MAIN_KB())
            await q.message.reply_text("Accepted")
        return
    if d.startswith("reject_"):
        uid2 = d.split("_")[1]
        u = LU()
        if uid2 in u:
            u[uid2]["status"] = "rejected"
            SU(u)
            await q.message.reply_text("Rejected")
        return
    if d.startswith("gpc_"):
        try:
            i = int(d.split("_")[1])
            lst = ctx.user_data.get("gpc", [])
            if 0 <= i < len(lst):
                await q.message.reply_text(FB(lst[i]))
        except:
            pass
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
    # جدد مؤقت الـ 10 دقايق مع كل رسالة
    try:
        ctx.job_queue.run_once(END_JOB, END_TIMEOUT, chat_id=update.message.chat_id, name="end_timer")
    except:
        pass
    # كلمة نهاية - ابعت الصورة مرة واحدة
    if NA(text) in [NA(w) for w in END_WORDS]:
        await update.message.reply_text("العفو، نورتنا ❤️")
        await SEND_END_IMAGE(ctx, update.message.chat_id)
        return
    # 3 - مستندات
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
            ctx.user_data["gpc"] = b[1]
            rows = []
            for i, x in enumerate(b[1]):
                rows.append([B("📦 " + x["t"][:50], "gpc_" + str(i))])
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
    if re.match(r"^[\d\.\s]+$", text.strip()):
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
        if rs:
            SAVE_RECENT(ctx, text)
            await update.message.reply_text(FROT(rs, ps[0], ps[1]))
            return
    # ===== البحث الذكي المتعدد - لو لقى في اكتر من ايقونة خيّر المستخدم =====
    # بنفحص من غير ما نرجع نتيجة واحدة غلط
    k_line = RL(text.strip())
    comp_hits = GC(text.strip(), "")
    tariff_hits = ST(text.strip())
    gpc_hit = SGPC(text)
    found = []
    if k_line: found.append("line")
    if comp_hits: found.append("comp")
    if tariff_hits: found.append("tariff")
    if gpc_hit is not None: found.append("gpc")
    # لو لقى في 2 او اكتر، خيّر المستخدم - وده الجديد
    if len(found) >= 2:
        SAVE_RECENT(ctx, text)
        ctx.user_data["multi_q"] = text.strip()
        names = {"line": "🚢 التوكيلات", "comp": "🏢 الشركات", "tariff": "📜 التعريفة", "gpc": "📦 GPC"}
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
    if not OK(uid):
        await update.message.reply_text("ابعت start الاول", reply_markup=MAIN_KB())
        return
    if gpc_hit is not None:
        SAVE_RECENT(ctx, text)
        if isinstance(gpc_hit, tuple):
            ctx.user_data["gpc"] = gpc_hit[1]
            rows = []
            for i, x in enumerate(gpc_hit[1]):
                rows.append([B("📦 " + x["t"][:50], "gpc_" + str(i))])
            kb = MK(rows)
            await update.message.reply_text("اختار صنف:", reply_markup=kb)
            return
        await update.message.reply_text(FB(gpc_hit))
        return
    await update.message.reply_text("وصلت", reply_markup=MAIN_KB())

def MAIN():
    print("Bot starting...")
    app = ApplicationBuilder().token(BOT).build()
    app.add_handler(CommandHandler("start", START))
    app.add_handler(CallbackQueryHandler(BTN))
    app.add_handler(MessageHandler(filters.PHOTO | filters.Document.ALL, SAVE_DOC))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, HANDLE))
    app.run_polling()
    print("Bot stopped")
MAIN()
