import os, json, time, random, string, re, difflib
import openpyxl
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, CallbackQueryHandler, ContextTypes, filters

BOT_TOKEN = os.getenv("BOT_TOKEN")
ADMIN_ID = os.getenv("ADMIN_ID")
USERS_FILE = "users.json"
REJECT_WAIT = 5 * 60
DOCS_FILE = "docs_index.json"
STATS_FILE = "hits.json"
OVERRIDE_FILE = "lines_override.json"
os.makedirs("docs", exist_ok=True)
EXCEL_FILES = ["gs1xlsx.xlsx", "GPC as of May 2023 v20230605 EN.xlsx"]

# ===== GPC =====
BRICK_INDEX = {}; BRICK_BY_CODE = {}; BRICK_ATTRS = {}
def build_index():
    print("Building light index...")
    for fname in EXCEL_FILES:
        if not os.path.exists(fname): continue
        try:
            wb = openpyxl.load_workbook(fname, read_only=True, data_only=True)
            ws = wb["Schema"] if "Schema" in wb.sheetnames else wb[wb.sheetnames[0]]
            for r in ws.iter_rows(values_only=True):
                try:
                    if not r or len(r) < 11: continue
                    if str(r[0]).lower().startswith("segment"): continue
                    brick_code = str(r[9]) if r[9] is not None else ""
                    brick_title = str(r[10]) if r[10] is not None else ""
                    if not brick_title: continue
                    key = brick_title.strip().lower()
                    if key not in BRICK_INDEX:
                        BRICK_INDEX[key] = {"brick_code": brick_code, "brick_title": brick_title, "segment": r[1], "segment_code": str(r[0]), "family": r[4], "family_code": str(r[3]), "class": r[7], "class_code": str(r[6]), "includes": r[11] or "", "excludes": r[12] if len(r) > 12 else "" or ""}
                        if brick_code: BRICK_BY_CODE[brick_code] = key
                    if len(r) > 14 and r[13] is not None:
                        attr_code = str(r[13]); attr_title = r[14]; val_title = str(r[17]) if len(r) > 17 and r[17] is not None else None
                        if key not in BRICK_ATTRS: BRICK_ATTRS[key] = []
                        ex = None
                        for a in BRICK_ATTRS[key]:
                            if a["attribute_code"] == attr_code: ex = a; break
                        if not ex:
                            ex = {"attribute_code": attr_code, "attribute_title": attr_title, "values": []}; BRICK_ATTRS[key].append(ex)
                        if val_title and val_title not in ex["values"]: ex["values"].append(val_title)
                except: continue
            wb.close(); print(f"Indexed {fname}: {len(BRICK_INDEX)} bricks")
        except Exception as e: print(f"Index error {fname}: {e}")
    print(f"Index done: {len(BRICK_INDEX)} bricks")
build_index()

# ===== Cities.xlsx =====
CITIES_FILE = "cities.xlsx"
CITY_INDEX = {}; CITY_BY_CODE = {}; COUNTRY_INDEX = {}; COUNTRY_BY_CODE = {}; CURRENCY_INDEX = {}; CURRENCY_BY_CODE = {}; PACKAGE_INDEX = {}; PACKAGE_BY_CODE = {}; UOM_INDEX = {}; UOM_BY_CODE = {}
def build_reference_index():
    if not os.path.exists(CITIES_FILE): print("cities.xlsx not found"); return
    try:
        wb = openpyxl.load_workbook(CITIES_FILE, read_only=True, data_only=True)
        if "City List" in wb.sheetnames:
            ws = wb["City List"]
            for r in ws.iter_rows(values_only=True):
                try:
                    if not r or len(r) < 2 or not r[0] or not r[1]: continue
                    if str(r[0]).lower() == "citycode": continue
                    code = str(r[0]).strip().upper(); name = str(r[1]).strip(); country = str(r[2]).strip() if len(r) > 2 and r[2] else ""
                    CITY_BY_CODE[code] = {"code": code, "name": name, "country": country}
                    k = name.lower()
                    if k not in CITY_INDEX: CITY_INDEX[k] = []
                    CITY_INDEX[k].append({"code": code, "name": name, "country": country})
                except: continue
        if "Cities" in wb.sheetnames:
            ws = wb["Cities"]
            for r in ws.iter_rows(values_only=True):
                try:
                    if not r or len(r) < 2 or not r[0] or not r[1]: continue
                    if str(r[0]).lower() == "citycode": continue
                    code = str(r[0]).strip().upper(); name = str(r[1]).strip()
                    if code in CITY_BY_CODE: continue
                    CITY_BY_CODE[code] = {"code": code, "name": name, "country": ""}
                    k = name.lower()
                    if k not in CITY_INDEX: CITY_INDEX[k] = []
                    CITY_INDEX[k].append({"code": code, "name": name, "country": ""})
                except: continue
        if "Country List" in wb.sheetnames:
            ws = wb["Country List"]
            for r in ws.iter_rows(values_only=True):
                try:
                    if not r or len(r) < 2 or not r[0] or not r[1]: continue
                    if str(r[0]).lower() == "code": continue
                    code = str(r[0]).strip().upper(); name = str(r[1]).strip()
                    COUNTRY_BY_CODE[code] = {"code": code, "name": name}
                    k = name.lower()
                    if k not in COUNTRY_INDEX: COUNTRY_INDEX[k] = []
                    COUNTRY_INDEX[k].append({"code": code, "name": name})
                except: continue
        if "Currency List" in wb.sheetnames:
            ws = wb["Currency List"]
            for r in ws.iter_rows(values_only=True):
                try:
                    if not r or len(r) < 2 or not r[0] or not r[1]: continue
                    if "currency" in str(r[0]).lower(): continue
                    name = str(r[0]).strip(); code = str(r[1]).strip().upper()
                    CURRENCY_BY_CODE[code] = {"code": code, "name": name}
                    k = name.lower()
                    if k not in CURRENCY_INDEX: CURRENCY_INDEX[k] = []
                    CURRENCY_INDEX[k].append({"code": code, "name": name})
                except: continue
        if "Package Type list" in wb.sheetnames:
            ws = wb["Package Type list"]
            for r in ws.iter_rows(values_only=True):
                try:
                    if not r or len(r) < 2 or not r[0] or not r[1]: continue
                    if str(r[0]).lower() == "code": continue
                    code = str(r[0]).strip().upper(); name = str(r[1]).strip()
                    PACKAGE_BY_CODE[code] = {"code": code, "name": name}
                    k = name.lower()
                    if k not in PACKAGE_INDEX: PACKAGE_INDEX[k] = []
                    PACKAGE_INDEX[k].append({"code": code, "name": name})
                except: continue
        if "UOM List" in wb.sheetnames:
            ws = wb["UOM List"]
            for r in ws.iter_rows(values_only=True):
                try:
                    if not r or len(r) < 2 or not r[0] or not r[1]: continue
                    if "uomcode" in str(r[0]).lower(): continue
                    code = str(r[0]).strip().upper(); name = str(r[1]).strip()
                    UOM_BY_CODE[code] = {"code": code, "name": name}
                    k = name.lower()
                    if k not in UOM_INDEX: UOM_INDEX[k] = []
                    UOM_INDEX[k].append({"code": code, "name": name})
                except: continue
        wb.close(); print(f"Ref indexed: Cities {len(CITY_BY_CODE)} | Country {len(COUNTRY_BY_CODE)} | Currency {len(CURRENCY_BY_CODE)} | Package {len(PACKAGE_BY_CODE)} | UOM {len(UOM_BY_CODE)}")
    except Exception as e: print(f"Ref index error: {e}")
build_reference_index()

# ===== التوكيلات =====
SHIPPING_FILE = "ايميلات_وارقام_توكيلات.xlsx"
LINE_CONTACTS = {}; LINE_NAMES = {}; ROTATION = []; SLMC_LIST = []
def clean_cell(v): return "" if v is None else str(v).strip()
def build_shipping_index():
    if not os.path.exists(SHIPPING_FILE): print("ايميلات_وارقام_توكيلات.xlsx not found"); return
    try:
        wb = openpyxl.load_workbook(SHIPPING_FILE, read_only=True, data_only=True)
        if "EMAILS" in wb.sheetnames:
            ws = wb["EMAILS"]; rows = list(ws.iter_rows(values_only=True)); cur_line = ""
            for r in rows[2:]:
                try:
                    if not r: continue
                    if r[0]: cur_line = str(r[0]).split("\n")[0].strip()
                    if not cur_line: continue
                    key = cur_line.upper()
                    if key not in LINE_NAMES: LINE_NAMES[key] = cur_line
                    name = clean_cell(r[1]) if len(r) > 1 else ""; pos = clean_cell(r[2]) if len(r) > 2 else ""; email = clean_cell(r[3]) if len(r) > 3 else ""; num = clean_cell(r[4]) if len(r) > 4 else ""; booking = clean_cell(r[5]) if len(r) > 5 else ""; final = clean_cell(r[6]) if len(r) > 6 else ""; vgm = clean_cell(r[7]) if len(r) > 7 else ""; draft = clean_cell(r[8]) if len(r) > 8 else ""; invoice = clean_cell(r[9]) if len(r) > 9 else ""; pickup = clean_cell(r[10]) if len(r) > 10 else ""
                    if not any([name, pos, email, num, booking, final, vgm, draft, invoice, pickup]): continue
                    if key not in LINE_CONTACTS: LINE_CONTACTS[key] = []
                    LINE_CONTACTS[key].append({"line": cur_line, "name": name, "position": pos, "email": email, "number": num, "booking": booking, "final": final, "vgm": vgm, "draft": draft, "invoice": invoice, "pickup": pickup})
                except: continue
        if "SERVICE ROTATION" in wb.sheetnames:
            ws = wb["SERVICE ROTATION"]
            for r in ws.iter_rows(values_only=True):
                try:
                    if not r or not r[0] or not r[1]: continue
                    if str(r[0]).upper() == "POL": continue
                    ROTATION.append({"pol": clean_cell(r[0]), "pod": clean_cell(r[1]), "line": clean_cell(r[2]), "tt": clean_cell(r[3]), "country": clean_cell(r[4])})
                except: continue
        if "SLMC" in wb.sheetnames:
            ws = wb["SLMC"]
            for r in ws.iter_rows(values_only=True):
                try:
                    if not r or not r[1] or not r[2]: continue
                    if str(r[1]).upper() == "POL": continue
                    SLMC_LIST.append({"pol": clean_cell(r[1]), "pod": clean_cell(r[2]), "line": clean_cell(r[3]), "tt": clean_cell(r[4]), "country": clean_cell(r[5])})
                except: continue
        wb.close()
        if os.path.exists(OVERRIDE_FILE):
            try:
                with open(OVERRIDE_FILE, "r", encoding="utf-8") as f: ov = json.load(f)
                for k, lst in ov.items():
                    if k not in LINE_NAMES: LINE_NAMES[k] = k
                    if k not in LINE_CONTACTS: LINE_CONTACTS[k] = []
                    LINE_CONTACTS[k].extend(lst)
            except: pass
        print(f"Shipping indexed: Lines {len(LINE_NAMES)} | Rotation {len(ROTATION)} | SLMC {len(SLMC_LIST)}")
    except Exception as e: print(f"Shipping index error: {e}")
build_shipping_index()

# ===== helpers =====
ARABIC_LINE_MAP = {"ميرسك":"MAERSK","هاباج":"HAPAG","هاباق":"HAPAG","كوسكو":"COSCO","سي ام ايه":"CMA","سي ام اي":"CMA","ايفرجرين":"EVERGREEN","ايفر جرين":"EVERGREEN","ام اس سي":"MSC","سيجو":"SEAGO","سيغو":"SEAGO","زيم":"ZIM","وان هاي":"WAN HAI","وانهاي":"WAN HAI","اركاس":"ARKAS","اركس":"ARKAS","وان":"ONE","او او سي ال":"OOCL","واي ام ال":"YML","اي بي ال":"APL","بي اي ال":"PIL","بيل":"PIL","تاروس":"TARROS","تاروز":"TARROS","توركون":"TURKON","ميدكون":"MEDKON","ميسينا":"MESSINA","ساف مارين":"SAF MARINE","هامبورج سود":"HAMBURG SUD","هامبورغ سود":"HAMBURG SUD","بي ام سي":"BMC","ترانسمار":"TRANSMAR","ترانس مار":"TRANSMAR","ويك لاين":"WEC LINE","جريمالدي":"GRIMALDI","هيونداي":"HYUNDAI","هيونداى":"HYUNDAI","دلتا":"DELTA","ام جي ال":"MGL","روديمار":"RODYMAR","نايل داتش":"NILE DUTCH","ادميرال":"ADMIRAL","بورشارد":"BOARSHARD","مارين":"MARINE"}
SHORTCUT_MAP = {"م":"MAERSK","ه":"HAPAG","ك":"COSCO","س":"CMA","ا":"EVERGREEN"}
def load_hits():
    if not os.path.exists(STATS_FILE): return {"lines":{},"ports":{}}
    try:
        with open(STATS_FILE,"r",encoding="utf-8") as f: return json.load(f)
    except: return {"lines":{},"ports":{}}
def save_hits(h):
    try:
        with open(STATS_FILE,"w",encoding="utf-8") as f: json.dump(h,f,ensure_ascii=False,indent=2)
    except: pass
def hit_line(line):
    h = load_hits(); h["lines"][line] = h["lines"].get(line,0)+1; save_hits(h)
def hit_port(port):
    h = load_hits(); h["ports"][port] = h["ports"].get(port,0)+1; save_hits(h)
def resolve_line(q):
    qq = (q or "").strip()
    if not qq: return None
    if qq.upper() in LINE_CONTACTS: return qq.upper()
    if qq.lower() in ARABIC_LINE_MAP: return ARABIC_LINE_MAP[qq.lower()]
    if qq.strip() in SHORTCUT_MAP: return SHORTCUT_MAP[qq.strip()]
    ql = qq.lower()
    for k in LINE_NAMES:
        if ql in k.lower() or ql in LINE_NAMES[k].lower(): return k
    return None
def suggest_line(q):
    m = difflib.get_close_matches((q or "").strip().upper(), list(LINE_NAMES.keys()), n=1, cutoff=0.55)
    if m: return m[0]
    m2 = difflib.get_close_matches((q or "").strip().lower(), list(ARABIC_LINE_MAP.keys()), n=1, cutoff=0.55)
    if m2: return ARABIC_LINE_MAP[m2[0]]
    return None
def add_recent(context, txt):
    lst = context.user_data.get("recent", []); txt = (txt or "").strip()
    if txt and txt not in lst: lst.insert(0, txt)
    context.user_data["recent"] = lst[:5]

# ===== المنيو الكامل - فيه كل حاجة =====
def main_menu_keyboard():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🚢 توكيلات", callback_data="menu_lines"), InlineKeyboardButton("📦 GPC", callback_data="menu_gpc")],
        [InlineKeyboardButton("🚢 Rotation", callback_data="menu_rotation"), InlineKeyboardButton("📁 المستندات / تزيل", callback_data="menu_docs")],
        [InlineKeyboardButton("🌍 Country", callback_data="menu_country"), InlineKeyboardButton("🏙️ Cities", callback_data="menu_cities")],
        [InlineKeyboardButton("💵 Currency", callback_data="menu_currency"), InlineKeyboardButton("📦 Package", callback_data="menu_package")],
        [InlineKeyboardButton("📏 UOM", callback_data="menu_uom"), InlineKeyboardButton("🕘 اخر بحث", callback_data="menu_recent")],
        [InlineKeyboardButton("❓ مساعدة / Help", callback_data="menu_help")],
    ])
def line_category_keyboard(line_key):
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("📘 Booking", callback_data=f"linecat_{line_key}_BOOKING"), InlineKeyboardButton("🧾 Invoice", callback_data=f"linecat_{line_key}_INVOICE")],
        [InlineKeyboardButton("📧 كل الايميلات", callback_data=f"linecat_{line_key}_ALL"), InlineKeyboardButton("📞 كل الارقام", callback_data=f"linecat_{line_key}_PHONE")],
        [InlineKeyboardButton("🏠 القايمة الرئيسية", callback_data="menu_main")],
    ])
def format_line_contacts(line_key, cat="ALL"):
    contacts = LINE_CONTACTS.get(line_key, []); display = LINE_NAMES.get(line_key, line_key)
    if not contacts: return f"❌ ملقيتش توكيل اسمه: {display}"
    txt = f"🚢 {display}\n"
    if cat == "BOOKING":
        txt += "📘 Booking:\n\n"; c = 0
        for x in contacts:
            if x["booking"] or x["email"]:
                txt += f"• {x['name'] or x['position'] or ''} - {x['email'] or ''} {x['number'] or ''}\nBooking: {x['booking']}\n\n"; c+=1
        if c==0: txt += "مفيش Booking مسجل\n"
        return txt.strip()
    if cat == "INVOICE":
        txt += "🧾 Invoice:\n\n"
        for x in contacts:
            if x["invoice"]: txt += f"• {x['name'] or x['position'] or ''} - {x['email']}\nInvoice: {x['invoice']}\n\n"
        return txt.strip() or f"مفيش Invoice لـ {display}"
    if cat == "PHONE":
        txt += "📞 الارقام:\n\n"
        for x in contacts:
            if x["number"]: txt += f"• {x['name'] or x['position'] or ''}: {x['number']}\n"
        return txt.strip()
    txt += f"📞 عدد جهات الاتصال: {len(contacts)}\n\n"
    for c in contacts[:12]:
        if c["name"]: txt += f"👤 {c['name']}\n"
        if c["position"]: txt += f"💼 {c['position']}\n"
        if c["email"]: txt += f"📧 {c['email']}\n"
        if c["number"]: txt += f"📞 {c['number']}\n"
        if c["booking"]: txt += f"📘 Booking: {c['booking']}\n"
        if c["invoice"]: txt += f"🧾 Invoice: {c['invoice']}\n"
        txt += "\n"
    if len(contacts) > 12: txt += f"... و {len(contacts)-12} تانيين\n"
    return txt.strip()
def copy_keyboard(line_key):
    contacts = LINE_CONTACTS.get(line_key, []); kb = []
    for i, c in enumerate(contacts[:5]):
        if c["email"]: kb.append([InlineKeyboardButton(f"📋 نسخ {c['email'][:30]}", callback_data=f"copy_{line_key}_{i}")])
    kb.append([InlineKeyboardButton("🏠 القايمة", callback_data="menu_main")])
    return InlineKeyboardMarkup(kb)
def search_rotation(pol, pod):
    pol_l = pol.lower().strip(); pod_l = pod.lower().strip(); out = []
    for r in ROTATION:
        if pol_l in r["pol"].lower() and pod_l in r["pod"].lower(): out.append(r)
    for r in SLMC_LIST:
        if pol_l in r["pol"].lower() and pod_l in r["pod"].lower():
            if r not in out: out.append(r)
    return out[:10]
def format_rotation(results, pol, pod):
    if not results: return f"❌ ملقيتش خط من {pol} لـ {pod}"
    txt = f"🚢 من {pol} لـ {pod} - لقيت {len(results)}:\n\n"
    for r in results:
        tt = f" - T.T: {r['tt']}" if r["tt"] else ""
        txt += f"• {r['line']} | {r['pol']} -> {r['pod']}{tt} | {r['country']}\n"
    return txt
def get_filtered_lines(query=""):
    lst = sorted(LINE_NAMES.keys(), key=lambda x: LINE_NAMES[x]); q = (query or "").strip().lower()
    if not q: return lst
    return [k for k in lst if q in LINE_NAMES[k].lower() or q in k.lower()]
def line_list_keyboard(page=0, query=""):
    lst = get_filtered_lines(query); per = 10; start = page*per; end = start+per; total_pages = max(1, (len(lst)+per-1)//per); kb = []
    for i in range(start, min(end, len(lst))):
        k = lst[i]; kb.append([InlineKeyboardButton(LINE_NAMES[k][:55], callback_data=f"linesel_{i}")])
    nav = []
    if page > 0: nav.append(InlineKeyboardButton("⬅️ السابق", callback_data=f"linepage_{page-1}"))
    if page < total_pages-1: nav.append(InlineKeyboardButton("التالي ➡️", callback_data=f"linepage_{page+1}"))
    if nav: kb.append(nav)
    kb.append([InlineKeyboardButton("🏠 القايمة الرئيسية", callback_data="menu_main")])
    title = f"بحث: {query} - " if query else ""
    return InlineKeyboardMarkup(kb), f"{title}اختار التوكيل - صفحة {page+1} من {total_pages} - العدد {len(lst)}"
async def send_line_list(update, context, page=0, query=""):
    uid = None
    if update.message: uid = update.message.from_user.id
    elif update.callback_query: uid = update.callback_query.from_user.id
    if uid is not None: context.bot_data[f"line_filter_{uid}"] = {"query": query}
    kb, txt = line_list_keyboard(page, query)
    if update.message: await update.message.reply_text(txt, reply_markup=kb)
    elif update.callback_query:
        try: await update.callback_query.message.edit_text(txt, reply_markup=kb)
        except: await update.callback_query.message.reply_text(txt, reply_markup=kb)

def load_docs():
    if not os.path.exists(DOCS_FILE): return []
    try:
        with open(DOCS_FILE, "r", encoding="utf-8") as f: return json.load(f)
    except: return []
def save_docs(docs):
    with open(DOCS_FILE, "w", encoding="utf-8") as f: json.dump(docs, f, ensure_ascii=False, indent=2)
WELCOME_MESSAGE = """Welcome to Captain Emad Customs Bot.

After the trial and subscription for 150 EGP, you can inquire about the customs item by code or by item name in Arabic or English.

You can also search for GS1, and you can calculate the approximate customs."""
HELP_TEXT = """❓ مساعدة / Help

🚢 توكيلات - اكتب: توكيلات او MAERSK او ميرسك او توكيل MAERSK
🚢 خط - اكتب: Alex Izmir او EGYPT LATTAKIA او MAERSK Alex Izmir
📦 GPC - اكتب: /gpc Airbrushes او اسم الصنف
🌍 Country - اكتب: Country Eg او Country Egypt
🏙️ Cities - اكتب: Cities Cai
💵 Currency - اكتب: Currency US
📦 Package - اكتب: Package B
📏 UOM - اكتب: UOM K
📁 مستندات - اكتب: /docs او تزيل 1
🕘 اخر بحث - اكتب: اخر
🏠 القايمة - اكتب: /menu"""
def load_users():
    if not os.path.exists(USERS_FILE): return {}
    try:
        with open(USERS_FILE, "r", encoding="utf-8") as f: return json.load(f)
    except: return {}
def save_users(users):
    with open(USERS_FILE, "w", encoding="utf-8") as f: json.dump(users, f, ensure_ascii=False, indent=2)
def generate_code(users):
    while True:
        code = "USR-" + "".join(random.choices(string.digits, k=4))
        if all(u.get("code") != code for u in users.values()): return code
def clean_text(text):
    if not text: return ""
    for ch in ["\u200e", "\u200f", "\u202a", "\u202b", "\u202c", "\u202d", "\u202e", "\u200c", "\u200d", "\u0640", "ـ"]: text = text.replace(ch, "")
    arabic_digits = "٠١٢٣٤٥٦٧٨٩"
    for i, d in enumerate(arabic_digits): text = text.replace(d, str(i))
    text = text.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا").replace("ة", "ه")
    return text.strip()
def is_download_request(text):
    t = clean_text(text); return bool(re.search(r'(تزيل|تنزيل|تحميل|download|نزل)', t))
def extract_number(text):
    t = clean_text(text); m = re.search(r'(\d+)', t); return int(m.group(1)) if m else None
def normalize(text): return re.sub(r'\s+', ' ', (text or "").strip().lower())
def search_gpc_fast(query):
    q = normalize(query)
    if not q: return None, None
    q_digits = re.sub(r'\D', '', q)
    if q_digits and q_digits in BRICK_BY_CODE:
        key = BRICK_BY_CODE[q_digits]; return BRICK_INDEX[key], None
    if q in BRICK_INDEX: return BRICK_INDEX[q], None
    matches = [BRICK_INDEX[k] for k in BRICK_INDEX.keys() if q in k]
    if len(matches) == 1: return matches[0], None
    if len(matches) > 1: return ("multiple", matches[:10]), None
    return None, None
def get_attrs_for_brick(brick_title): return BRICK_ATTRS.get(brick_title.strip().lower(), [])
def format_brick(brick, attrs):
    txt = f"📦 {brick['brick_title']}\n🔢 Brick Code: {brick['brick_code']}\n\n📂 Segment: {brick['segment']} ({brick['segment_code']})\n📁 Family: {brick['family']} ({brick['family_code']})\n📄 Class: {brick['class']} ({brick['class_code']})\n\n"
    if brick.get("includes"): txt += f"✅ Includes:\n{brick['includes']}\n\n"
    if brick.get("excludes"): txt += f"⛔ Excludes:\n{brick['excludes']}\n\n"
    if attrs:
        txt += "⚙️ Attributes:\n"
        for a in attrs:
            txt += f"\n• {a['attribute_title']} ({a['attribute_code']})\n"
            if a.get("values"):
                vals = a["values"][:15]; txt += "  " + ", ".join(vals)
                if len(a["values"]) > 15: txt += f" ... (+{len(a['values'])-15} more)"
                txt += "\n"
    return txt
def search_ref(kind, query):
    q = query.strip(); qu = q.upper().strip()
    if kind == "cities": by_code, index = CITY_BY_CODE, CITY_INDEX
    elif kind == "country": by_code, index = COUNTRY_BY_CODE, COUNTRY_INDEX
    elif kind == "currency": by_code, index = CURRENCY_BY_CODE, CURRENCY_INDEX
    elif kind == "package": by_code, index = PACKAGE_BY_CODE, PACKAGE_INDEX
    elif kind == "uom": by_code, index = UOM_BY_CODE, UOM_INDEX
    else: return None, None
    if qu in by_code: return by_code[qu], None
    ql = q.lower(); matches = []
    for key, lst in index.items():
        if ql in key:
            matches.extend(lst)
            if len(matches) >= 10: break
    if not matches:
        for code, item in by_code.items():
            if qu in code:
                matches.append(item)
                if len(matches) >= 10: break
    if len(matches) == 1: return matches[0], None
    if len(matches) > 1: return ("multiple", matches[:10]), None
    return None, None
def format_ref(kind, item):
    if kind == "cities":
        extra = f"\n🌍 Country: {item.get('country','')}" if item.get("country") else ""
        return f"🏙️ {item['name']}\n🔢 City Code: {item['code']}{extra}"
    if kind == "country": return f"🌍 {item['name']}\n🔢 Country Code: {item['code']}"
    if kind == "currency": return f"💵 {item['name']}\n🔢 Currency Code: {item['code']}"
    if kind == "package": return f"📦 {item['name']}\n🔢 Package Code: {item['code']}"
    if kind == "uom": return f"📏 {item['name']}\n🔢 UOM Code: {item['code']}"
    return f"{item['name']} - {item['code']}"
def search_all_refs(query):
    results = []; q = query.strip(); qu = q.upper().strip(); ql = q.lower()
    if qu in CITY_BY_CODE: results.append(("cities", CITY_BY_CODE[qu]))
    if qu in COUNTRY_BY_CODE: results.append(("country", COUNTRY_BY_CODE[qu]))
    if qu in CURRENCY_BY_CODE: results.append(("currency", CURRENCY_BY_CODE[qu]))
    if qu in PACKAGE_BY_CODE: results.append(("package", PACKAGE_BY_CODE[qu]))
    if qu in UOM_BY_CODE: results.append(("uom", UOM_BY_CODE[qu]))
    if results: return results
    for kind, index in [("cities", CITY_INDEX), ("country", COUNTRY_INDEX), ("currency", CURRENCY_INDEX), ("package", PACKAGE_INDEX), ("uom", UOM_INDEX)]:
        for key, lst in index.items():
            if ql in key:
                for item in lst:
                    results.append((kind, item))
                    if len(results) >= 10: return results
    return results
def get_sorted_ref_list(kind):
    if kind == "cities": return sorted(CITY_BY_CODE.values(), key=lambda x: x["code"])
    if kind == "country": return sorted(COUNTRY_BY_CODE.values(), key=lambda x: x["code"])
    if kind == "currency": return sorted(CURRENCY_BY_CODE.values(), key=lambda x: x["code"])
    if kind == "package": return sorted(PACKAGE_BY_CODE.values(), key=lambda x: x["code"])
    if kind == "uom": return sorted(UOM_BY_CODE.values(), key=lambda x: x["code"])
    return []
def get_filtered_ref_list(kind, query=""):
    lst = get_sorted_ref_list(kind); q = (query or "").strip().lower()
    if not q: return lst
    qu = q.upper(); out = []
    for item in lst:
        if q in item["name"].lower() or qu in item["code"]: out.append(item)
    return out
def ref_list_keyboard(kind, page=0, query=""):
    lst = get_filtered_ref_list(kind, query); per = 10; start = page*per; end = start+per; total_pages = max(1, (len(lst)+per-1)//per); kb = []
    for i in range(start, min(end, len(lst))):
        item = lst[i]; label = f"{item['name'][:40]} - {item['code']}"; kb.append([InlineKeyboardButton(label, callback_data=f"refsel_{kind}_{i}")])
    nav = []
    if page > 0: nav.append(InlineKeyboardButton("⬅️ السابق", callback_data=f"refpage_{kind}_{page-1}"))
    if page < total_pages-1: nav.append(InlineKeyboardButton("التالي ➡️", callback_data=f"refpage_{kind}_{page+1}"))
    if nav: kb.append(nav)
    kb.append([InlineKeyboardButton("🏠 القايمة الرئيسية", callback_data="menu_main")])
    title = f"بحث: {query} - " if query else ""
    return InlineKeyboardMarkup(kb), f"{title}اختار من {kind} - صفحة {page+1} من {total_pages} - العدد {len(lst)}"
async def send_ref_list(update, context, kind, page=0, query=""):
    uid = None
    if update.message: uid = update.message.from_user.id
    elif update.callback_query: uid = update.callback_query.from_user.id
    if uid is not None: context.bot_data[f"ref_filter_{uid}"] = {"kind": kind, "query": query}
    kb, txt = ref_list_keyboard(kind, page, query)
    if update.message: await update.message.reply_text(txt, reply_markup=kb)
    elif update.callback_query:
        try: await update.callback_query.message.edit_text(txt, reply_markup=kb)
        except: await update.callback_query.message.reply_text(txt, reply_markup=kb)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.message.from_user; users = load_users(); uid = str(user.id); now = time.time()
    if uid in users and users[uid].get("status") == "rejected":
        last = users[uid].get("rejected_at", 0); remaining = int((REJECT_WAIT - (now - last)) / 60) + 1
        if now - last < REJECT_WAIT:
            kb = InlineKeyboardMarkup([[InlineKeyboardButton("🔁 اطلب قبول تاني", callback_data="request_again")]])
            await update.message.reply_text(f"Your request was rejected. You can send again after {remaining} minutes.", reply_markup=kb); return
        users[uid]["status"] = "pending"; users[uid]["requested_at"] = now; save_users(users)
    if uid in users and users[uid].get("status") == "accepted":
        await update.message.reply_text(WELCOME_MESSAGE + "\n\n✅ Your account is accepted. Send the item code or item name.", reply_markup=main_menu_keyboard()); return
    if uid not in users:
        code = generate_code(users); users[uid] = {"code": code, "name": user.first_name or "", "username": user.username or "", "id": user.id, "status": "pending", "requested_at": now}; save_users(users)
    else:
        code = users[uid]["code"]; users[uid]["status"] = "pending"; users[uid]["requested_at"] = now; save_users(users)
    await update.message.reply_text(WELCOME_MESSAGE + f"\n\nYour Code: {code}\nWaiting for acceptance from Captain Emad.", reply_markup=main_menu_keyboard())
    if ADMIN_ID:
        keyboard = [[InlineKeyboardButton("✅ قبول / Accept", callback_data=f"accept_{uid}"), InlineKeyboardButton("❌ رفض / Reject", callback_data=f"reject_{uid}")]]
        try: await context.bot.send_message(chat_id=int(ADMIN_ID), text=f"📦 New User Request\n\nCode: {code}\nName: {user.first_name}\nUsername: @{user.username}\nID: {user.id}", reply_markup=InlineKeyboardMarkup(keyboard))
        except Exception as e: print(f"Admin send error: {e}")
async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(HELP_TEXT, reply_markup=main_menu_keyboard())
async def menu_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("🏠 القايمة الرئيسية - اختار:", reply_markup=main_menu_keyboard())
async def top_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if str(update.message.from_user.id) != str(ADMIN_ID): return
    h = load_hits(); lines = sorted(h["lines"].items(), key=lambda x: x[1], reverse=True)[:10]; ports = sorted(h["ports"].items(), key=lambda x: x[1], reverse=True)[:10]
    txt = "📊 اكتر حاجة اتسأل عليها:\n\n🚢 خطوط:\n"
    for k, v in lines: txt += f"• {k}: {v}\n"
    txt += "\n⚓ موانئ:\n"
    for k, v in ports: txt += f"• {k}: {v}\n"
    await update.message.reply_text(txt or "لسه مفيش بحث")

async def button_click(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    try: await query.answer()
    except: pass
    users = load_users(); data = query.data
    if data == "menu_main":
        await query.message.reply_text("🏠 القايمة الرئيسية - اختار:", reply_markup=main_menu_keyboard()); return
    if data == "menu_lines": await send_line_list(query, context, 0, ""); return
    if data == "menu_gpc": await query.message.reply_text("اكتب: /gpc Airbrushes او اسم الصنف", reply_markup=main_menu_keyboard()); return
    if data == "menu_country": await send_ref_list(query, context, "country", 0, ""); return
    if data == "menu_cities": await send_ref_list(query, context, "cities", 0, ""); return
    if data == "menu_currency": await send_ref_list(query, context, "currency", 0, ""); return
    if data == "menu_package": await send_ref_list(query, context, "package", 0, ""); return
    if data == "menu_uom": await send_ref_list(query, context, "uom", 0, ""); return
    if data == "menu_rotation": await query.message.reply_text("اكتب: Alex Izmir او EGYPT LATTAKIA", reply_markup=main_menu_keyboard()); return
    if data == "menu_docs":
        docs = load_docs()
        if not docs: await query.message.reply_text("مفيش مستندات لسه", reply_markup=main_menu_keyboard()); return
        txt = "📁 المستندات:\n\n"
        for d in docs: txt += f"{d['num']} - {d['name']}\n"
        txt += "\nاكتب: تزيل 1"
        await query.message.reply_text(txt, reply_markup=main_menu_keyboard()); return
    if data == "menu_help": await query.message.reply_text(HELP_TEXT, reply_markup=main_menu_keyboard()); return
    if data == "menu_recent":
        lst = context.user_data.get("recent", [])
        if not lst: await query.message.reply_text("لسه مفيش بحث"); return
        await query.message.reply_text("🕘 اخر 5:\n" + "\n".join([f"• {x}" for x in lst])); return
    if data == "request_again":
        uid = str(query.from_user.id); users2 = load_users()
        if uid in users2:
            users2[uid]["status"] = "pending"; users2[uid]["requested_at"] = time.time(); save_users(users2)
            await query.message.reply_text("✅ طلبك اتبعت تاني لكابتن عماد، استنى القبول")
            if ADMIN_ID: await context.bot.send_message(chat_id=int(ADMIN_ID), text=f"🔁 طلب قبول تاني\nCode: {users2[uid]['code']} - {users2[uid]['name']} - ID: {uid}")
        return
    if data.startswith("refpage_"):
        parts = data.split("_"); kind = parts[1]; page = int(parts[2]); filt = context.bot_data.get(f"ref_filter_{query.from_user.id}", {}); q = filt.get("query", ""); kb, txt = ref_list_keyboard(kind, page, q)
        try: await query.message.edit_text(txt, reply_markup=kb)
        except: await query.message.reply_text(txt, reply_markup=kb)
        return
    if data.startswith("refsel_"):
        parts = data.split("_"); kind = parts[1]; idx = int(parts[2]); filt = context.bot_data.get(f"ref_filter_{query.from_user.id}", {}); q = filt.get("query", ""); lst = get_filtered_ref_list(kind, q)
        if 0 <= idx < len(lst): await query.message.reply_text(format_ref(kind, lst[idx]))
        return
    if data.startswith("linepage_"):
        page = int(data.split("_")[1]); filt = context.bot_data.get(f"line_filter_{query.from_user.id}", {}); q = filt.get("query", ""); kb, txt = line_list_keyboard(page, q)
        try: await query.message.edit_text(txt, reply_markup=kb)
        except: await query.message.reply_text(txt, reply_markup=kb)
        return
    if data.startswith("linesel_"):
        idx = int(data.split("_")[1]); filt = context.bot_data.get(f"line_filter_{query.from_user.id}", {}); q = filt.get("query", ""); lst = get_filtered_lines(q)
        if 0 <= idx < len(lst):
            k = lst[idx]; hit_line(k); add_recent(context, LINE_NAMES[k])
            await query.message.reply_text(format_line_contacts(k), reply_markup=line_category_keyboard(k))
            await query.message.reply_text("📋 دوس عشان تنسخ:", reply_markup=copy_keyboard(k))
        return
    if data.startswith("linecat_"):
        parts = data.split("_"); line_key = parts[1]; cat = parts[2]
        await query.message.reply_text(format_line_contacts(line_key, cat), reply_markup=line_category_keyboard(line_key)); return
    if data.startswith("copy_"):
        parts = data.split("_"); line_key = parts[1]; idx = int(parts[2]); contacts = LINE_CONTACTS.get(line_key, [])
        if 0 <= idx < len(contacts):
            c = contacts[idx]; await query.message.reply_text(f"{c['email']}\n{c['number']}")
        return
    if data.startswith("accept_"):
        uid = data.split("_")[1]
        if uid in users:
            users[uid]["status"] = "accepted"; save_users(users)
            await context.bot.send_message(chat_id=int(uid), text=WELCOME_MESSAGE + "\n\n✅ Your request has been accepted by Captain Emad.", reply_markup=main_menu_keyboard())
            await query.message.reply_text(f"Accepted ✅\nCode: {users[uid]['code']} - {users[uid]['name']}")
    elif data.startswith("reject_"):
        uid = data.split("_")[1]
        if uid in users:
            users[uid]["status"] = "rejected"; users[uid]["rejected_at"] = time.time(); save_users(users)
            kb = InlineKeyboardMarkup([[InlineKeyboardButton("🔁 اطلب قبول تاني", callback_data="request_again")]])
            await context.bot.send_message(chat_id=int(uid), text="❌ Your request was rejected by Captain Emad. You can send again after 5 minutes.", reply_markup=kb)
            await query.message.reply_text(f"Rejected ❌\nCode: {users[uid]['code']} - {users[uid]['name']}")
    elif data.startswith("gpc_"):
        try:
            idx = int(data[4:]); keys = context.user_data.get("gpc_matches", [])
            if not keys and context.bot_data.get(f"gpc_{query.from_user.id}"): keys = context.bot_data.get(f"gpc_{query.from_user.id}", [])
            if 0 <= idx < len(keys):
                b = keys[idx]; attrs = get_attrs_for_brick(b["brick_title"]); await query.message.reply_text(format_brick(b, attrs))
            else: await query.message.reply_text("القايمة دي انتهت، ابعت /gpc تاني واختار من جديد")
        except Exception as e: print(f"gpc callback error: {e}")

async def save_doc(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if str(update.message.from_user.id) != str(ADMIN_ID): return
    caption = (update.message.caption or "").strip() or "بدون اسم"; file = None; ext = "jpg"
    if update.message.photo: file = await update.message.photo[-1].get_file()
    elif update.message.document:
        file = await update.message.document.get_file(); ext = update.message.document.file_name.split(".")[-1] if "." in update.message.document.file_name else "pdf"
    else: return
    docs = load_docs(); num = max([d["num"] for d in docs], default=0) + 1; filename = f"docs/{num}.{ext}"
    await file.download_to_drive(filename); docs.append({"num": num, "name": caption, "path": filename}); save_docs(docs)
    await update.message.reply_text(f"✅ اتحفظ برقم {num}\n📄 {caption}\n\nعشان تنزله اكتب: تزيل {num}")
async def list_docs(update: Update, context: ContextTypes.DEFAULT_TYPE):
    docs = load_docs()
    if not docs: await update.message.reply_text("مفيش مستندات لسه"); return
    txt = "📁 المستندات:\n\n"
    for d in docs: txt += f"{d['num']} - {d['name']}\n"
    txt += "\nاكتب: تزيل 1"; await update.message.reply_text(txt)
async def download_doc(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.message.from_user; uid = str(user.id)
    if str(user.id) != str(ADMIN_ID):
        users = load_users()
        if uid not in users or users[uid].get("status") != "accepted": await update.message.reply_text("Please send /start first."); return
    num = extract_number(update.message.text)
    if num is None: await update.message.reply_text("اكتب: تزيل 1"); return
    docs = load_docs()
    for d in docs:
        if d["num"] == num:
            try: await update.message.reply_document(document=open(d["path"], "rb"), caption=f"📄 رقم {num} - {d['name']}")
            except FileNotFoundError: await update.message.reply_text(f"❌ الملف رقم {num} مش موجود على السيرفر، ارفعه تاني")
            return
    await update.message.reply_text(f"❌ مفيش مستند رقم {num}\nاكتب /docs عشان تشوف الارقام")
async def gpc_search_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    users = load_users(); uid = str(update.message.from_user.id)
    if str(update.message.from_user.id) != str(ADMIN_ID):
        if uid not in users or users[uid].get("status") != "accepted": await update.message.reply_text("Please send /start first."); return
    if not context.args: await update.message.reply_text("اكتب: /gpc اسم الصنف او الكود\nمثال: /gpc Airbrushes\nاو: /gpc 10001686"); return
    query = " ".join(context.args); add_recent(context, query); brick, _ = search_gpc_fast(query)
    if brick is None: await update.message.reply_text(f"❌ ملقيتش صنف اسمه: {query}"); return
    if isinstance(brick, tuple) and brick[0] == "multiple":
        matches = brick[1]; context.user_data["gpc_matches"] = matches; context.bot_data[f"gpc_{update.message.from_user.id}"] = matches
        keyboard = [[InlineKeyboardButton(m["brick_title"][:60], callback_data=f"gpc_{i}")] for i, m in enumerate(matches)]
        await update.message.reply_text(f"لقيت {len(matches)} صنف، اختار واحد:", reply_markup=InlineKeyboardMarkup(keyboard)); return
    attrs = get_attrs_for_brick(brick["brick_title"]); await update.message.reply_text(format_brick(brick, attrs))
async def ref_command(update: Update, context: ContextTypes.DEFAULT_TYPE, kind, label, example):
    users = load_users(); uid = str(update.message.from_user.id)
    if str(update.message.from_user.id) != str(ADMIN_ID):
        if uid not in users or users[uid].get("status") != "accepted": await update.message.reply_text("Please send /start first."); return
    if not context.args: await send_ref_list(update, context, kind, 0, ""); return
    query = " ".join(context.args); add_recent(context, query); await send_ref_list(update, context, kind, 0, query); return
async def cities_command(update: Update, context: ContextTypes.DEFAULT_TYPE): await ref_command(update, context, "cities", "مدينة", "/cities Cairo")
async def country_command(update: Update, context: ContextTypes.DEFAULT_TYPE): await ref_command(update, context, "country", "دولة", "/country Egypt")
async def currency_command(update: Update, context: ContextTypes.DEFAULT_TYPE): await ref_command(update, context, "currency", "عملة", "/currency USD")
async def package_command(update: Update, context: ContextTypes.DEFAULT_TYPE): await ref_command(update, context, "package", "عبوة", "/package Box")
async def uom_command(update: Update, context: ContextTypes.DEFAULT_TYPE): await ref_command(update, context, "uom", "وحدة قياس", "/uom KGM")
async def lines_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    users = load_users(); uid = str(update.message.from_user.id)
    if str(update.message.from_user.id) != str(ADMIN_ID):
        if uid not in users or users[uid].get("status") != "accepted": await update.message.reply_text("Please send /start first."); return
    if not context.args: await send_line_list(update, context, 0, ""); return
    q = " ".join(context.args); add_recent(context, q); k = resolve_line(q)
    if k:
        hit_line(k); await update.message.reply_text(format_line_contacts(k), reply_markup=line_category_keyboard(k)); return
    sug = suggest_line(q)
    if sug: await update.message.reply_text(f"قصدك {LINE_NAMES.get(sug,sug)}؟\nاكتب: توكيل {LINE_NAMES.get(sug,sug)}"); return
    await send_line_list(update, context, 0, q); return
async def line_command(update: Update, context: ContextTypes.DEFAULT_TYPE): await lines_command(update, context)
async def rotation_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    users = load_users(); uid = str(update.message.from_user.id)
    if str(update.message.from_user.id) != str(ADMIN_ID):
        if uid not in users or users[uid].get("status") != "accepted": await update.message.reply_text("Please send /start first."); return
    if len(context.args) < 2: await update.message.reply_text("اكتب: /rotation Alex Izmir\nاو: Alex Izmir"); return
    pol = context.args[0]; pod = " ".join(context.args[1:]); hit_port(f"{pol}->{pod}"); add_recent(context, f"{pol} {pod}"); res = search_rotation(pol, pod); await update.message.reply_text(format_rotation(res, pol, pod))
async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.message.from_user; uid = str(user.id); text = update.message.text or ""
    if is_download_request(text): return await download_doc(update, context)
    if text.strip().lower().startswith("تعديل") and str(user.id) == str(ADMIN_ID):
        body = text.strip().split(" ", 1)[1] if " " in text.strip() else ""
        parts = [p.strip() for p in body.split("|")]
        if len(parts) >= 2:
            line_key = resolve_line(parts[0]) or parts[0].upper()
            new_c = {"line": line_key, "name": parts[1] if len(parts) > 1 else "", "position": "", "email": parts[2] if len(parts) > 2 else "", "number": parts[3] if len(parts) > 3 else "", "booking": "", "final": "", "vgm": "", "draft": "", "invoice": "", "pickup": ""}
            ov = {}
            if os.path.exists(OVERRIDE_FILE):
                try:
                    with open(OVERRIDE_FILE, "r", encoding="utf-8") as f: ov = json.load(f)
                except: ov = {}
            if line_key not in ov: ov[line_key] = []
            ov[line_key].append(new_c)
            with open(OVERRIDE_FILE, "w", encoding="utf-8") as f: json.dump(ov, f, ensure_ascii=False, indent=2)
            if line_key not in LINE_NAMES: LINE_NAMES[line_key] = line_key
            if line_key not in LINE_CONTACTS: LINE_CONTACTS[line_key] = []
            LINE_CONTACTS[line_key].append(new_c)
            await update.message.reply_text(f"✅ اتضاف لـ {line_key}\n👤 {new_c['name']} - {new_c['email']} - {new_c['number']}"); return
        await update.message.reply_text("اكتب: تعديل MAERSK | الاسم | الايميل | الرقم"); return
    bare = text.strip().lstrip("/").strip().lower()
    if bare in ["اخر", "recent", "اخر بحث"]:
        lst = context.user_data.get("recent", [])
        await update.message.reply_text("🕘 اخر 5:\n" + "\n".join([f"• {x}" for x in lst]) if lst else "لسه مفيش بحث"); return
    if bare in ["توكيلات", "توكيل", "lines", "line", "shipping lines", "خطوط", "خط"]:
        await send_line_list(update, context, 0, ""); return
    if bare in ["menu", "القايمة", "القايمه", "menu_main", "القائمة"]:
        await update.message.reply_text("🏠 القايمة الرئيسية:", reply_markup=main_menu_keyboard()); return
    low = text.strip().lower()
    parts_all = text.strip().split()
    if len(parts_all) >= 3:
        k = resolve_line(parts_all[0])
        if k:
            pol = parts_all[1]; pod = " ".join(parts_all[2:]); hit_line(k); hit_port(f"{pol}->{pod}"); add_recent(context, text)
            res = search_rotation(pol, pod)
            await update.message.reply_text(format_line_contacts(k), reply_markup=line_category_keyboard(k))
            await update.message.reply_text(format_rotation(res, pol, pod)); return
    for pref in ["توكيل ", "توكيلات ", "line ", "lines ", "shipping "]:
        if low.startswith(pref) or low.startswith("/"+pref):
            q = text.strip().lstrip("/").strip().split(" ", 1)[1] if " " in text.strip() else ""
            if q:
                add_recent(context, q); k = resolve_line(q)
                if k:
                    hit_line(k); await update.message.reply_text(format_line_contacts(k), reply_markup=line_category_keyboard(k)); return
                sug = suggest_line(q)
                if sug: await update.message.reply_text(f"قصدك {LINE_NAMES.get(sug,sug)}؟\nاكتب: توكيل {LINE_NAMES.get(sug,sug)}"); return
                await send_line_list(update, context, 0, q); return
    for pref in ["rotation ", "route ", "rot "]:
        if low.startswith(pref) or low.startswith("/"+pref):
            parts = text.strip().lstrip("/").strip().split(" ")
            if len(parts) >= 3:
                pol = parts[1]; pod = " ".join(parts[2:]); hit_port(f"{pol}->{pod}"); add_recent(context, text); res = search_rotation(pol, pod); await update.message.reply_text(format_rotation(res, pol, pod)); return
    k_single = resolve_line(text.strip())
    if k_single:
        hit_line(k_single); add_recent(context, text); await update.message.reply_text(format_line_contacts(k_single), reply_markup=line_category_keyboard(k_single)); return
    if len(text.strip().split()) == 1 and len(text.strip()) >= 3:
        sug = suggest_line(text.strip())
        if sug:
            await update.message.reply_text(f"قصدك {LINE_NAMES.get(sug,sug)}؟\nاكتب: توكيل {LINE_NAMES.get(sug,sug)}"); return
    parts = text.strip().split()
    if len(parts) == 2:
        res = search_rotation(parts[0], parts[1])
        if res:
            hit_port(f"{parts[0]}->{parts[1]}"); add_recent(context, text); await update.message.reply_text(format_rotation(res, parts[0], parts[1])); return
    bare_map = {"country": "country", "countries": "country", "دولة": "country", "دول": "country", "cities": "cities", "city": "cities", "مدينة": "cities", "مدن": "cities", "currency": "currency", "currencies": "currency", "عملة": "currency", "عملات": "currency", "package": "package", "packages": "package", "packaging": "package", "عبوة": "package", "عبوات": "package", "uom": "uom", "unit": "uom", "units": "uom", "وحدة": "uom"}
    if bare in bare_map: await send_ref_list(update, context, bare_map[bare], 0, ""); return
    prefixes = [("cities ", "cities"), ("city ", "cities"), ("مدينة ", "cities"), ("country ", "country"), ("دولة ", "country"), ("currency ", "currency"), ("عملة ", "currency"), ("package ", "package"), ("عبوة ", "package"), ("packaging ", "package"), ("uom ", "uom"), ("unit ", "uom")]
    for pref, kind in prefixes:
        if low.startswith(pref) or low.startswith("/" + pref):
            q = text.strip().lstrip("/").strip().split(" ", 1)[1] if " " in text.strip() else ""; add_recent(context, q); await send_ref_list(update, context, kind, 0, q); return
    q_single = text.strip().lstrip("/").strip()
    if len(q_single) >= 2:
        ref_results = search_all_refs(q_single)
        if ref_results:
            txt = ""
            for kind, item in ref_results: txt += format_ref(kind, item) + "\n\n"
            await update.message.reply_text(txt.strip()); return
    users = load_users()
    if str(user.id) == str(ADMIN_ID):
        if len(text.strip()) >= 3:
            brick, _ = search_gpc_fast(text)
            if brick is not None:
                if isinstance(brick, tuple) and brick[0] == "multiple":
                    matches = brick[1]; context.user_data["gpc_matches"] = matches; context.bot_data[f"gpc_{user.id}"] = matches
                    keyboard = [[InlineKeyboardButton(m["brick_title"][:60], callback_data=f"gpc_{i}")] for i, m in enumerate(matches)]
                    await update.message.reply_text(f"لقيت {len(matches)} صنف قريب من: {text}\nاختار واحد:", reply_markup=InlineKeyboardMarkup(keyboard)); return
                attrs = get_attrs_for_brick(brick["brick_title"]); await update.message.reply_text(format_brick(brick, attrs)); return
        return
    if uid not in users: await update.message.reply_text("Please send /start first.", reply_markup=main_menu_keyboard()); return
    if users[uid].get("status") == "pending": await update.message.reply_text("Your request is still waiting for acceptance from Captain Emad."); return
    if users[uid].get("status") == "rejected":
        kb = InlineKeyboardMarkup([[InlineKeyboardButton("🔁 اطلب قبول تاني", callback_data="request_again")]])
        await update.message.reply_text("Your request was rejected. Send /start again after 5 minutes.", reply_markup=kb); return
    if len(text.strip()) >= 3:
        brick, _ = search_gpc_fast(text)
        if brick is not None:
            add_recent(context, text)
            if isinstance(brick, tuple) and brick[0] == "multiple":
                matches = brick[1]; context.user_data["gpc_matches"] = matches; context.bot_data[f"gpc_{user.id}"] = matches
                keyboard = [[InlineKeyboardButton(m["brick_title"][:60], callback_data=f"gpc_{i}")] for i, m in enumerate(matches)]
                await update.message.reply_text(f"لقيت {len(matches)} صنف قريب من: {text}\nاختار واحد:", reply_markup=InlineKeyboardMarkup(keyboard)); return
            attrs = get_attrs_for_brick(brick["brick_title"]); await update.message.reply_text(format_brick(brick, attrs)); return
    if ADMIN_ID: await context.bot.send_message(chat_id=int(ADMIN_ID), text=f"Inquiry from {users[uid]['code']} - {users[uid]['name']}:\n{update.message.text}")
    await update.message.reply_text("Received. Captain Emad will reply to you.\n\nجرب القايمة دي، هتلاقي طلبك اسرع:", reply_markup=main_menu_keyboard())
async def list_users(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if str(update.message.from_user.id) != str(ADMIN_ID): return
    users = load_users()
    if not users: await update.message.reply_text("No users yet."); return
    text = f"Total Users: {len(users)}\n\n"
    for u in users.values(): text += f"Code: {u['code']} | Name: {u['name']} | @{u.get('username','')} | ID: {u['id']} | Status: {u['status']}\n"
    await update.message.reply_text(text)
async def delete_user(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if str(update.message.from_user.id) != str(ADMIN_ID): return
    if not context.args: await update.message.reply_text("Usage: /delete USR-1234  or  /delete <ID>"); return
    target = context.args[0]; users = load_users(); found = None
    for uid, u in list(users.items()):
        if u.get("code") == target or str(u.get("id")) == target: found = uid; break
    if not found: await update.message.reply_text("User not found."); return
    del users[found]; save_users(users); await update.message.reply_text(f"Deleted user {target} ✅")
async def stats(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if str(update.message.from_user.id) != str(ADMIN_ID): return
    users = load_users(); accepted = sum(1 for u in users.values() if u.get("status") == "accepted"); pending = sum(1 for u in users.values() if u.get("status") == "pending"); rejected = sum(1 for u in users.values() if u.get("status") == "rejected")
    await update.message.reply_text(f"Total: {len(users)}\nAccepted: {accepted}\nPending: {pending}\nRejected: {rejected}")
def main():
    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("menu", menu_command))
    app.add_handler(CommandHandler("top", top_command))
    app.add_handler(CommandHandler("users", list_users))
    app.add_handler(CommandHandler("delete", delete_user))
    app.add_handler(CommandHandler("stats", stats))
    app.add_handler(CommandHandler("docs", list_docs))
    app.add_handler(CommandHandler("gpc", gpc_search_command))
    app.add_handler(CommandHandler("cities", cities_command))
    app.add_handler(CommandHandler("city", cities_command))
    app.add_handler(CommandHandler("country", country_command))
    app.add_handler(CommandHandler("currency", currency_command))
    app.add_handler(CommandHandler("package", package_command))
    app.add_handler(CommandHandler("uom", uom_command))
    app.add_handler(CommandHandler("lines", lines_command))
    app.add_handler(CommandHandler("line", line_command))
    app.add_handler(CommandHandler("rotation", rotation_command))
    app.add_handler(CommandHandler("route", rotation_command))
    app.add_handler(CallbackQueryHandler(button_click))
    app.add_handler(MessageHandler(filters.PHOTO | filters.Document.ALL, save_doc))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    print(f"Bot is running... {len(BRICK_INDEX)} bricks, {len(CITY_BY_CODE)} cities, {len(COUNTRY_BY_CODE)} countries, {len(CURRENCY_BY_CODE)} currencies, {len(PACKAGE_BY_CODE)} packages, {len(UOM_BY_CODE)} uom, {len(LINE_NAMES)} lines, {len(ROTATION)} rotation, {len(SLMC_LIST)} slmc")
    app.run_polling()
if __name__ == "__main__":
    main()
