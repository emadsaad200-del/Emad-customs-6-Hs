import os
import json
import time
import random
import string
import re
import difflib
import openpyxl
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, CallbackQueryHandler, ContextTypes, filters
BOT_TOKEN = os.getenv("BOT_TOKEN")
ADMIN_ID = os.getenv("ADMIN_ID")
SLASH = chr(47)
USERS_FILE = "users.json"
REJECT_WAIT = 5 * 60
DOCS_FILE = "docs_index.json"
STATS_FILE = "hits.json"
OVERRIDE_FILE = "lines_override.json"
os.makedirs("docs", exist_ok=True)
EXCEL_FILES = ["gs1xlsx.xlsx", "GPC as of May 2023 v20230605 EN.xlsx"]
BRICK_INDEX = {}
BRICK_BY_CODE = {}
BRICK_ATTRS = {}
def build_index():
    print("Building light index...")
    for fname in EXCEL_FILES:
        if not os.path.exists(fname):
            continue
        try:
            wb = openpyxl.load_workbook(fname, read_only=True, data_only=True)
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
                    brick_code = str(r[9]) if r[9] is not None else ""
                    brick_title = str(r[10]) if r[10] is not None else ""
                    if not brick_title:
                        continue
                    key = brick_title.strip().lower()
                    if key not in BRICK_INDEX:
                        BRICK_INDEX[key] = {"brick_code": brick_code, "brick_title": brick_title, "segment": r[1], "segment_code": str(r[0]), "family": r[4], "family_code": str(r[3]), "class": r[7], "class_code": str(r[6]), "includes": r[11] or "", "excludes": r[12] if len(r) > 12 else ""}
                        if brick_code:
                            BRICK_BY_CODE[brick_code] = key
                    if len(r) > 14 and r[13] is not None:
                        attr_code = str(r[13])
                        attr_title = r[14]
                        val_title = str(r[17]) if len(r) > 17 and r[17] is not None else None
                        if key not in BRICK_ATTRS:
                            BRICK_ATTRS[key] = []
                        ex = None
                        for a in BRICK_ATTRS[key]:
                            if a["attribute_code"] == attr_code:
                                ex = a
                                break
                        if not ex:
                            ex = {"attribute_code": attr_code, "attribute_title": attr_title, "values": []}
                            BRICK_ATTRS[key].append(ex)
                        if val_title and val_title not in ex["values"]:
                            ex["values"].append(val_title)
                except:
                    continue
            wb.close()
            print("Indexed " + fname)
        except Exception as e:
            print("Index error " + fname + ": " + str(e))
    print("Index done")
build_index()
CITIES_FILE = "cities.xlsx"
CITY_INDEX = {}
CITY_BY_CODE = {}
COUNTRY_INDEX = {}
COUNTRY_BY_CODE = {}
CURRENCY_INDEX = {}
CURRENCY_BY_CODE = {}
PACKAGE_INDEX = {}
PACKAGE_BY_CODE = {}
UOM_INDEX = {}
UOM_BY_CODE = {}
def build_reference_index():
    if not os.path.exists(CITIES_FILE):
        print("cities.xlsx not found")
        return
    try:
        wb = openpyxl.load_workbook(CITIES_FILE, read_only=True, data_only=True)
        if "City List" in wb.sheetnames:
            ws = wb["City List"]
            for r in ws.iter_rows(values_only=True):
                try:
                    if not r or len(r) < 2 or not r[0] or not r[1]:
                        continue
                    if str(r[0]).lower() == "citycode":
                        continue
                    code = str(r[0]).strip().upper()
                    name = str(r[1]).strip()
                    country = str(r[2]).strip() if len(r) > 2 and r[2] else ""
                    CITY_BY_CODE[code] = {"code": code, "name": name, "country": country}
                    k = name.lower()
                    if k not in CITY_INDEX:
                        CITY_INDEX[k] = []
                    CITY_INDEX[k].append({"code": code, "name": name, "country": country})
                except:
                    continue
        if "Cities" in wb.sheetnames:
            ws = wb["Cities"]
            for r in ws.iter_rows(values_only=True):
                try:
                    if not r or len(r) < 2 or not r[0] or not r[1]:
                        continue
                    if str(r[0]).lower() == "citycode":
                        continue
                    code = str(r[0]).strip().upper()
                    name = str(r[1]).strip()
                    if code in CITY_BY_CODE:
                        continue
                    CITY_BY_CODE[code] = {"code": code, "name": name, "country": ""}
                    k = name.lower()
                    if k not in CITY_INDEX:
                        CITY_INDEX[k] = []
                    CITY_INDEX[k].append({"code": code, "name": name, "country": ""})
                except:
                    continue
        if "Country List" in wb.sheetnames:
            ws = wb["Country List"]
            for r in ws.iter_rows(values_only=True):
                try:
                    if not r or len(r) < 2 or not r[0] or not r[1]:
                        continue
                    if str(r[0]).lower() == "code":
                        continue
                    code = str(r[0]).strip().upper()
                    name = str(r[1]).strip()
                    COUNTRY_BY_CODE[code] = {"code": code, "name": name}
                    k = name.lower()
                    if k not in COUNTRY_INDEX:
                        COUNTRY_INDEX[k] = []
                    COUNTRY_INDEX[k].append({"code": code, "name": name})
                except:
                    continue
        if "Currency List" in wb.sheetnames:
            ws = wb["Currency List"]
            for r in ws.iter_rows(values_only=True):
                try:
                    if not r or len(r) < 2 or not r[0] or not r[1]:
                        continue
                    if "currency" in str(r[0]).lower():
                        continue
                    name = str(r[0]).strip()
                    code = str(r[1]).strip().upper()
                    CURRENCY_BY_CODE[code] = {"code": code, "name": name}
                    k = name.lower()
                    if k not in CURRENCY_INDEX:
                        CURRENCY_INDEX[k] = []
                    CURRENCY_INDEX[k].append({"code": code, "name": name})
                except:
                    continue
        if "Package Type list" in wb.sheetnames:
            ws = wb["Package Type list"]
            for r in ws.iter_rows(values_only=True):
                try:
                    if not r or len(r) < 2 or not r[0] or not r[1]:
                        continue
                    if str(r[0]).lower() == "code":
                        continue
                    code = str(r[0]).strip().upper()
                    name = str(r[1]).strip()
                    PACKAGE_BY_CODE[code] = {"code": code, "name": name}
                    k = name.lower()
                    if k not in PACKAGE_INDEX:
                        PACKAGE_INDEX[k] = []
                    PACKAGE_INDEX[k].append({"code": code, "name": name})
                except:
                    continue
        if "UOM List" in wb.sheetnames:
            ws = wb["UOM List"]
            for r in ws.iter_rows(values_only=True):
                try:
                    if not r or len(r) < 2 or not r[0] or not r[1]:
                        continue
                    if "uomcode" in str(r[0]).lower():
                        continue
                    code = str(r[0]).strip().upper()
                    name = str(r[1]).strip()
                    UOM_BY_CODE[code] = {"code": code, "name": name}
                    k = name.lower()
                    if k not in UOM_INDEX:
                        UOM_INDEX[k] = []
                    UOM_INDEX[k].append({"code": code, "name": name})
                except:
                    continue
        wb.close()
        print("Ref indexed")
    except Exception as e:
        print("Ref index error: " + str(e))
build_reference_index()
COMPANIES_FILE = "ارقام_تسجيل_الشركات_111111.xlsx"
COMPANIES_FILE_2026 = "ارقام_تسجيل_الشركات_2026.xlsx"
COMPANY_LIST = []
COMPANY_BY_REG = {}
COMPANY_BY_FILE = {}
COMPANY_NORM_CACHE = {}
def clean_cell(v):
    if v is None:
        return ""
    return str(v).strip()
def clean_invisible(s):
    if not s:
        return ""
    out = []
    for ch in str(s):
        o = ord(ch)
        if o < 32:
            continue
        if o == 160:
            continue
        if o == 1600:
            continue
        if o == 8204:
            continue
        if o == 8205:
            continue
        if o == 8206:
            continue
        if o == 8207:
            continue
        if o >= 8234 and o <= 8238:
            continue
        if o >= 1611 and o <= 1626:
            continue
        if o == 1648:
            continue
        out.append(ch)
    return "".join(out)
def norm_ar(s):
    if not s:
        return ""
    s = clean_invisible(str(s)).strip().lower()
    s = s.replace("أ", "ا")
    s = s.replace("إ", "ا")
    s = s.replace("آ", "ا")
    s = s.replace("ة", "ه")
    s = s.replace("ى", "ي")
    s = s.replace("ئ", "ي")
    s = s.replace("ؤ", "و")
    s = s.replace("ھ", "ه")
    s = s.replace("ی", "ي")
    s = s.replace("ڪ", "ك")
    s = s.replace("ک", "ك")
    s = s.translate(str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789"))
    s = re.sub(r"\s+", " ", s)
    return s.strip()
def clean_search_words(text):
    stop = ["ال", "شركه", "شركة", "لل", "ل", "شرك"]
    words = norm_ar(text).split()
    out = []
    for w in words:
        if w not in stop and len(w) >= 2:
            out.append(w)
    if out:
        return out
    return words
def _add_company(name, file_no, reg_no, entity, activity, status, address, ctype=""):
    name = clean_cell(name)
    reg_no = clean_cell(reg_no).replace(" ", "").replace(",", "")
    if not name or not reg_no:
        return
    if reg_no in COMPANY_BY_REG:
        return
    file_no = clean_cell(file_no).replace(" ", "").replace(",", "")
    item = {"name": name, "file_no": file_no, "reg_no": reg_no, "entity": clean_cell(entity), "activity": clean_cell(activity), "status": clean_cell(status), "address": clean_cell(address), "ctype": ctype}
    COMPANY_LIST.append(item)
    COMPANY_BY_REG[reg_no] = item
    if file_no:
        COMPANY_BY_FILE[file_no] = item
    idx = len(COMPANY_LIST) - 1
    COMPANY_NORM_CACHE[idx] = {"name": norm_ar(name), "activity": norm_ar(activity), "address": norm_ar(address), "entity": norm_ar(entity)}
def build_companies_index():
    if os.path.exists(COMPANIES_FILE):
        try:
            wb = openpyxl.load_workbook(COMPANIES_FILE, read_only=True, data_only=True)
            ws = wb[wb.sheetnames[0]]
            for r in ws.iter_rows(values_only=True):
                try:
                    if not r:
                        continue
                    name = clean_cell(r[0]) if len(r) > 0 else ""
                    if not name or "الاسم" in name:
                        continue
                    f1 = clean_cell(r[1]) if len(r) > 1 else ""
                    f2 = clean_cell(r[2]) if len(r) > 2 else ""
                    f3 = clean_cell(r[3]) if len(r) > 3 else ""
                    f4 = clean_cell(r[4]) if len(r) > 4 else ""
                    f5 = clean_cell(r[5]) if len(r) > 5 else ""
                    f6 = clean_cell(r[6]) if len(r) > 6 else ""
                    _add_company(name, f1, f2, f3, f4, f5, f6, "private")
                except:
                    continue
            wb.close()
            print("Companies OLD: " + str(len(COMPANY_LIST)))
        except Exception as e:
            print("Companies OLD error: " + str(e))
    if not os.path.exists(COMPANIES_FILE_2026):
        print("2026 file not found")
        return
    try:
        wb = openpyxl.load_workbook(COMPANIES_FILE_2026, read_only=True, data_only=True)
        if "خاص" in wb.sheetnames:
            ws = wb["خاص"]
            for r in ws.iter_rows(values_only=True):
                try:
                    if not r or not r[0] or not r[1]:
                        continue
                    if str(r[0]).lower() == "registrationnumber":
                        continue
                    ent = "خاص - " + clean_cell(r[3])
                    addr = "تاريخ التسجيل: " + clean_cell(r[4])
                    _add_company(r[1], "", r[0], ent, "", clean_cell(r[5]), addr, "private")
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
                    _add_company(r[2], "", r[1], "حكومي", "", "", clean_cell(r[3]), "gov")
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
                    _add_company(r[2], "", r[1], "فرع اجنبي", "", "", clean_cell(r[3]), "foreign")
                except:
                    continue
        wb.close()
        print("Companies TOTAL: " + str(len(COMPANY_LIST)))
    except Exception as e:
        print("Companies 2026 error: " + str(e))
build_companies_index()
def search_company(q, ctype=""):
    qq = (q or "").strip()
    if not qq:
        return []
    if qq in COMPANY_BY_REG:
        c = COMPANY_BY_REG[qq]
        if not ctype or c.get("ctype") == ctype:
            return [c]
        return []
    if qq in COMPANY_BY_FILE:
        c = COMPANY_BY_FILE[qq]
        if not ctype or c.get("ctype") == ctype:
            return [c]
        return []
    qq_digits = qq.translate(str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")).replace(" ", "").replace(",", "")
    if qq_digits.isdigit() and len(qq_digits) >= 3:
        out = []
        for c in COMPANY_LIST:
            if (qq_digits in c["reg_no"] or qq_digits in c["file_no"]) and (not ctype or c.get("ctype") == ctype):
                out.append(c)
        if out:
            return out[:10]
    nq = norm_ar(qq)
    words = clean_search_words(qq)
    out = []
    for idx, c in enumerate(COMPANY_LIST):
        if ctype and c.get("ctype") != ctype:
            continue
        nc = COMPANY_NORM_CACHE.get(idx, {}).get("name", norm_ar(c["name"]))
        if words and all(w in nc for w in words):
            out.append(c)
        elif nq and nq in nc:
            out.append(c)
        if len(out) >= 10:
            break
    if out:
        return out
    out2 = []
    for idx, c in enumerate(COMPANY_LIST):
        if ctype and c.get("ctype") != ctype:
            continue
        cache = COMPANY_NORM_CACHE.get(idx, {})
        if nq and (nq in cache.get("activity", "") or nq in cache.get("address", "") or nq in cache.get("entity", "")):
            out2.append(c)
        if len(out2) >= 10:
            break
    if out2:
        return out2
    try:
        names_norm = []
        for i in range(len(COMPANY_LIST)):
            if not ctype or COMPANY_LIST[i].get("ctype") == ctype:
                names_norm.append(COMPANY_NORM_CACHE.get(i, {}).get("name", ""))
        close = difflib.get_close_matches(nq, names_norm, n=10, cutoff=0.6)
        if close:
            out3 = []
            for cn in close:
                for idx, c in enumerate(COMPANY_LIST):
                    if COMPANY_NORM_CACHE.get(idx, {}).get("name", "") == cn:
                        out3.append(c)
                        break
            if out3:
                return out3[:10]
    except:
        pass
    return out
def format_company(c):
    txt = "🏢 " + c["name"] + "\n"
    if c["reg_no"]:
        txt += "🔢 رقم التسجيل: " + c["reg_no"] + "\n"
    if c["file_no"]:
        txt += "📁 رقم الملف: " + c["file_no"] + "\n"
    if c["entity"]:
        txt += "🏛️ الكيان: " + c["entity"] + "\n"
    if c["activity"]:
        txt += "📦 النشاط: " + c["activity"] + "\n"
    if c["status"]:
        txt += "📄 حاله الملف: " + c["status"] + "\n"
    if c["address"]:
        txt += "📍 العنوان: " + c["address"] + "\n"
    return txt.strip()
def get_filtered_companies(query="", ctype=""):
    base = []
    for c in COMPANY_LIST:
        if not ctype or c.get("ctype") == ctype:
            base.append(c)
    q = norm_ar(query)
    if not q:
        return base
    qd = query.strip().translate(str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")).replace(" ", "")
    if qd.isdigit() and len(qd) >= 3:
        out = []
        for c in base:
            if qd in c["reg_no"] or qd in c["file_no"]:
                out.append(c)
        return out
    words = clean_search_words(query)
    out = []
    for c in base:
        nc = norm_ar(c["name"])
        if words and all(w in nc for w in words):
            out.append(c)
        elif q in nc or q in norm_ar(c.get("activity", "")) or q in norm_ar(c.get("address", "")) or q in norm_ar(c.get("entity", "")):
            out.append(c)
    return out
def companies_submenu_keyboard():
    kb = []
    kb.append([InlineKeyboardButton("🏢 الشركات الخاصة", callback_data="comp_type_private")])
    kb.append([InlineKeyboardButton("🏛️ الشركات الحكومية", callback_data="comp_type_gov")])
    kb.append([InlineKeyboardButton("🌍 الشركات الاجنبية", callback_data="comp_type_foreign")])
    kb.append([InlineKeyboardButton("🏠 القايمة الرئيسية", callback_data="menu_main")])
    return InlineKeyboardMarkup(kb)
async def send_companies_submenu(update, context):
    txt = "🏢 ارقام تسجيل الشركات - اختار النوع:"
    if update.message:
        await update.message.reply_text(txt, reply_markup=companies_submenu_keyboard())
    elif update.callback_query:
        try:
            await update.callback_query.message.edit_text(txt, reply_markup=companies_submenu_keyboard())
        except:
            await update.callback_query.message.reply_text(txt, reply_markup=companies_submenu_keyboard())
def company_list_keyboard(page=0, query="", ctype=""):
    lst = get_filtered_companies(query, ctype)
    per = 10
    start = page * per
    end = start + per
    total_pages = max(1, (len(lst) + per - 1) // per)
    kb = []
    for i in range(start, min(end, len(lst))):
        c = lst[i]
        btn = c["name"][:50] + " - " + c["reg_no"]
        kb.append([InlineKeyboardButton(btn, callback_data="compsel_" + str(i))])
    nav = []
    if page > 0:
        nav.append(InlineKeyboardButton("⬅️ السابق", callback_data="comppage_" + str(page - 1)))
    if page < total_pages - 1:
        nav.append(InlineKeyboardButton("التالي ➡️", callback_data="comppage_" + str(page + 1)))
    if nav:
        kb.append(nav)
    kb.append([InlineKeyboardButton("🔙 انواع الشركات", callback_data="menu_companies")])
    kb.append([InlineKeyboardButton("🏠 القايمة الرئيسية", callback_data="menu_main")])
    title = "بحث: " + query + " - " if query else ""
    tname = ""
    if ctype == "private":
        tname = "الشركات الخاصة - "
    if ctype == "gov":
        tname = "الشركات الحكومية - "
    if ctype == "foreign":
        tname = "الشركات الاجنبية - "
    txt = tname + title + "اختار الشركة - صفحة " + str(page + 1) + " من " + str(total_pages) + " - العدد " + str(len(lst))
    return InlineKeyboardMarkup(kb), txt
async def send_company_list(update, context, page=0, query="", ctype=""):
    uid = None
    if update.message:
        uid = update.message.from_user.id
    elif update.callback_query:
        uid = update.callback_query.from_user.id
    if uid is not None:
        context.bot_data["comp_filter_" + str(uid)] = {"query": query, "ctype": ctype}
    kb, txt = company_list_keyboard(page, query, ctype)
    if update.message:
        await update.message.reply_text(txt, reply_markup=kb)
    elif update.callback_query:
        try:
            await update.callback_query.message.edit_text(txt, reply_markup=kb)
        except:
            await update.callback_query.message.reply_text(txt, reply_markup=kb)
TARIFF_LIST = []
TARIFF_BY_CODE = {}
def norm_tariff_code(s):
    if s is None:
        return ""
    s = str(s).strip().replace(" ", "").replace(",", "")
    if re.match(r"^\d{6}$", s):
        s = s[:4] + "." + s[4:]
    if re.match(r"^\d{8}$", s):
        s = s[:4] + "." + s[4:6] + "." + s[6:]
    return s
def clean_tariff_tax(v):
    if v is None:
        return ""
    s = str(v).strip()
    if s == "" or s == "-" or s == "None":
        return ""
    if s == "معفاه" or s == "معفاة":
        return "معفاه"
    tmp = s.replace("%", "").strip()
    if tmp:
        return tmp + "%"
    return ""
def build_tariff_index():
    files = []
    if os.path.isdir("tarifa"):
        for f in sorted(os.listdir("tarifa")):
            if f.lower().endswith(".xlsx"):
                files.append(os.path.join("tarifa", f))
    for f in sorted(os.listdir(".")):
        if f.lower().startswith("tarifa") and f.lower().endswith(".xlsx") and f not in files:
            files.append(f)
    for f in ["30285-1.xlsx", "70490-2.xlsx"]:
        if os.path.exists(f) and f not in files:
            files.append(f)
    if not files:
        print("Tariff: no files")
        return
    for fname in files:
        try:
            wb = openpyxl.load_workbook(fname, read_only=True, data_only=True)
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
                    first = str(r[0]).strip() if r[0] is not None else ""
                    if "رمز" in first or "الفصل" in first:
                        continue
                    code = ""
                    desc = ""
                    tax = ""
                    chapter = ""
                    main_band = ""
                    if len(r) >= 5 and r[2] is not None and "رمز" not in str(r[2]):
                        code = norm_tariff_code(r[2])
                        desc = clean_cell(r[3])
                        tax = clean_tariff_tax(r[4])
                        chapter = clean_cell(r[0])
                        main_band = clean_cell(r[1])
                    elif len(r) >= 3 and r[0] is not None:
                        code = norm_tariff_code(r[0])
                        desc = clean_cell(r[1])
                        tax = clean_tariff_tax(r[2])
                        if "." in code:
                            chapter = code.split(".")[0]
                        else:
                            chapter = code[:2]
                        main_band = code
                    else:
                        continue
                    if not code or code == "-":
                        continue
                    if not desc:
                        continue
                    item = {"code": code, "desc": desc, "tax": tax, "chapter": chapter, "main_band": main_band, "file": fname}
                    if code in TARIFF_BY_CODE:
                        if not TARIFF_BY_CODE[code].get("tax") and tax:
                            TARIFF_BY_CODE[code] = item
                        continue
                    TARIFF_BY_CODE[code] = item
                    TARIFF_LIST.append(item)
                except:
                    continue
            wb.close()
            print("Tariff indexed " + fname)
        except Exception as e:
            print("Tariff index error " + fname + ": " + str(e))
    print("Tariff done: " + str(len(TARIFF_LIST)))
build_tariff_index()
def search_tariff(q):
    qq = (q or "").strip()
    if not qq:
        return []
    code_q = norm_tariff_code(qq)
    if code_q in TARIFF_BY_CODE:
        return [TARIFF_BY_CODE[code_q]]
    q_digits = re.sub(r"\D", "", qq)
    if q_digits and len(q_digits) >= 2:
        out = []
        for t in TARIFF_LIST:
            if q_digits in re.sub(r"\D", "", t["code"]):
                out.append(t)
        if out:
            return out[:10]
    nq = norm_ar(qq)
    if len(nq) < 2:
        return []
    out = []
    for t in TARIFF_LIST:
        if nq in norm_ar(t["desc"]):
            out.append(t)
            if len(out) >= 10:
                break
    return out
def format_tariff(t):
    txt = "🧾 " + t["desc"] + "\n"
    txt += "🔢 رمز النظام المنسق / البند: " + t["code"] + "\n"
    if t.get("chapter"):
        txt += "📂 الفصل: " + t["chapter"] + "\n"
    if t.get("main_band") and t["main_band"] != t["code"]:
        txt += "📁 البند الرئيسي: " + t["main_band"] + "\n"
    taxval = t["tax"] if t.get("tax") else "- (بند رئيسي / عنوان)"
    txt += "💰 فئة الضريبة: " + taxval + "\n"
    return txt.strip()
def get_filtered_tariff(query=""):
    q = norm_ar(query)
    q_digits = re.sub(r"\D", "", query or "")
    if not q:
        return TARIFF_LIST
    out = []
    for t in TARIFF_LIST:
        if q in norm_ar(t["desc"]):
            out.append(t)
        elif q_digits and len(q_digits) >= 2 and q_digits in re.sub(r"\D", "", t["code"]):
            out.append(t)
    return out
def tariff_list_keyboard(page=0, query=""):
    lst = get_filtered_tariff(query)
    per = 10
    start = page * per
    end = start + per
    total_pages = max(1, (len(lst) + per - 1) // per)
    kb = []
    for i in range(start, min(end, len(lst))):
        t = lst[i]
        btn = t["code"] + " - " + t["desc"][:45]
        kb.append([InlineKeyboardButton(btn, callback_data="tarifsel_" + str(i))])
    nav = []
    if page > 0:
        nav.append(InlineKeyboardButton("⬅️ السابق", callback_data="tarifpage_" + str(page - 1)))
    if page < total_pages - 1:
        nav.append(InlineKeyboardButton("التالي ➡️", callback_data="tarifpage_" + str(page + 1)))
    if nav:
        kb.append(nav)
    kb.append([InlineKeyboardButton("🏠 القايمة الرئيسية", callback_data="menu_main")])
    title = "بحث: " + query + " - " if query else ""
    txt = title + "اختار البند - صفحة " + str(page + 1) + " من " + str(total_pages) + " - العدد " + str(len(lst))
    return InlineKeyboardMarkup(kb), txt
async def send_tariff_list(update, context, page=0, query=""):
    uid = None
    if update.message:
        uid = update.message.from_user.id
    elif update.callback_query:
        uid = update.callback_query.from_user.id
    if uid is not None:
        context.bot_data["tariff_filter_" + str(uid)] = {"query": query}
    kb, txt = tariff_list_keyboard(page, query)
    if update.message:
        await update.message.reply_text(txt, reply_markup=kb)
    elif update.callback_query:
        try:
            await update.callback_query.message.edit_text(txt, reply_markup=kb)
        except:
            await update.callback_query.message.reply_text(txt, reply_markup=kb)
SHIPPING_FILE = "ايميلات_وارقام_توكيلات.xlsx"
LINE_CONTACTS = {}
LINE_NAMES = {}
ROTATION = []
SLMC_LIST = []
def build_shipping_index():
    if not os.path.exists(SHIPPING_FILE):
        print("shipping file not found")
        return
    try:
        wb = openpyxl.load_workbook(SHIPPING_FILE, read_only=True, data_only=True)
        if "EMAILS" in wb.sheetnames:
            ws = wb["EMAILS"]
            rows = list(ws.iter_rows(values_only=True))
            cur_line = ""
            for r in rows[2:]:
                try:
                    if not r:
                        continue
                    if r[0]:
                        cur_line = str(r[0]).split("\n")[0].strip()
                    if not cur_line:
                        continue
                    key = cur_line.upper()
                    if key not in LINE_NAMES:
                        LINE_NAMES[key] = cur_line
                    name = clean_cell(r[1]) if len(r) > 1 else ""
                    pos = clean_cell(r[2]) if len(r) > 2 else ""
                    email = clean_cell(r[3]) if len(r) > 3 else ""
                    num = clean_cell(r[4]) if len(r) > 4 else ""
                    booking = clean_cell(r[5]) if len(r) > 5 else ""
                    final = clean_cell(r[6]) if len(r) > 6 else ""
                    vgm = clean_cell(r[7]) if len(r) > 7 else ""
                    draft = clean_cell(r[8]) if len(r) > 8 else ""
                    invoice = clean_cell(r[9]) if len(r) > 9 else ""
                    pickup = clean_cell(r[10]) if len(r) > 10 else ""
                    vals = [name, pos, email, num, booking, final, vgm, draft, invoice, pickup]
                    if not any(vals):
                        continue
                    if key not in LINE_CONTACTS:
                        LINE_CONTACTS[key] = []
                    LINE_CONTACTS[key].append({"line": cur_line, "name": name, "position": pos, "email": email, "number": num, "booking": booking, "final": final, "vgm": vgm, "draft": draft, "invoice": invoice, "pickup": pickup})
                except:
                    continue
        if "SERVICE ROTATION" in wb.sheetnames:
            ws = wb["SERVICE ROTATION"]
            last_pol = ""
            for r in ws.iter_rows(values_only=True):
                try:
                    if not r:
                        continue
                    if str(r[0]).upper() == "POL":
                        continue
                    pol = clean_cell(r[0]) or last_pol
                    if pol:
                        last_pol = pol
                    pod = clean_cell(r[1])
                    if not pol or not pod:
                        continue
                    ROTATION.append({"pol": pol, "pod": pod, "line": clean_cell(r[2]), "tt": clean_cell(r[3]), "country": clean_cell(r[4])})
                except:
                    continue
        if "SLMC" in wb.sheetnames:
            ws = wb["SLMC"]
            last_pol = ""
            for r in ws.iter_rows(values_only=True):
                try:
                    if not r:
                        continue
                    vals = []
                    for x in r:
                        vals.append(clean_cell(x))
                    if not any(vals):
                        continue
                    up = []
                    for v in vals:
                        up.append(v.upper())
                    if "POL" in up:
                        continue
                    pol = ""
                    pod = ""
                    line = ""
                    tt = ""
                    country = ""
                    if len(vals) >= 4 and vals[1] and vals[2]:
                        pol = vals[1]
                        pod = vals[2]
                        line = vals[3]
                        tt = vals[4] if len(vals) > 4 else ""
                        country = vals[5] if len(vals) > 5 else ""
                    elif len(vals) >= 3 and vals[0] and vals[1]:
                        pol = vals[0]
                        pod = vals[1]
                        line = vals[2]
                        tt = vals[3] if len(vals) > 3 else ""
                        country = vals[4] if len(vals) > 4 else ""
                    elif len(vals) >= 3 and vals[2]:
                        pol = last_pol
                        pod = vals[2]
                        line = vals[3] if len(vals) > 3 else ""
                        tt = vals[4] if len(vals) > 4 else ""
                        country = vals[5] if len(vals) > 5 else ""
                    elif len(vals) >= 2 and vals[1]:
                        pol = last_pol
                        pod = vals[1]
                        line = vals[2] if len(vals) > 2 else ""
                    if pol:
                        last_pol = pol
                    if not pol or not pod:
                        continue
                    SLMC_LIST.append({"pol": pol, "pod": pod, "line": line, "tt": tt, "country": country})
                except:
                    continue
        wb.close()
        if os.path.exists(OVERRIDE_FILE):
            try:
                with open(OVERRIDE_FILE, "r", encoding="utf-8") as f:
                    ov = json.load(f)
                for k, lst in ov.items():
                    if k not in LINE_NAMES:
                        LINE_NAMES[k] = k
                    if k not in LINE_CONTACTS:
                        LINE_CONTACTS[k] = []
                    LINE_CONTACTS[k].extend(lst)
            except:
                pass
        print("Shipping indexed")
    except Exception as e:
        print("Shipping index error: " + str(e))
build_shipping_index()
ARABIC_LINE_MAP = {}
ARABIC_LINE_MAP["ميرسك"] = "MAERSK"
ARABIC_LINE_MAP["هاباج"] = "HAPAG"
ARABIC_LINE_MAP["هاباق"] = "HAPAG"
ARABIC_LINE_MAP["كوسكو"] = "COSCO"
ARABIC_LINE_MAP["سي ام ايه"] = "CMA"
ARABIC_LINE_MAP["سي ام اي"] = "CMA"
ARABIC_LINE_MAP["ايفرجرين"] = "EVERGREEN"
ARABIC_LINE_MAP["ايفر جرين"] = "EVERGREEN"
ARABIC_LINE_MAP["ام اس سي"] = "MSC"
ARABIC_LINE_MAP["سيجو"] = "SEAGO"
ARABIC_LINE_MAP["سيغو"] = "SEAGO"
ARABIC_LINE_MAP["زيم"] = "ZIM"
ARABIC_LINE_MAP["وان هاي"] = "WAN HAI"
ARABIC_LINE_MAP["وانهاي"] = "WAN HAI"
ARABIC_LINE_MAP["اركاس"] = "ARKAS"
ARABIC_LINE_MAP["اركس"] = "ARKAS"
ARABIC_LINE_MAP["وان"] = "ONE"
ARABIC_LINE_MAP["او او سي ال"] = "OOCL"
ARABIC_LINE_MAP["واي ام ال"] = "YML"
ARABIC_LINE_MAP["اي بي ال"] = "APL"
ARABIC_LINE_MAP["بي اي ال"] = "PIL"
ARABIC_LINE_MAP["بيل"] = "PIL"
ARABIC_LINE_MAP["تاروس"] = "TARROS"
ARABIC_LINE_MAP["تاروز"] = "TARROS"
ARABIC_LINE_MAP["توركون"] = "TURKON"
ARABIC_LINE_MAP["ميدكون"] = "MEDKON"
ARABIC_LINE_MAP["ميسينا"] = "MESSINA"
ARABIC_LINE_MAP["ساف مارين"] = "SAF MARINE"
ARABIC_LINE_MAP["هامبورج سود"] = "HAMBURG SUD"
ARABIC_LINE_MAP["هامبورغ سود"] = "HAMBURG SUD"
ARABIC_LINE_MAP["بي ام سي"] = "BMC"
ARABIC_LINE_MAP["ترانسمار"] = "TRANSMAR"
ARABIC_LINE_MAP["ترانس مار"] = "TRANSMAR"
ARABIC_LINE_MAP["ويك لاين"] = "WEC LINE"
ARABIC_LINE_MAP["جريمالدي"] = "GRIMALDI"
ARABIC_LINE_MAP["هيونداي"] = "HYUNDAI"
ARABIC_LINE_MAP["هيونداى"] = "HYUNDAI"
ARABIC_LINE_MAP["دلتا"] = "DELTA"
ARABIC_LINE_MAP["ام جي ال"] = "MGL"
ARABIC_LINE_MAP["روديمار"] = "RODYMAR"
ARABIC_LINE_MAP["نايل داتش"] = "NILE DUTCH"
ARABIC_LINE_MAP["ادميرال"] = "ADMIRAL"
ARABIC_LINE_MAP["بورشارد"] = "BOARSHARD"
ARABIC_LINE_MAP["مارين"] = "MARINE"
SHORTCUT_MAP = {}
SHORTCUT_MAP["م"] = "MAERSK"
SHORTCUT_MAP["ه"] = "HAPAG"
SHORTCUT_MAP["ك"] = "COSCO"
SHORTCUT_MAP["س"] = "CMA"
SHORTCUT_MAP["ا"] = "EVERGREEN"
def load_hits():
    if not os.path.exists(STATS_FILE):
        return {"lines": {}, "ports": {}}
    try:
        with open(STATS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return {"lines": {}, "ports": {}}
def save_hits(h):
    try:
        with open(STATS_FILE, "w", encoding="utf-8") as f:
            json.dump(h, f, ensure_ascii=False, indent=2)
    except:
        pass
def hit_line(line):
    h = load_hits()
    h["lines"][line] = h["lines"].get(line, 0) + 1
    save_hits(h)
def hit_port(port):
    h = load_hits()
    h["ports"][port] = h["ports"].get(port, 0) + 1
    save_hits(h)
def resolve_line(q):
    qq = (q or "").strip()
    if not qq:
        return None
    if qq.upper() in LINE_CONTACTS:
        return qq.upper()
    if qq.lower() in ARABIC_LINE_MAP:
        return ARABIC_LINE_MAP[qq.lower()]
    if qq.strip() in SHORTCUT_MAP:
        return SHORTCUT_MAP[qq.strip()]
    ql = qq.lower()
    for k in LINE_NAMES:
        if ql in k.lower() or ql in LINE_NAMES[k].lower():
            return k
    return None
def suggest_line(q):
    m = difflib.get_close_matches((q or "").strip().upper(), list(LINE_NAMES.keys()), n=1, cutoff=0.55)
    if m:
        return m[0]
    m2 = difflib.get_close_matches((q or "").strip().lower(), list(ARABIC_LINE_MAP.keys()), n=1, cutoff=0.55)
    if m2:
        return ARABIC_LINE_MAP[m2[0]]
    return None
def add_recent(context, txt):
    lst = context.user_data.get("recent", [])
    txt = (txt or "").strip()
    if txt and txt not in lst:
        lst.insert(0, txt)
    context.user_data["recent"] = lst[:5]
def main_menu_keyboard():
    kb = []
    kb.append([InlineKeyboardButton("🚢 توكيلات", callback_data="menu_lines"), InlineKeyboardButton("📦 GPC", callback_data="menu_gpc")])
    kb.append([InlineKeyboardButton("🧾 التعريفة الجمركية", callback_data="menu_tariff"), InlineKeyboardButton("📁 المستندات / تزيل", callback_data="menu_docs")])
    kb.append([InlineKeyboardButton("🚢 Rotation", callback_data="menu_rotation"), InlineKeyboardButton("🏢 ارقام تسجيل الشركات", callback_data="menu_companies")])
    kb.append([InlineKeyboardButton("🌍 Country", callback_data="menu_country"), InlineKeyboardButton("🏙️ Cities", callback_data="menu_cities")])
    kb.append([InlineKeyboardButton("💵 Currency", callback_data="menu_currency"), InlineKeyboardButton("📦 Package", callback_data="menu_package")])
    kb.append([InlineKeyboardButton("📏 UOM", callback_data="menu_uom"), InlineKeyboardButton("🕘 اخر بحث", callback_data="menu_recent")])
    kb.append([InlineKeyboardButton("❓ مساعدة / Help", callback_data="menu_help")])
    return InlineKeyboardMarkup(kb)
def line_category_keyboard(line_key):
    kb = []
    kb.append([InlineKeyboardButton("📘 Booking", callback_data="linecat_" + line_key + "_BOOKING"), InlineKeyboardButton("🧾 Invoice", callback_data="linecat_" + line_key + "_INVOICE")])
    kb.append([InlineKeyboardButton("📧 كل الايميلات", callback_data="linecat_" + line_key + "_ALL"), InlineKeyboardButton("📞 كل الارقام", callback_data="linecat_" + line_key + "_PHONE")])
    kb.append([InlineKeyboardButton("🏠 القايمة الرئيسية", callback_data="menu_main")])
    return InlineKeyboardMarkup(kb)
def format_line_contacts(line_key, cat="ALL"):
    contacts = LINE_CONTACTS.get(line_key, [])
    display = LINE_NAMES.get(line_key, line_key)
    if not contacts:
        return "❌ ملقيتش توكيل اسمه: " + display
    txt = "🚢 " + display + "\n"
    if cat == "BOOKING":
        txt += "📘 Booking:\n\n"
        c = 0
        for x in contacts:
            if x["booking"] or x["email"]:
                txt += "• " + (x["name"] or x["position"] or "") + " - " + (x["email"] or "") + " " + (x["number"] or "") + "\nBooking: " + x["booking"] + "\n\n"
                c += 1
        if c == 0:
            txt += "مفيش Booking مسجل\n"
        return txt.strip()
    if cat == "INVOICE":
        txt += "🧾 Invoice:\n\n"
        for x in contacts:
            if x["invoice"]:
                txt += "• " + (x["name"] or x["position"] or "") + " - " + x["email"] + "\nInvoice: " + x["invoice"] + "\n\n"
        out = txt.strip()
        if out:
            return out
        return "مفيش Invoice لـ " + display
    if cat == "PHONE":
        txt += "📞 الارقام:\n\n"
        for x in contacts:
            if x["number"]:
                txt += "• " + (x["name"] or x["position"] or "") + ": " + x["number"] + "\n"
        return txt.strip()
    txt += "📞 عدد جهات الاتصال: " + str(len(contacts)) + "\n\n"
    for c in contacts[:12]:
        if c["name"]:
            txt += "👤 " + c["name"] + "\n"
        if c["position"]:
            txt += "💼 " + c["position"] + "\n"
        if c["email"]:
            txt += "📧 " + c["email"] + "\n"
        if c["number"]:
            txt += "📞 " + c["number"] + "\n"
        if c["booking"]:
            txt += "📘 Booking: " + c["booking"] + "\n"
        if c["invoice"]:
            txt += "🧾 Invoice: " + c["invoice"] + "\n"
        txt += "\n"
    if len(contacts) > 12:
        txt += "... و " + str(len(contacts) - 12) + " تانيين\n"
    return txt.strip()
def copy_keyboard(line_key):
    contacts = LINE_CONTACTS.get(line_key, [])
    kb = []
    for i, c in enumerate(contacts[:5]):
        if c["email"]:
            kb.append([InlineKeyboardButton("📋 نسخ " + c["email"][:30], callback_data="copy_" + line_key + "_" + str(i))])
    kb.append([InlineKeyboardButton("🏠 القايمة", callback_data="menu_main")])
    return InlineKeyboardMarkup(kb)
def search_rotation(pol, pod):
    pol_l = pol.lower().strip()
    pod_l = pod.lower().strip()
    out = []
    for r in ROTATION:
        if pol_l in r["pol"].lower() and pod_l in r["pod"].lower():
            out.append(r)
    for r in SLMC_LIST:
        if pol_l in r["pol"].lower() and pod_l in r["pod"].lower():
            if r not in out:
                out.append(r)
    return out[:10]
def format_rotation(results, pol, pod):
    if not results:
        return "❌ ملقيتش خط من " + pol + " لـ " + pod
    txt = "🚢 من " + pol + " لـ " + pod + " - لقيت " + str(len(results)) + ":\n\n"
    for r in results:
        tt = " - T.T: " + r["tt"] if r["tt"] else ""
        txt += "• " + r["line"] + " | " + r["pol"] + " -> " + r["pod"] + tt + " | " + r["country"] + "\n"
    return txt
def get_filtered_lines(query=""):
    lst = sorted(LINE_NAMES.keys(), key=lambda x: LINE_NAMES[x])
    q = (query or "").strip().lower()
    if not q:
        return lst
    out = []
    for k in lst:
        if q in LINE_NAMES[k].lower() or q in k.lower():
            out.append(k)
    return out
def line_list_keyboard(page=0, query=""):
    lst = get_filtered_lines(query)
    per = 10
    start = page * per
    end = start + per
    total_pages = max(1, (len(lst) + per - 1) // per)
    kb = []
    for i in range(start, min(end, len(lst))):
        k = lst[i]
        kb.append([InlineKeyboardButton(LINE_NAMES[k][:55], callback_data="linesel_" + str(i))])
    nav = []
    if page > 0:
        nav.append(InlineKeyboardButton("⬅️ السابق", callback_data="linepage_" + str(page - 1)))
    if page < total_pages - 1:
        nav.append(InlineKeyboardButton("التالي ➡️", callback_data="linepage_" + str(page + 1)))
    if nav:
        kb.append(nav)
    kb.append([InlineKeyboardButton("🏠 القايمة الرئيسية", callback_data="menu_main")])
    title = "بحث: " + query + " - " if query else ""
    txt = title + "اختار التوكيل - صفحة " + str(page + 1) + " من " + str(total_pages) + " - العدد " + str(len(lst))
    return InlineKeyboardMarkup(kb), txt
async def send_line_list(update, context, page=0, query=""):
    uid = None
    if update.message:
        uid = update.message.from_user.id
    elif update.callback_query:
        uid = update.callback_query.from_user.id
    if uid is not None:
        context.bot_data["line_filter_" + str(uid)] = {"query": query}
    kb, txt = line_list_keyboard(page, query)
    if update.message:
        await update.message.reply_text(txt, reply_markup=kb)
    elif update.callback_query:
        try:
            await update.callback_query.message.edit_text(txt, reply_markup=kb)
        except:
            await update.callback_query.message.reply_text(txt, reply_markup=kb)
def load_docs():
    if not os.path.exists(DOCS_FILE):
        return []
    try:
        with open(DOCS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return []
def save_docs(docs):
    with open(DOCS_FILE, "w", encoding="utf-8") as f:
        json.dump(docs, f, ensure_ascii=False, indent=2)
WELCOME_MESSAGE = "Welcome to Captain Emad Customs Bot. After the trial and subscription for 150 EGP, you can inquire about the customs item by code or by item name in Arabic or English. You can also search for GS1, and you can calculate the approximate customs."
HELP_TEXT = "مساعدة - تعريفة اكتب 0302.44 او خيول - توكيلات اكتب توكيلات او MAERSK او ميرسك - خط اكتب Alex Izmir - شركات اكتب شركات او رقم التسجيل 202444732 - GPC اكتب gpc Airbrushes - Country اكتب Country Eg - Cities اكتب Cities Cai - Currency اكتب Currency US - Package اكتب Package B - UOM اكتب UOM K - مستندات اكتب docs او تزيل 1 - اخر بحث اكتب اخر - القايمة اكتب menu"
def load_users():
    if not os.path.exists(USERS_FILE):
        return {}
    try:
        with open(USERS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return {}
def save_users(users):
    with open(USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(users, f, ensure_ascii=False, indent=2)
def generate_code(users):
    while True:
        code = "USR-" + "".join(random.choices(string.digits, k=4))
        found = False
        for u in users.values():
            if u.get("code") == code:
                found = True
        if not found:
            return code
def clean_text(text):
    if not text:
        return ""
    text = clean_invisible(text)
    arabic_digits = "٠١٢٣٤٥٦٧٨٩"
    for i, d in enumerate(arabic_digits):
        text = text.replace(d, str(i))
    text = text.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا").replace("ة", "ه")
    return text.strip()
def is_download_request(text):
    t = clean_text(text)
    return bool(re.search(r"(تزيل|تنزيل|تحميل|download|نزل)", t))
def extract_number(text):
    t = clean_text(text)
    m = re.search(r"(\d+)", t)
    if m:
        return int(m.group(1))
    return None
def normalize(text):
    return re.sub(r"\s+", " ", (text or "").strip().lower())
def search_gpc_fast(query):
    q = normalize(query)
    if not q:
        return None, None
    q_digits = re.sub(r"\D", "", q)
    if q_digits and q_digits in BRICK_BY_CODE:
        return BRICK_INDEX[BRICK_BY_CODE[q_digits]], None
    if q in BRICK_INDEX:
        return BRICK_INDEX[q], None
    matches = []
    for k in BRICK_INDEX.keys():
        if q in k:
            matches.append(BRICK_INDEX[k])
    if len(matches) == 1:
        return matches[0], None
    if len(matches) > 1:
        return ("multiple", matches[:10]), None
    return None, None
def get_attrs_for_brick(brick_title):
    return BRICK_ATTRS.get(brick_title.strip().lower(), [])
def format_brick(brick, attrs):
    txt = "📦 " + brick["brick_title"] + "\n"
    txt += "🔢 Brick Code: " + brick["brick_code"] + "\n\n"
    txt += "📂 Segment: " + str(brick["segment"]) + " (" + str(brick["segment_code"]) + ")\n"
    txt += "📁 Family: " + str(brick["family"]) + " (" + str(brick["family_code"]) + ")\n"
    txt += "📄 Class: " + str(brick["class"]) + " (" + str(brick["class_code"]) + ")\n\n"
    if brick.get("includes"):
        txt += "✅ Includes:\n" + str(brick["includes"]) + "\n\n"
    if brick.get("excludes"):
        txt += "⛔ Excludes:\n" + str(brick["excludes"]) + "\n\n"
    if attrs:
        txt += "⚙️ Attributes:\n"
        for a in attrs:
            txt += "\n• " + str(a["attribute_title"]) + " (" + str(a["attribute_code"]) + ")\n"
            if a.get("values"):
                vals = a["values"][:15]
                txt += "  " + ", ".join(vals)
                if len(a["values"]) > 15:
                    txt += " ... (+" + str(len(a["values"]) - 15) + " more)"
                txt += "\n"
    return txt
def search_ref(kind, query):
    q = query.strip()
    qu = q.upper().strip()
    if kind == "cities":
        by_code = CITY_BY_CODE
        index = CITY_INDEX
    elif kind == "country":
        by_code = COUNTRY_BY_CODE
        index = COUNTRY_INDEX
    elif kind == "currency":
        by_code = CURRENCY_BY_CODE
        index = CURRENCY_INDEX
    elif kind == "package":
        by_code = PACKAGE_BY_CODE
        index = PACKAGE_INDEX
    elif kind == "uom":
        by_code = UOM_BY_CODE
        index = UOM_INDEX
    else:
        return None, None
    if qu in by_code:
        return by_code[qu], None
    ql = q.lower()
    matches = []
    for key, lst in index.items():
        if ql in key:
            matches.extend(lst)
        if len(matches) >= 10:
            break
    if not matches:
        for code, item in by_code.items():
            if qu in code:
                matches.append(item)
            if len(matches) >= 10:
                break
    if len(matches) == 1:
        return matches[0], None
    if len(matches) > 1:
        return ("multiple", matches[:10]), None
    return None, None
def format_ref(kind, item):
    if kind == "cities":
        extra = ""
        if item.get("country"):
            extra = "\n🌍 Country: " + str(item.get("country", ""))
        return "🏙️ " + item["name"] + "\n🔢 City Code: " + item["code"] + extra
    if kind == "country":
        return "🌍 " + item["name"] + "\n🔢 Country Code: " + item["code"]
    if kind == "currency":
        return "💵 " + item["name"] + "\n🔢 Currency Code: " + item["code"]
    if kind == "package":
        return "📦 " + item["name"] + "\n🔢 Package Code: " + item["code"]
    if kind == "uom":
        return "📏 " + item["name"] + "\n🔢 UOM Code: " + item["code"]
    return item["name"] + " - " + item["code"]
def search_all_refs(query):
    results = []
    q = query.strip()
    qu = q.upper().strip()
    ql = q.lower()
    if qu in CITY_BY_CODE:
        results.append(("cities", CITY_BY_CODE[qu]))
    if qu in COUNTRY_BY_CODE:
        results.append(("country", COUNTRY_BY_CODE[qu]))
    if qu in CURRENCY_BY_CODE:
        results.append(("currency", CURRENCY_BY_CODE[qu]))
    if qu in PACKAGE_BY_CODE:
        results.append(("package", PACKAGE_BY_CODE[qu]))
    if qu in UOM_BY_CODE:
        results.append(("uom", UOM_BY_CODE[qu]))
    if results:
        return results
    pairs = []
    pairs.append(("cities", CITY_INDEX))
    pairs.append(("country", COUNTRY_INDEX))
    pairs.append(("currency", CURRENCY_INDEX))
    pairs.append(("package", PACKAGE_INDEX))
    pairs.append(("uom", UOM_INDEX))
    for kind, index in pairs:
        for key, lst in index.items():
            if ql in key:
                for item in lst:
                    results.append((kind, item))
                if len(results) >= 10:
                    return results
    return results
def get_sorted_ref_list(kind):
    if kind == "cities":
        return sorted(CITY_BY_CODE.values(), key=lambda x: x["code"])
    if kind == "country":
        return sorted(COUNTRY_BY_CODE.values(), key=lambda x: x["code"])
    if kind == "currency":
        return sorted(CURRENCY_BY_CODE.values(), key=lambda x: x["code"])
    if kind == "package":
        return sorted(PACKAGE_BY_CODE.values(), key=lambda x: x["code"])
    if kind == "uom":
        return sorted(UOM_BY_CODE.values(), key=lambda x: x["code"])
    return []
def get_filtered_ref_list(kind, query=""):
    lst = get_sorted_ref_list(kind)
    q = (query or "").strip().lower()
    if not q:
        return lst
    qu = q.upper()
    out = []
    for item in lst:
        if q in item["name"].lower() or qu in item["code"]:
            out.append(item)
    return out
def ref_list_keyboard(kind, page=0, query=""):
    lst = get_filtered_ref_list(kind, query)
    per = 10
    start = page * per
    end = start + per
    total_pages = max(1, (len(lst) + per - 1) // per)
    kb = []
    for i in range(start, min(end, len(lst))):
        item = lst[i]
        btn = item["name"][:40] + " - " + item["code"]
        kb.append([InlineKeyboardButton(btn, callback_data="refsel_" + kind + "_" + str(i))])
    nav = []
    if page > 0:
        nav.append(InlineKeyboardButton("⬅️ السابق", callback_data="refpage_" + kind + "_" + str(page - 1)))
    if page < total_pages - 1:
        nav.append(InlineKeyboardButton("التالي ➡️", callback_data="refpage_" + kind + "_" + str(page + 1)))
    if nav:
        kb.append(nav)
    kb.append([InlineKeyboardButton("🏠 القايمة الرئيسية", callback_data="menu_main")])
    title = "بحث: " + query + " - " if query else ""
    txt = title + "اختار من " + kind + " - صفحة " + str(page + 1) + " من " + str(total_pages) + " - العدد " + str(len(lst))
    return InlineKeyboardMarkup(kb), txt
async def send_ref_list(update, context, kind, page=0, query=""):
    uid = None
    if update.message:
        uid = update.message.from_user.id
    elif update.callback_query:
        uid = update.callback_query.from_user.id
    if uid is not None:
        context.bot_data["ref_filter_" + str(uid)] = {"kind": kind, "query": query}
    kb, txt = ref_list_keyboard(kind, page, query)
    if update.message:
        await update.message.reply_text(txt, reply_markup=kb)
    elif update.callback_query:
        try:
            await update.callback_query.message.edit_text(txt, reply_markup=kb)
        except:
            await update.callback_query.message.reply_text(txt, reply_markup=kb)
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.message.from_user
    users = load_users()
    uid = str(user.id)
    now = time.time()
    if uid in users and users[uid].get("status") == "rejected":
        last = users[uid].get("rejected_at", 0)
        remaining = int((REJECT_WAIT - (now - last)) / 60) + 1
        if now - last < REJECT_WAIT:
            kb = InlineKeyboardMarkup([[InlineKeyboardButton("🔁 اطلب قبول تاني", callback_data="request_again")]])
            await update.message.reply_text("Your request was rejected. You can send again after " + str(remaining) + " minutes.", reply_markup=kb)
            return
        users[uid]["status"] = "pending"
        users[uid]["requested_at"] = now
        save_users(users)
    if uid in users and users[uid].get("status") == "accepted":
        await update.message.reply_text(WELCOME_MESSAGE + "\n\n✅ Your account is accepted. Send the item code or item name.", reply_markup=main_menu_keyboard())
        return
    if uid not in users:
        code = generate_code(users)
        users[uid] = {"code": code, "name": user.first_name or "", "username": user.username or "", "id": user.id, "status": "pending", "requested_at": now}
        save_users(users)
    else:
        code = users[uid]["code"]
        users[uid]["status"] = "pending"
        users[uid]["requested_at"] = now
        save_users(users)
    await update.message.reply_text(WELCOME_MESSAGE + "\n\nYour Code: " + code + "\nWaiting for acceptance from Captain Emad.", reply_markup=main_menu_keyboard())
    if ADMIN_ID:
        kb2 = []
        kb2.append([InlineKeyboardButton("✅ قبول / Accept", callback_data="accept_" + uid), InlineKeyboardButton("❌ رفض / Reject", callback_data="reject_" + uid)])
        try:
            await context.bot.send_message(chat_id=int(ADMIN_ID), text="📦 New User Request\n\nCode: " + code + "\nName: " + str(user.first_name) + "\nUsername: @" + str(user.username) + "\nID: " + str(user.id), reply_markup=InlineKeyboardMarkup(kb2))
        except Exception as e:
            print("Admin send error: " + str(e))
async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(HELP_TEXT, reply_markup=main_menu_keyboard())
async def menu_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("🏠 القايمة الرئيسية - اختار:", reply_markup=main_menu_keyboard())
async def top_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if str(update.message.from_user.id) != str(ADMIN_ID):
        return
    h = load_hits()
    lines = sorted(h["lines"].items(), key=lambda x: x[1], reverse=True)[:10]
    ports = sorted(h["ports"].items(), key=lambda x: x[1], reverse=True)[:10]
    txt = "📊 اكتر حاجة اتسأل عليها:\n\n🚢 خطوط:\n"
    for k, v in lines:
        txt += "• " + k + ": " + str(v) + "\n"
    txt += "\n⚓ موانئ:\n"
    for k, v in ports:
        txt += "• " + k + ": " + str(v) + "\n"
    await update.message.reply_text(txt or "لسه مفيش بحث")
async def button_click(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    try:
        await query.answer()
    except:
        pass
    users = load_users()
    data = query.data
    if data == "menu_main":
        await query.message.reply_text("🏠 القايمة الرئيسية - اختار:", reply_markup=main_menu_keyboard())
        return
    if data == "menu_lines":
        await send_line_list(query, context, 0, "")
        return
    if data == "menu_companies":
        await send_companies_submenu(query, context)
        return
    if data == "comp_type_private":
        await send_company_list(query, context, 0, "", "private")
        return
    if data == "comp_type_gov":
        await send_company_list(query, context, 0, "", "gov")
        return
    if data == "comp_type_foreign":
        await send_company_list(query, context, 0, "", "foreign")
        return
    if data == "menu_tariff":
        await send_tariff_list(query, context, 0, "")
        return
    if data == "menu_gpc":
        await query.message.reply_text("اكتب: gpc Airbrushes او اسم الصنف", reply_markup=main_menu_keyboard())
        return
    if data == "menu_country":
        await send_ref_list(query, context, "country", 0, "")
        return
    if data == "menu_cities":
        await send_ref_list(query, context, "cities", 0, "")
        return
    if data == "menu_currency":
        await send_ref_list(query, context, "currency", 0, "")
        return
    if data == "menu_package":
        await send_ref_list(query, context, "package", 0, "")
        return
    if data == "menu_uom":
        await send_ref_list(query, context, "uom", 0, "")
        return
    if data == "menu_rotation":
        await query.message.reply_text("اكتب: Alex Izmir او EGYPT LATTAKIA", reply_markup=main_menu_keyboard())
        return
    if data == "menu_docs":
        docs = load_docs()
        if not docs:
            await query.message.reply_text("مفيش مستندات لسه", reply_markup=main_menu_keyboard())
            return
        txt = "📁 المستندات:\n\n"
        for d in docs:
            txt += str(d["num"]) + " - " + d["name"] + "\n"
        txt += "\nاكتب: تزيل 1"
        await query.message.reply_text(txt, reply_markup=main_menu_keyboard())
        return
    if data == "menu_help":
        await query.message.reply_text(HELP_TEXT, reply_markup=main_menu_keyboard())
        return
    if data == "menu_recent":
        lst = context.user_data.get("recent", [])
        if not lst:
            await query.message.reply_text("لسه مفيش بحث")
            return
        out = []
        for x in lst:
            out.append("• " + x)
        await query.message.reply_text("🕘 اخر 5:\n" + "\n".join(out))
        return
    if data == "request_again":
        uid = str(query.from_user.id)
        users2 = load_users()
        if uid in users2:
            users2[uid]["status"] = "pending"
            users2[uid]["requested_at"] = time.time()
            save_users(users2)
            await query.message.reply_text("✅ طلبك اتبعت تاني لكابتن عماد، استنى القبول")
            if ADMIN_ID:
                await context.bot.send_message(chat_id=int(ADMIN_ID), text="🔁 طلب قبول تاني\nCode: " + users2[uid]["code"] + " - " + users2[uid]["name"] + " - ID: " + uid)
        return
    if data.startswith("tarifpage_"):
        page = int(data.split("_")[1])
        filt = context.bot_data.get("tariff_filter_" + str(query.from_user.id), {})
        q = filt.get("query", "")
        kb, txt = tariff_list_keyboard(page, q)
        try:
            await query.message.edit_text(txt, reply_markup=kb)
        except:
            await query.message.reply_text(txt, reply_markup=kb)
        return
    if data.startswith("tarifsel_"):
        idx = int(data.split("_")[1])
        filt = context.bot_data.get("tariff_filter_" + str(query.from_user.id), {})
        q = filt.get("query", "")
        lst = get_filtered_tariff(q)
        if 0 <= idx < len(lst):
            await query.message.reply_text(format_tariff(lst[idx]))
        return
    if data.startswith("comppage_"):
        page = int(data.split("_")[1])
        filt = context.bot_data.get("comp_filter_" + str(query.from_user.id), {})
        q = filt.get("query", "")
        ct = filt.get("ctype", "")
        kb, txt = company_list_keyboard(page, q, ct)
        try:
            await query.message.edit_text(txt, reply_markup=kb)
        except:
            await query.message.reply_text(txt, reply_markup=kb)
        return
    if data.startswith("compsel_"):
        idx = int(data.split("_")[1])
        filt = context.bot_data.get("comp_filter_" + str(query.from_user.id), {})
        q = filt.get("query", "")
        ct = filt.get("ctype", "")
        lst = get_filtered_companies(q, ct)
        if 0 <= idx < len(lst):
            await query.message.reply_text(format_company(lst[idx]))
        return
    if data.startswith("refpage_"):
        parts = data.split("_")
        kind = parts[1]
        page = int(parts[2])
        filt = context.bot_data.get("ref_filter_" + str(query.from_user.id), {})
        q = filt.get("query", "")
        kb, txt = ref_list_keyboard(kind, page, q)
        try:
            await query.message.edit_text(txt, reply_markup=kb)
        except:
            await query.message.reply_text(txt, reply_markup=kb)
        return
    if data.startswith("refsel_"):
        parts = data.split("_")
        kind = parts[1]
        idx = int(parts[2])
        filt = context.bot_data.get("ref_filter_" + str(query.from_user.id), {})
        q = filt.get("query", "")
        lst = get_filtered_ref_list(kind, q)
        if 0 <= idx < len(lst):
            await query.message.reply_text(format_ref(kind, lst[idx]))
        return
    if data.startswith("linepage_"):
        page = int(data.split("_")[1])
        filt = context.bot_data.get("line_filter_" + str(query.from_user.id), {})
        q = filt.get("query", "")
        kb, txt = line_list_keyboard(page, q)
        try:
            await query.message.edit_text(txt, reply_markup=kb)
        except:
            await query.message.reply_text(txt, reply_markup=kb)
        return
    if data.startswith("linesel_"):
        idx = int(data.split("_")[1])
        filt = context.bot_data.get("line_filter_" + str(query.from_user.id), {})
        q = filt.get("query", "")
        lst = get_filtered_lines(q)
        if 0 <= idx < len(lst):
            k = lst[idx]
            hit_line(k)
            add_recent(context, LINE_NAMES[k])
            await query.message.reply_text(format_line_contacts(k), reply_markup=line_category_keyboard(k))
            await query.message.reply_text("📋 دوس عشان تنسخ:", reply_markup=copy_keyboard(k))
        return
    if data.startswith("linecat_"):
        parts = data.split("_")
        line_key = parts[1]
        cat = parts[2]
        await query.message.reply_text(format_line_contacts(line_key, cat), reply_markup=line_category_keyboard(line_key))
        return
    if data.startswith("copy_"):
        parts = data.split("_")
        line_key = parts[1]
        idx = int(parts[2])
        contacts = LINE_CONTACTS.get(line_key, [])
        if 0 <= idx < len(contacts):
            c = contacts[idx]
            await query.message.reply_text(c["email"] + "\n" + c["number"])
        return
    if data.startswith("accept_"):
        uid = data.split("_")[1]
        if uid in users:
            users[uid]["status"] = "accepted"
            save_users(users)
            await context.bot.send_message(chat_id=int(uid), text=WELCOME_MESSAGE + "\n\n✅ Your request has been accepted by Captain Emad.", reply_markup=main_menu_keyboard())
            await query.message.reply_text("Accepted ✅\nCode: " + users[uid]["code"] + " - " + users[uid]["name"])
    elif data.startswith("reject_"):
        uid = data.split("_")[1]
        if uid in users:
            users[uid]["status"] = "rejected"
            users[uid]["rejected_at"] = time.time()
            save_users(users)
            kb = InlineKeyboardMarkup([[InlineKeyboardButton("🔁 اطلب قبول تاني", callback_data="request_again")]])
            await context.bot.send_message(chat_id=int(uid), text="❌ Your request was rejected by Captain Emad. You can send again after 5 minutes.", reply_markup=kb)
            await query.message.reply_text("Rejected ❌\nCode: " + users[uid]["code"] + " - " + users[uid]["name"])
    elif data.startswith("gpc_"):
        try:
            idx = int(data[4:])
            keys = context.user_data.get("gpc_matches", [])
            if not keys and context.bot_data.get("gpc_" + str(query.from_user.id)):
                keys = context.bot_data.get("gpc_" + str(query.from_user.id), [])
            if 0 <= idx < len(keys):
                b = keys[idx]
                attrs = get_attrs_for_brick(b["brick_title"])
                await query.message.reply_text(format_brick(b, attrs))
            else:
                await query.message.reply_text("القايمة دي انتهت، ابعت gpc تاني واختار من جديد")
        except Exception as e:
            print("gpc callback error: " + str(e))
async def save_doc(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if str(update.message.from_user.id) != str(ADMIN_ID):
        return
    caption = (update.message.caption or "").strip() or "بدون اسم"
    file = None
    ext = "jpg"
    if update.message.photo:
        file = await update.message.photo[-1].get_file()
    elif update.message.document:
        file = await update.message.document.get_file()
        if "." in update.message.document.file_name:
            ext = update.message.document.file_name.split(".")[-1]
        else:
            ext = "pdf"
    else:
        return
    docs = load_docs()
    num = max([d["num"] for d in docs], default=0) + 1
    filename = "docs/" + str(num) + "." + ext
    await file.download_to_drive(filename)
    docs.append({"num": num, "name": caption, "path": filename})
    save_docs(docs)
    await update.message.reply_text("✅ اتحفظ برقم " + str(num) + "\n📄 " + caption + "\n\nعشان تنزله اكتب: تزيل " + str(num))
async def list_docs(update: Update, context: ContextTypes.DEFAULT_TYPE):
    docs = load_docs()
    if not docs:
        await update.message.reply_text("مفيش مستندات لسه")
        return
    txt = "📁 المستندات:\n\n"
    for d in docs:
        txt += str(d["num"]) + " - " + d["name"] + "\n"
    txt += "\nاكتب: تزيل 1"
    await update.message.reply_text(txt)
async def download_doc(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.message.from_user
    uid = str(user.id)
    if str(user.id) != str(ADMIN_ID):
        users = load_users()
        if uid not in users or users[uid].get("status") != "accepted":
            await update.message.reply_text("Please send start first.")
            return
    num = extract_number(update.message.text)
    if num is None:
        await update.message.reply_text("اكتب: تزيل 1")
        return
    docs = load_docs()
    for d in docs:
        if d["num"] == num:
            try:
                await update.message.reply_document(document=open(d["path"], "rb"), caption="📄 رقم " + str(num) + " - " + d["name"])
            except FileNotFoundError:
                await update.message.reply_text("❌ الملف رقم " + str(num) + " مش موجود على السيرفر، ارفعه تاني")
            return
    await update.message.reply_text("❌ مفيش مستند رقم " + str(num) + "\nاكتب docs عشان تشوف الارقام")
async def gpc_search_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    users = load_users()
    uid = str(update.message.from_user.id)
    if str(update.message.from_user.id) != str(ADMIN_ID):
        if uid not in users or users[uid].get("status") != "accepted":
            await update.message.reply_text("Please send start first.")
            return
    if not context.args:
        await update.message.reply_text("اكتب: gpc اسم الصنف او الكود\nمثال: gpc Airbrushes\nاو: gpc 10001686")
        return
    query = " ".join(context.args)
    add_recent(context, query)
    brick, _ = search_gpc_fast(query)
    if brick is None:
        await update.message.reply_text("❌ ملقيتش صنف اسمه: " + query)
        return
    if isinstance(brick, tuple) and brick[0] == "multiple":
        matches = brick[1]
        context.user_data["gpc_matches"] = matches
        context.bot_data["gpc_" + str(update.message.from_user.id)] = matches
        keyboard = []
        for i, m in enumerate(matches):
            keyboard.append([InlineKeyboardButton(m["brick_title"][:60], callback_data="gpc_" + str(i))])
        await update.message.reply_text("لقيت " + str(len(matches)) + " صنف، اختار واحد:",