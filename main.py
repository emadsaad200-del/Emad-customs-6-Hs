import os, json, time, random, string, re, difflib
import openpyxl
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, CallbackQueryHandler, ContextTypes, filters
BOT_TOKEN=os.getenv("BOT_TOKEN"); ADMIN_ID=os.getenv("ADMIN_ID")
SLASH=chr(47)
USERS_FILE="users.json"; REJECT_WAIT=5*60; DOCS_FILE="docs_index.json"; STATS_FILE="hits.json"; OVERRIDE_FILE="lines_override.json"
os.makedirs("docs",exist_ok=True)
EXCEL_FILES=["gs1xlsx.xlsx","GPC as of May 2023 v20230605 EN.xlsx"]
BRICK_INDEX={}; BRICK_BY_CODE={}; BRICK_ATTRS={}
def build_index():
    print("Building light index...")
    for fname in EXCEL_FILES:
        if not os.path.exists(fname): continue
        try:
            wb=openpyxl.load_workbook(fname,read_only=True,data_only=True)
            ws=wb["Schema"] if "Schema" in wb.sheetnames else wb[wb.sheetnames[0]]
            for r in ws.iter_rows(values_only=True):
                try:
                    if not r or len(r)<11: continue
                    if str(r[0]).lower().startswith("segment"): continue
                    brick_code=str(r[9]) if r[9] is not None else ""; brick_title=str(r[10]) if r[10] is not None else ""
                    if not brick_title: continue
                    key=brick_title.strip().lower()
                    if key not in BRICK_INDEX:
                        BRICK_INDEX[key]={"brick_code":brick_code,"brick_title":brick_title,"segment":r[1],"segment_code":str(r[0]),"family":r[4],"family_code":str(r[3]),"class":r[7],"class_code":str(r[6]),"includes":r[11] or "","excludes":r[12] if len(r)>12 else "" or ""}
                        if brick_code: BRICK_BY_CODE[brick_code]=key
                    if len(r)>14 and r[13] is not None:
                        attr_code=str(r[13]); attr_title=r[14]; val_title=str(r[17]) if len(r)>17 and r[17] is not None else None
                        if key not in BRICK_ATTRS: BRICK_ATTRS[key]=[]
                        ex=None
                        for a in BRICK_ATTRS[key]:
                            if a["attribute_code"]==attr_code: ex=a; break
                        if not ex: ex={"attribute_code":attr_code,"attribute_title":attr_title,"values":[]}; BRICK_ATTRS[key].append(ex)
                        if val_title and val_title not in ex["values"]: ex["values"].append(val_title)
                except: continue
            wb.close(); print("Indexed "+fname)
        except Exception as e: print("Index error "+fname+": "+str(e))
    print("Index done")
build_index()
CITIES_FILE="cities.xlsx"
CITY_INDEX={}; CITY_BY_CODE={}; COUNTRY_INDEX={}; COUNTRY_BY_CODE={}; CURRENCY_INDEX={}; CURRENCY_BY_CODE={}; PACKAGE_INDEX={}; PACKAGE_BY_CODE={}; UOM_INDEX={}; UOM_BY_CODE={}
def build_reference_index():
    if not os.path.exists(CITIES_FILE): print("cities.xlsx not found"); return
    try:
        wb=openpyxl.load_workbook(CITIES_FILE,read_only=True,data_only=True)
        if "City List" in wb.sheetnames:
            ws=wb["City List"]
            for r in ws.iter_rows(values_only=True):
                try:
                    if not r or len(r)<2 or not r[0] or not r[1]: continue
                    if str(r[0]).lower()=="citycode": continue
                    code=str(r[0]).strip().upper(); name=str(r[1]).strip(); country=str(r[2]).strip() if len(r)>2 and r[2] else ""
                    CITY_BY_CODE[code]={"code":code,"name":name,"country":country}; k=name.lower()
                    if k not in CITY_INDEX: CITY_INDEX[k]=[]
                    CITY_INDEX[k].append({"code":code,"name":name,"country":country})
                except: continue
        if "Cities" in wb.sheetnames:
            ws=wb["Cities"]
            for r in ws.iter_rows(values_only=True):
                try:
                    if not r or len(r)<2 or not r[0] or not r[1]: continue
                    if str(r[0]).lower()=="citycode": continue
                    code=str(r[0]).strip().upper(); name=str(r[1]).strip()
                    if code in CITY_BY_CODE: continue
                    CITY_BY_CODE[code]={"code":code,"name":name,"country":""}; k=name.lower()
                    if k not in CITY_INDEX: CITY_INDEX[k]=[]
                    CITY_INDEX[k].append({"code":code,"name":name,"country":""})
                except: continue
        if "Country List" in wb.sheetnames:
            ws=wb["Country List"]
            for r in ws.iter_rows(values_only=True):
                try:
                    if not r or len(r)<2 or not r[0] or not r[1]: continue
                    if str(r[0]).lower()=="code": continue
                    code=str(r[0]).strip().upper(); name=str(r[1]).strip(); COUNTRY_BY_CODE[code]={"code":code,"name":name}; k=name.lower()
                    if k not in COUNTRY_INDEX: COUNTRY_INDEX[k]=[]
                    COUNTRY_INDEX[k].append({"code":code,"name":name})
                except: continue
        if "Currency List" in wb.sheetnames:
            ws=wb["Currency List"]
            for r in ws.iter_rows(values_only=True):
                try:
                    if not r or len(r)<2 or not r[0] or not r[1]: continue
                    if "currency" in str(r[0]).lower(): continue
                    name=str(r[0]).strip(); code=str(r[1]).strip().upper(); CURRENCY_BY_CODE[code]={"code":code,"name":name}; k=name.lower()
                    if k not in CURRENCY_INDEX: CURRENCY_INDEX[k]=[]
                    CURRENCY_INDEX[k].append({"code":code,"name":name})
                except: continue
        if "Package Type list" in wb.sheetnames:
            ws=wb["Package Type list"]
            for r in ws.iter_rows(values_only=True):
                try:
                    if not r or len(r)<2 or not r[0] or not r[1]: continue
                    if str(r[0]).lower()=="code": continue
                    code=str(r[0]).strip().upper(); name=str(r[1]).strip(); PACKAGE_BY_CODE[code]={"code":code,"name":name}; k=name.lower()
                    if k not in PACKAGE_INDEX: PACKAGE_INDEX[k]=[]
                    PACKAGE_INDEX[k].append({"code":code,"name":name})
                except: continue
        if "UOM List" in wb.sheetnames:
            ws=wb["UOM List"]
            for r in ws.iter_rows(values_only=True):
                try:
                    if not r or len(r)<2 or not r[0] or not r[1]: continue
                    if "uomcode" in str(r[0]).lower(): continue
                    code=str(r[0]).strip().upper(); name=str(r[1]).strip(); UOM_BY_CODE[code]={"code":code,"name":name}; k=name.lower()
                    if k not in UOM_INDEX: UOM_INDEX[k]=[]
                    UOM_INDEX[k].append({"code":code,"name":name})
                except: continue
        wb.close(); print("Ref indexed")
    except Exception as e: print("Ref index error: "+str(e))
build_reference_index()
COMPANIES_FILE="ارقام_تسجيل_الشركات_111111.xlsx"
COMPANIES_FILE_2026="ارقام_تسجيل_الشركات_2026.xlsx"
COMPANY_LIST=[]; COMPANY_BY_REG={}; COMPANY_BY_FILE={}
COMPANY_NORM_CACHE={}
def clean_cell(v):
    if v is None: return ""
    return str(v).strip()
def norm_ar(s):
    if not s: return ""
    s=str(s).strip().lower()
    for ch in ["\u200e","\u200f","\u202a","\u202b","\u202c","\u202d","\u202e","\u200c","\u200d","\u0640"]:
        s=s.replace(ch,"")
    s=s.replace("أ","ا").replace("إ","ا").replace("آ","ا").replace("ٱ","ا")
    s=s.replace("ة","ه").replace("ھ","ه").replace("ۀ","ه")
    s=s.replace("ى","ي").replace("ی","ي").replace("ئ","ي").replace("ې","ي")
    s=s.replace("ؤ","و").replace("ڪ","ك").replace("ک","ك")
    s=re.sub(r'[\u064B-\u0652\u0670]', '', s)
    s=s.translate(str.maketrans("٠١٢٣٤٥٦٧٨٩","0123456789"))
    s=re.sub(r'\s+',' ',s)
    return s.strip()
def clean_search_words(text):
    stop=["ال","شركه","شركة","لل","ل","شرك","وشركاه","وشريكه","وشريك","وشركاؤه"]
    words=norm_ar(text).split()
    out=[w for w in words if w not in stop and len(w)>=2]
    return out if out else words
def _add_company(name, file_no, reg_no, entity, activity, status, address):
    name=clean_cell(name); reg_no=clean_cell(reg_no).replace(" ","").replace("،","").replace(",","")
    if not name or not reg_no: return
    if reg_no in COMPANY_BY_REG: return
    file_no=clean_cell(file_no).replace(" ","").replace("،","").replace(",","")
    item={"name":name,"file_no":file_no,"reg_no":reg_no,"entity":clean_cell(entity),"activity":clean_cell(activity),"status":clean_cell(status),"address":clean_cell(address)}
    COMPANY_LIST.append(item)
    COMPANY_BY_REG[reg_no]=item
    if file_no: COMPANY_BY_FILE[file_no]=item
    COMPANY_NORM_CACHE[len(COMPANY_LIST)-1]={"name":norm_ar(name),"activity":norm_ar(activity),"address":norm_ar(address),"entity":norm_ar(entity)}
def build_companies_index():
    if os.path.exists(COMPANIES_FILE):
        try:
            wb=openpyxl.load_workbook(COMPANIES_FILE,read_only=True,data_only=True); ws=wb[wb.sheetnames[0]]
            for r in ws.iter_rows(values_only=True):
                try:
                    if not r: continue
                    name=clean_cell(r[0]) if len(r)>0 else ""
                    if not name or "الاسم" in name: continue
                    _add_company(name, clean_cell(r[1]) if len(r)>1 else "", clean_cell(r[2]) if len(r)>2 else "", clean_cell(r[3]) if len(r)>3 else "", clean_cell(r[4]) if len(r)>4 else "", clean_cell(r[5]) if len(r)>5 else "", clean_cell(r[6]) if len(r)>6 else "")
                except: continue
            wb.close(); print("Companies OLD: "+str(len(COMPANY_LIST)))
        except Exception as e: print("Companies OLD error: "+str(e))
    if not os.path.exists(COMPANIES_FILE_2026):
        print("2026 file not found"); return
    try:
        wb=openpyxl.load_workbook(COMPANIES_FILE_2026,read_only=True,data_only=True)
        if "خاص" in wb.sheetnames:
            ws=wb["خاص"]
            for r in ws.iter_rows(values_only=True):
                try:
                    if not r or not r[0] or not r[1]: continue
                    if str(r[0]).lower()=="registrationnumber": continue
                    ent="خاص - "+clean_cell(r[3])
                    addr="تاريخ التسجيل: "+clean_cell(r[4])
                    _add_company(r[1], "", r[0], ent, "", clean_cell(r[5]), addr)
                except: continue
        if "حكومي" in wb.sheetnames:
            ws=wb["حكومي"]
            for r in ws.iter_rows(values_only=True):
                try:
                    if not r or not r[1] or not r[2]: continue
                    if str(r[1]).lower()=="رقم الجهة": continue
                    if str(r[0])=="قطاع حكومي": continue
                    _add_company(r[2], "", r[1], "حكومي", "", "", clean_cell(r[3]))
                except: continue
        if "فرع اجنبي" in wb.sheetnames:
            ws=wb["فرع اجنبي"]
            for r in ws.iter_rows(values_only=True):
                try:
                    if not r or not r[1] or not r[2]: continue
                    if str(r[1]).lower()=="رقم التسجيل": continue
                    _add_company(r[2], "", r[1], "فرع اجنبي", "", "", clean_cell(r[3]))
                except: continue
        wb.close(); print("Companies TOTAL: "+str(len(COMPANY_LIST)))
    except Exception as e: print("Companies 2026 error: "+str(e))
build_companies_index()
def search_company(q):
    qq=(q or "").strip()
    if not qq: return []
    if qq in COMPANY_BY_REG: return [COMPANY_BY_REG[qq]]
    if qq in COMPANY_BY_FILE: return [COMPANY_BY_FILE[qq]]
    qq_digits=qq.translate(str.maketrans("٠١٢٣٤٥٦٧٨٩","0123456789")).replace(" ","").replace("،","").replace(",","")
    if qq_digits.isdigit() and len(qq_digits)>=3:
        out=[c for c in COMPANY_LIST if qq_digits in c["reg_no"] or qq_digits in c["file_no"]]
        if out: return out[:10]
    nq=norm_ar(qq); words=clean_search_words(qq); out=[]
    for idx,c in enumerate(COMPANY_LIST):
        nc=COMPANY_NORM_CACHE.get(idx,{}).get("name",norm_ar(c["name"]))
        if words and all(w in nc for w in words): out.append(c)
        elif nq and nq in nc: out.append(c)
        if len(out)>=10: break
    if out: return out
    out2=[]
    for idx,c in enumerate(COMPANY_LIST):
        cache=COMPANY_NORM_CACHE.get(idx,{})
        if nq and (nq in cache.get("activity","") or nq in cache.get("address","") or nq in cache.get("entity","")): out2.append(c)
        if len(out2)>=10: break
    if out2: return out2
    try:
        names_norm=[COMPANY_NORM_CACHE.get(i,{}).get("name","") for i in range(len(COMPANY_LIST))]
        close=difflib.get_close_matches(nq,names_norm,n=10,cutoff=0.6)
        if close:
            out3=[]
            for cn in close:
                for idx,c in enumerate(COMPANY_LIST):
                    if COMPANY_NORM_CACHE.get(idx,{}).get("name","")==cn: out3.append(c); break
            if out3: return out3[:10]
    except: pass
    return out
def format_company(c):
    txt="🏢 "+c["name"]+"\n"
    if c["reg_no"]: txt+="🔢 رقم التسجيل: "+c["reg_no"]+"\n"
    if c["file_no"]: txt+="📁 رقم الملف: "+c["file_no"]+"\n"
    if c["entity"]: txt+="🏛️ الكيان: "+c["entity"]+"\n"
    if c["activity"]: txt+="📦 النشاط: "+c["activity"]+"\n"
    if c["status"]: txt+="📄 حاله الملف: "+c["status"]+"\n"
    if c["address"]: txt+="📍 العنوان: "+c["address"]+"\n"
    return txt.strip()
def filter_by_cat(lst, cat):
    if not cat or cat=="all": return lst
    out=[]
    for c in lst:
        ent=c.get("entity","")
        if cat=="private":
            if "حكومي" not in ent and "فرع اجنبي" not in ent: out.append(c)
        elif cat=="gov":
            if "حكومي" in ent: out.append(c)
        elif cat=="foreign":
            if "فرع اجنبي" in ent: out.append(c)
    return out
def get_filtered_companies(query="", cat="all"):
    base=filter_by_cat(COMPANY_LIST, cat)
    q=norm_ar(query)
    if not q: return base
    qd=query.strip().translate(str.maketrans("٠١٢٣٤٥٦٧٨٩","0123456789")).replace(" ","")
    if qd.isdigit() and len(qd)>=3: return [c for c in base if qd in c["reg_no"] or qd in c["file_no"]]
    words=clean_search_words(query); out=[]
    for c in base:
        idx=COMPANY_LIST.index(c) if c in COMPANY_LIST else -1
        cache=COMPANY_NORM_CACHE.get(idx,{}) if idx>=0 else {"name":norm_ar(c["name"]),"activity":norm_ar(c["activity"]),"address":norm_ar(c["address"]),"entity":norm_ar(c["entity"])}
        nc=cache.get("name","")
        if words and all(w in nc for w in words): out.append(c)
        elif q in nc or q in cache.get("activity","") or q in cache.get("address","") or q in cache.get("entity",""): out.append(c)
    return out
def company_cat_title(cat):
    if cat=="private": return "الشركات الخاصة"
    if cat=="gov": return "الشركات الحكومية"
    if cat=="foreign": return "الشركات الاجنبية"
    return "كل الشركات"
def company_list_keyboard(page=0, query="", cat="all"):
    lst=get_filtered_companies(query, cat); per=10; start=page*per; end=start+per; total_pages=max(1,(len(lst)+per-1)//per); kb=[]
    for i in range(start,min(end,len(lst))):
        c=lst[i]; kb.append([InlineKeyboardButton(c["name"][:50]+" - "+c["reg_no"],callback_data="compsel_"+str(i))])
    nav=[]
    if page>0: nav.append(InlineKeyboardButton("⬅️ السابق",callback_data="comppage_"+str(page-1)))
    if page<total_pages-1: nav.append(InlineKeyboardButton("التالي ➡️",callback_data="comppage_"+str(page+1)))
    if nav: kb.append(nav)
    kb.append([InlineKeyboardButton("🏢 الشركات",callback_data="menu_companies")])
    kb.append([InlineKeyboardButton("🏠 القايمة الرئيسية",callback_data="menu_main")])
    title="بحث: "+query+" - " if query else ""
    return InlineKeyboardMarkup(kb),title+company_cat_title(cat)+" - اختار الشركة - صفحة "+str(page+1)+" من "+str(total_pages)+" - العدد "+str(len(lst))
def companies_menu_keyboard():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🏢 الشركات الخاصة",callback_data="compcat_private")],
        [InlineKeyboardButton("🏛️ الشركات الحكومية",callback_data="compcat_gov")],
        [InlineKeyboardButton("🌍 الشركات الاجنبية",callback_data="compcat_foreign")],
        [InlineKeyboardButton("🏠 القايمة الرئيسية",callback_data="menu_main")],
    ])
async def send_companies_menu(update, context):
    kb=companies_menu_keyboard()
    txt="🏢 ارقام تسجيل الشركات - اختار النوع:\n\n🏢 الشركات الخاصة - 42 الف شركة + القديم\n🏛️ الشركات الحكومية - 5 الاف جهة\n🌍 الشركات الاجنبية - فروع اجنبية"
    if update.message: await update.message.reply_text(txt,reply_markup=kb)
    elif update.callback_query:
        try: await update.callback_query.message.edit_text(txt,reply_markup=kb)
        except: await update.callback_query.message.reply_text(txt,reply_markup=kb)
async def send_company_list(update,context,page=0,query="",cat="all"):
    uid=None
    if update.message: uid=update.message.from_user.id
    elif update.callback_query: uid=update.callback_query.from_user.id
    if uid is not None: context.bot_data["comp_filter_"+str(uid)]={"query":query,"cat":cat}
    kb,txt=company_list_keyboard(page,query,cat)
    if update.message: await update.message.reply_text(txt,reply_markup=kb)
    elif update.callback_query:
        try: await update.callback_query.message.edit_text(txt,reply_markup=kb)
        except: await update.callback_query.message.reply_text(txt,reply_markup=kb)
TARIFF_LIST=[]; TARIFF_BY_CODE={}
def norm_tariff_code(s):
    if s is None: return ""
    s=str(s).strip().replace(" ","").replace("،","").replace(",","")
    if re.match(r'^\d{6}$',s): s=s[:4]+"."+s[4:]
    if re.match(r'^\d{8}$',s): s=s[:4]+"."+s[4:6]+"."+s[6:]
    return s
def clean_tariff_tax(v):
    if v is None: return ""
    s=str(v).strip()
    if s in ["","-","—","–","None"]: return ""
    if s in ["معفاه","معفاة"]: return "معفاه"
    return s.replace("%","").strip()+"%" if s.replace("%","").strip() else ""
def build_tariff_index():
    files=[]
    if os.path.isdir("tarifa"):
        for f in sorted(os.listdir("tarifa")):
            if f.lower().endswith(".xlsx"): files.append(os.path.join("tarifa",f))
    for f in sorted(os.listdir(".")):
        if f.lower().startswith("tarifa") and f.lower().endswith(".xlsx") and f not in files: files.append(f)
    for f in ["30285-1.xlsx","70490-2.xlsx"]:
        if os.path.exists(f) and f not in files: files.append(f)
    if not files: print("Tariff: no files"); return
    for fname in files:
        try:
            wb=openpyxl.load_workbook(fname,read_only=True,data_only=True); ws=wb[wb.sheetnames[0]]
            for r in ws.iter_rows(values_only=True):
                try:
                    if not r or not any([str(x).strip() if x is not None else "" for x in r]): continue
                    first=str(r[0]).strip() if r[0] is not None else ""
                    if "رمز" in first or "الفصل" in first: continue
                    if len(r)>=5 and r[2] is not None and "رمز" not in str(r[2]):
                        code=norm_tariff_code(r[2]); desc=clean_cell(r[3]); tax=clean_tariff_tax(r[4]); chapter=clean_cell(r[0]); main_band=clean_cell(r[1])
                    elif len(r)>=3 and r[0] is not None:
                        code=norm_tariff_code(r[0]); desc=clean_cell(r[1]); tax=clean_tariff_tax(r[2]); chapter=code.split(".")[0] if "." in code else code[:2]; main_band=code
                    else: continue
                    if not code or code in ["—","-","–"]: continue
                    if not desc: continue
                    item={"code":code,"desc":desc,"tax":tax,"chapter":chapter,"main_band":main_band,"file":fname}
                    if code in TARIFF_BY_CODE:
                        if not TARIFF_BY_CODE[code].get("tax") and tax: TARIFF_BY_CODE[code]=item
                        continue
                    TARIFF_BY_CODE[code]=item; TARIFF_LIST.append(item)
                except: continue
            wb.close(); print("Tariff indexed "+fname)
        except Exception as e: print("Tariff index error "+fname+": "+str(e))
    print("Tariff done: "+str(len(TARIFF_LIST)))
build_tariff_index()
def search_tariff(q):
    qq=(q or "").strip()
    if not qq: return []
    code_q=norm_tariff_code(qq)
    if code_q in TARIFF_BY_CODE: return [TARIFF_BY_CODE[code_q]]
    q_digits=re.sub(r'\D','',qq)
    if q_digits and len(q_digits)>=2:
        out=[t for