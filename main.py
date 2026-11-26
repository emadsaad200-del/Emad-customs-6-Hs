import os, json, re, time, glob, traceback, asyncio
from datetime import datetime
BASE = os.path.dirname(os.path.abspath(__file__))
os.makedirs(os.path.join(BASE, "data"), exist_ok=True)
TK = os.environ.get("TELEGRAM_TOKEN","").strip()
if not TK:
    try: TK=open(os.path.join(BASE,"token.txt")).read().strip()
    except: TK=""
CHAT_FILE=os.path.join(BASE,"data","chats.json")
STATE_FILE=os.path.join(BASE,"data","state.json")
def LOG(m):
    print(m,flush=True)
    try:
        open(os.path.join(BASE,"data","bot.log"),"a",encoding="utf-8").write(str(m)+"\n")
    except: pass
LC={}; LN={}; ROT=[]; ROT2=[]; ROT2_SEEN=set(); SL=[]; BL=[]; PORD={}; POCOMP={}; CBM={}; CBM_LOADED=False
def SAVE_CHATS():
    try: json.dump(LC,open(CHAT_FILE,"w",encoding="utf-8"),ensure_ascii=False)
    except: pass
def LOAD_CHATS():
    global LC
    try:
        if os.path.exists(CHAT_FILE): LC=json.load(open(CHAT_FILE,encoding="utf-8"))
    except: LC={}
def SAVE_STATE():
    try: json.dump({"LN":LN,"SL":SL},open(STATE_FILE,"w",encoding="utf-8"),ensure_ascii=False)
    except: pass
def LOAD_STATE():
    global LN,SL
    try:
        if os.path.exists(STATE_FILE):
            d=json.load(open(STATE_FILE,encoding="utf-8")); LN=d.get("LN",{}); SL=d.get("SL",[])
    except: pass
LOAD_CHATS(); LOAD_STATE()
# ===== PORTS AR <-> EN =====
PORT_AR_EN={"الاسكندريه":"alexandria","الإسكندرية":"alexandria","اسكندريه":"alexandria","دمياط":"damietta","بورسعيد":"port said","السخنه":"sokhna","العين السخنه":"sokhna","الدخيله":"dekheila","ابو قير":"abu qir","شرم الشيخ":"sharm el sheikh","الغردقه":"hurghada","نويبع":"nuweiba","سفاجا":"safaga","السويس":"suez","العقبه":"aqaba","جده":"jeddah","الرياض":"riyadh","الدمام":"dammam","دبي":"dubai","جبل علي":"jebel ali","ابو ظبي":"abu dhabi","الدوحه":"doha","الكويت":"kuwait","المنامه":"manama","مسقط":"muscat","صلاله":"salalah","عدن":"aden","الحوديد":"hodeidah","جيبوتي":"djibouti","مومباسا":"mombasa","دار السلام":"dar es salaam","شنغهاي":"shanghai","نينغبو":"ningbo","شنتشن":"shenzhen","قوانغتشو":"guangzhou","تشينغداو":"qingdao","تيانجين":"tianjin","شيامن":"xiamen","هونج كونج":"hong kong","بوسان":"busan","طوكيو":"tokyo","يوكوهاما":"yokohama","سنغافوره":"singapore","بورت كلانج":"port klang","هايفونج":"haiphong","فيتنام":"vietnam","بانكوك":"bangkok","روتردام":"rotterdam","انتويرب":"antwerp","هامبورج":"hamburg","فيلكسستو":"felixstowe","ليفربول":"liverpool","مرسيليا":"marseille","جنوه":"genoa","بيروت":"beirut","طرابلس":"tripoli","بنغازي":"benghazi","الجزائر":"algiers","تونس":"tunis","الدار البيضاء":"casablanca","طنجة":"tangier","لاجوس":"lagos","نيويورك":"new york","لوس انجلوس":"los angeles","هيوستن":"houston","ميامي":"miami","سافانا":"savannah","نورفولك":"norfolk"}
PORT_EN_AR={v:k for k,v in PORT_AR_EN.items()}
def NRM(s):
    s=str(s or "").strip().lower()
    s=re.sub(r"\s+"," ",s)
    return s
def TR_PORT(s):
    s=str(s or "").strip()
    if s in PORT_AR_EN: return PORT_AR_EN[s]
    if NRM(s) in [NRM(x) for x in PORT_AR_EN]: return PORT_AR_EN[[x for x in PORT_AR_EN if NRM(x)==NRM(s)][0]]
    return s
def TR_PORT_AR(s):
    s=NRM(s)
    return PORT_EN_AR.get(s, s)
# ===== ROT OLD =====
def LOAD_ROT_OLD():
    global ROT
    ROT=[]
    for f in glob.glob(os.path.join(BASE,"*.xlsx"))+glob.glob(os.path.join(BASE,"*.csv")):
        try:
            if f.endswith(".csv"):
                import csv
                for r in csv.DictReader(open(f,encoding="utf-8",errors="ignore")):
                    vals=list(r.values())
                    if len(vals)>=3: ROT.append({"from":str(vals[0]),"to":str(vals[1]),"line":str(vals[2]),"transit":str(vals[3] if len(vals)>3 else ""),"raw":" | ".join([str(x) for x in vals])})
            else:
                try:
                    import openpyxl
                    wb=openpyxl.load_workbook(f,data_only=True,read_only=True)
                    for ws in wb.worksheets:
                        for row in ws.iter_rows(values_only=True):
                            if not row or not row[0]: continue
                            vals=[str(x) for x in row if x is not None]
                            if len(vals)>=3: ROT.append({"from":vals[0],"to":vals[1],"line":vals[2],"transit":vals[3] if len(vals)>3 else "","raw":" | ".join(vals)})
                except: pass
        except Exception as e: LOG(f"ROT old load fail {f}: {e}")
    LOG(f"ROT old: {len(ROT)}")
def FROT(rs,pol,pod):
    if not rs: return "🔄 ملقيتش خط من "+pol+" لـ "+pod
    t="🚢 *روتيشن من "+pol+" لـ "+pod+"*\n\n"
    for r in rs[:15]:
        t+="• "+r.get("raw","")+"\n"
    return t[:3500]
# ===== ROT2 NEW =====
def LOAD_ONE_ROT2_FILE(path):
    global ROT2, ROT2_SEEN
    try:
        import openpyxl
        wb=openpyxl.load_workbook(path,data_only=True,read_only=True)
        c=0
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                if not row or not row[0]: continue
                vals=[str(x).strip() for x in row if x is not None]
                if len(vals)<3: continue
                key="|".join(vals[:4]).lower()
                if key in ROT2_SEEN: continue
                ROT2_SEEN.add(key)
                ROT2.append({"from":vals[0],"to":vals[1],"line":vals[2],"transit":vals[3] if len(vals)>3 else "","raw":" | ".join(vals),"file":os.path.basename(path)})
                c+=1
        LOG(f"ROT2 loaded {c} from {os.path.basename(path)}")
    except Exception as e: LOG(f"ROT2 load fail {path}: {e}")
def LOAD_ROT2_FILE():
    global ROT2, ROT2_SEEN
    ROT2=[]; ROT2_SEEN=set()
    found=False
    for pat in ["*egypt_shipping*.xlsx","*shipping*.xlsx","*rotation*.xlsx","*.xlsx"]:
        for f in glob.glob(os.path.join(BASE,pat)):
            if "token" in f.lower(): continue
            LOAD_ONE_ROT2_FILE(f); found=True
    if not found: LOG("ROT2 FILE MISSING: no egypt_shipping file found")
    LOG(f"ROT2 new: {len(ROT2)}")
def FROT2(pol,pod):
    pol_n=NRM(TR_PORT(pol)); pod_n=NRM(TR_PORT(pod))
    out=[]
    for r in ROT2:
        if NRM(TR_PORT(r.get("from","")))==pol_n and NRM(TR_PORT(r.get("to","")))==pod_n: out.append(r)
    if not out:
        for r in ROT2:
            if pol_n in NRM(r.get("from","")) and pod_n in NRM(r.get("to","")): out.append(r)
    if not out: return ""
    t=f"🆕 *روتيشن جديد من {pol} لـ {pod}*\n\n"
    for r in out[:15]: t+="• "+r["raw"]+"\n"
    return t[:3500]
# ===== CBM =====
def LOAD_CBM():
    global CBM, CBM_LOADED
    CBM={}
    for f in glob.glob(os.path.join(BASE,"*cbm*.xlsx"))+glob.glob(os.path.join(BASE,"*CBM*.xlsx"))+glob.glob(os.path.join(BASE,"*.xlsx")):
        try:
            import openpyxl
            wb=openpyxl.load_workbook(f,data_only=True,read_only=True)
            for ws in wb.worksheets:
                for row in ws.iter_rows(values_only=True):
                    if not row or not row[0]: continue
                    k=str(row[0]).strip().lower()
                    try: CBM[k]=float(row[1])
                    except: pass
        except: pass
    CBM_LOADED=True
    LOG(f"CBM: {len(CBM)}")
def FCBM(q):
    q=NRM(q)
    for k,v in CBM.items():
        if q in k or k in q: return f"📦 *CBM {q}*\n\n• {k} : {v} CBM"
    return "📦 ملقيتش CBM لـ "+q+"، ابعت اسم الصنف بالظبط"
# ===== TELEGRAM BOT - POLLING =====
import urllib.request, urllib.parse
def TG(method, data=None):
    url=f"https://api.telegram.org/bot{TK}/{method}"
    if data: data=urllib.parse.urlencode(data).encode()
    req=urllib.request.Request(url, data=data)
    try:
        with urllib.request.urlopen(req, timeout=30) as r: return json.loads(r.read().decode())
    except Exception as e: LOG(f"TG {method} fail {e}"); return {}
def SEND(chat_id, text):
    return TG("sendMessage", {"chat_id":chat_id,"text":text[:4000],"parse_mode":"Markdown"})
def HANDLE(txt, chat_id, user=""):
    t=(txt or "").strip()
    low=t.lower()
    if t.startswith("/start") or t=="start":
        return "👋 أهلاً بيك في بوت الشحن\n\nابعت كده:\n• روتيشن اسكندريه جدة\n• روتيشن جديد اسكندريه دبي\n• CBM موبايل\n• سعر حاوية"
    if "cbm" in low:
        q=t.replace("CBM","").replace("cbm","").strip()
        return FCBM(q) if q else "ابعت: CBM + اسم الصنف"
    if "روتيشن" in t or "rotation" in low:
        m=re.search(r"(?:روتيشن|rotation)\s*(?:جديد)?\s*(.+?)\s+(.+)", t, re.I)
        if m:
            pol, pod=m.group(1).strip(), m.group(2).strip()
            old=FROT([r for r in ROT if NRM(TR_PORT(r.get("from","")))==NRM(TR_PORT(pol)) and NRM(TR_PORT(r.get("to","")))==NRM(TR_PORT(pod))], pol, pod)
            new=FROT2(pol, pod)
            if new: return new+"\n\n---\n\n"+old
            return old
        return "ابعت: روتيشن اسكندريه جدة"
    if "سعر" in t or "price" in low:
        return "💰 ابعت تفاصيل الشحنة: من - إلى - نوع البضاعة - عدد الحاويات، وهبعتلك السعر"
    return "اكتب: روتيشن اسكندريه جدة\nأو: روتيشن جديد اسكندريه دبي\nأو: CBM موبايل"
def MAIN():
    LOG("Bot starting...")
    if not TK: LOG("NO TOKEN - حط TELEGRAM_TOKEN في Secrets"); return
    LOAD_ROT_OLD(); LOAD_ROT2_FILE(); LOAD_CBM()
    LOG(f"ROT old: {len(ROT)} | ROT2 new: {len(ROT2)} | CBM: {len(CBM)}")
    offset=0
    while True:
        try:
            res=TG("getUpdates", {"offset":offset,"timeout":25})
            for u in res.get("result",[]):
                offset=u["update_id"]+1
                msg=u.get("message",{})
                txt=msg.get("text","")
                cid=str(msg.get("chat",{}).get("id",""))
                if txt and cid:
                    LC[cid]=msg.get("chat",{}).get("title",msg.get("from",{}).get("first_name",""))
                    SAVE_CHATS()
                    ans=HANDLE(txt, cid)
                    if ans: SEND(cid, ans)
        except Exception as e:
            LOG(f"MAIN loop: {e}"); time.sleep(3)
if __name__=="__main__":
    MAIN()
