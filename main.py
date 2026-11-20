import os,json,time,random,string,re,difflib
import openpyxl
from telegram import Update,InlineKeyboardButton,InlineKeyboardMarkup
from telegram.ext import ApplicationBuilder,CommandHandler,MessageHandler,CallbackQueryHandler,ContextTypes,filters
BOT=os.getenv("BOT_TOKEN")
ADM=os.getenv("ADMIN_ID")
SL=chr(47)
UF="users.json";DF="docs_index.json";SF="hits.json";OF="lines_override.json"
os.makedirs("docs",exist_ok=True)
BI={};BC={};BA={}
def bidx():
 for f in ["gs1xlsx.xlsx","GPC as of May 2023 v20230605 EN.xlsx"]:
  if not os.path.exists(f): continue
  wb=openpyxl.load_workbook(f,read_only=True,data_only=True)
  ws=wb["Schema"] if "Schema" in wb.sheetnames else wb[wb.sheetnames[0]]
  for r in ws.iter_rows(values_only=True):
   try:
    if not r or len(r)<11: continue
    if str(r[0]).lower().startswith("segment"): continue
    bc=str(r[9]) if r[9] is not None else ""
    bt=str(r[10]) if r[10] is not None else ""
    if not bt: continue
    k=bt.strip().lower()
    if k not in BI:
     BI[k]={"bc":bc,"bt":bt,"seg":r[1],"segc":str(r[0]),"fam":r[4],"famc":str(r[3]),"cls":r[7],"clsc":str(r[6]),"inc":r[11] or "","exc":r[12] if len(r)>12 else ""}
     if bc: BC[bc]=k
    if len(r)>14 and r[13] is not None:
     ac=str(r[13]);at=r[14]
     vt=str(r[17]) if len(r)>17 and r[17] is not None else None
     if k not in BA: BA[k]=[]
     ex=None
     for a in BA[k]:
      if a["c"]==ac: ex=a;break
     if not ex:
      ex={"c":ac,"t":at,"v":[]};BA[k].append(ex)
     if vt and vt not in ex["v"]: ex["v"].append(vt)
   except: continue
  wb.close()
bidx()
CIF="cities.xlsx"
CI={};CB={};COI={};COB={};CUI={};CUB={};PAI={};PAB={};UOI={};UOB={}
def ridx():
 if not os.path.exists(CIF): return
 wb=openpyxl.load_workbook(CIF,read_only=True,data_only=True)
 def put(sheet,by,ix,skip):
  if sheet not in wb.sheetnames: return
  ws=wb[sheet]
  for r in ws.iter_rows(values_only=True):
   try:
    if not r or len(r)<2 or not r[0] or not r[1]: continue
    if str(r[0]).lower() in skip: continue
    code=str(r[0]).strip().upper();name=str(r[1]).strip()
    extra=str(r[2]).strip() if len(r)>2 and r[2] else ""
    by[code]={"code":code,"name":name,"country":extra}
    k=name.lower()
    if k not in ix: ix[k]=[]
    ix[k].append({"code":code,"name":name,"country":extra})
   except: continue
 put("City List",CB,CI,["citycode"])
 put("Cities",CB,CI,["citycode"])
 put("Country List",COB,COI,["code"])
 put("Currency List",CUB,CUI,["currency"])
 put("Package Type list",PAB,PAI,["code"])
 put("UOM List",UOB,UOI,["uomcode"])
 wb.close()
ridx()
CF1="ارقام_تسجيل_الشركات_111111.xlsx"
CF2="ارقام_تسجيل_الشركات_2026.xlsx"
CL=[];CBR={};CBF={};CNC={}
def cc(v):
 return "" if v is None else str(v).strip()
def cinv(s):
 if not s: return ""
 o=[]
 for ch in str(s):
  n=ord(ch)
  if n<32: continue
  if n in [160,1600,8204,8205,8206,8207,1648]: continue
  if 8234<=n<=8238: continue
  if 1611<=n<=1626: continue
  o.append(ch)
 return "".join(o)
def nar(s):
 if not s: return ""
 s=cinv(str(s)).strip().lower()
 s=s.replace("أ","ا").replace("إ","ا").replace("آ","ا")
 s=s.replace("ة","ه").replace("ى","ي").replace("ئ","ي").replace("ؤ","و")
 s=s.replace("ھ","ه").replace("ی","ي").replace("ڪ","ك").replace("ک","ك")
 s=s.translate(str.maketrans("٠١٢٣٤٥٦٧٨٩","0123456789"))
 return re.sub(r"\s+"," ",s).strip()
def csw(t):
 st=["ال","شركه","شركة","لل","ل","شرك"]
 w=nar(t).split();o=[x for x in w if x not in st and len(x)>=2]
 return o if o else w
def addc(n,f,r,e,a,s,ad,ct=""):
 n=cc(n);r=cc(r).replace(" ","").replace(",","")
 if not n or not r: return
 if r in CBR: return
 f=cc(f).replace(" ","").replace(",","")
 it={"name":n,"file":f,"reg":r,"ent":cc(e),"act":cc(a),"st":cc(s),"ad":cc(ad),"ct":ct}
 CL.append(it);CBR[r]=it
 if f: CBF[f]=it
 CNC[len(CL)-1]={"n":nar(n),"a":nar(a),"d":nar(ad),"e":nar(e)}
def cidx():
 if os.path.exists(CF1):
  wb=openpyxl.load_workbook(CF1,read_only=True,data_only=True)
  ws=wb[wb.sheetnames[0]]
  for r in ws.iter_rows(values_only=True):
   try:
    if not r: continue
    nm=cc(r[0]) if len(r)>0 else ""
    if not nm or "الاسم" in nm: continue
    addc(nm,cc(r[1]) if len(r)>1 else "",cc(r[2]) if len(r)>2 else "",cc(r[3]) if len(r)>3 else "",cc(r[4]) if len(r)>4 else "",cc(r[5]) if len(r)>5 else "",cc(r[6]) if len(r)>6 else "","private")
   except: continue
  wb.close()
 if not os.path.exists(CF2): return
 wb=openpyxl.load_workbook(CF2,read_only=True,data_only=True)
 if "خاص" in wb.sheetnames:
  ws=wb["خاص"]
  for r in ws.iter_rows(values_only=True):
   try:
    if not r or not r[0] or not r[1]: continue
    if str(r[0]).lower()=="registrationnumber": continue
    addc(r[1],"",r[0],"خاص - "+cc(r[3]),"",cc(r[5]),"تاريخ التسجيل: "+cc(r[4]),"private")
   except: continue
 if "حكومي" in wb.sheetnames:
  ws=wb["حكومي"]
  for r in ws.iter_rows(values_only=True):
   try:
    if not r or not r[1] or not r[2]: continue
    if str(r[1]).lower()=="رقم الجهة": continue
    if str(r[0])=="قطاع حكومي": continue
    addc(r[2],"",r[1],"حكومي","","",cc(r[3]),"gov")
   except: continue
 if "فرع اجنبي" in wb.sheetnames:
  ws=wb["فرع اجنبي"]
  for r in ws.iter_rows(values_only=True):
   try:
    if not r or not r[1] or not r[2]: continue
    if str(r[1]).lower()=="رقم التسجيل": continue
    addc(r[2],"",r[1],"فرع اجنبي","","",cc(r[3]),"foreign")
   except: continue
 wb.close()
cidx()
def scomp(q,ct=""):
 qq=(q or "").strip()
 if not qq: return []
 if qq in CBR:
  c=CBR[qq]
  return [c] if not ct or c["ct"]==ct else []
 if qq in CBF:
  c=CBF[qq]
  return [c] if not ct or c["ct"]==ct else []
 qd=qq.translate(str.maketrans("٠١٢٣٤٥٦٧٨٩","0123456789")).replace(" ","").replace(",","")
 if qd.isdigit() and len(qd)>=3:
  o=[c for c in CL if (qd in c["reg"] or qd in c["file"]) and (not ct or c["ct"]==ct)]
  if o: return o[:10]
 nq=nar(qq);wd=csw(qq);o=[]
 for i,c in enumerate(CL):
  if ct and c["ct"]!=ct: continue
  nc=CNC.get(i,{}).get("n",nar(c["name"]))
  if wd and all(w in nc for w in wd): o.append(c)
  elif nq and nq in nc: o.append(c)
  if len(o)>=10: break
 return o
def fcomp(c):
 t="🏢 "+c["name"]+"\n"
 if c["reg"]: t+="🔢 رقم التسجيل: "+c["reg"]+"\n"
 if c["file"]: t+="📁 رقم الملف: "+c["file"]+"\n"
 if c["ent"]: t+="🏛️ الكيان: "+c["ent"]+"\n"
 if c["act"]: t+="📦 النشاط: "+c["act"]+"\n"
 if c["st"]: t+="📄 حاله الملف: "+c["st"]+"\n"
 if c["ad"]: t+="📍 العنوان: "+c["ad"]+"\n"
 return t.strip()
def gcomp(q="",ct=""):
 base=[c for c in CL if not ct or c["ct"]==ct]
 qq=nar(q)
 if not qq: return base
 qd=q.strip().translate(str.maketrans("٠١٢٣٤٥٦٧٨٩","0123456789")).replace(" ","")
 if qd.isdigit() and len(qd)>=3:
  return [c for c in base if qd in c["reg"] or qd in c["file"]]
 wd=csw(q);o=[]
 for c in base:
  nc=nar(c["name"])
  if wd and all(w in nc for w in wd): o.append(c)
  elif qq in nc or qq in nar(c["act"]) or qq in nar(c["ad"]): o.append(c)
 return o
def csub():
 return InlineKeyboardMarkup([
  [InlineKeyboardButton("🏢 الشركات الخاصة",callback_data="comp_type_private")],
  [InlineKeyboardButton("🏛️ الشركات الحكومية",callback_data="comp_type_gov")],
  [InlineKeyboardButton("🌍 الشركات الاجنبية",callback_data="comp_type_foreign")],
  [InlineKeyboardButton("🏠 القايمة الرئيسية",callback_data="menu_main")]])
async def send_csub(up,ctx):
 t="🏢 ارقام تسجيل الشركات - اختار النوع:"
 if up.message: await up.message.reply_text(t,reply_markup=csub())
 else:
  try: await up.callback_query.message.edit_text(t,reply_markup=csub())
  except: await up.callback_query.message.reply_text(t,reply_markup=csub())
def ckb(p=0,q="",ct=""):
 lst=gcomp(q,ct);per=10;st=p*per;en=st+per
 tp=max(1,(len(lst)+per-1)//per);kb=[]
 for i in range(st,min(en,len(lst))):
  c=lst[i]
  kb.append([InlineKeyboardButton(c["name"][:50]+" - "+c["reg"],callback_data="compsel_"+str(i))])
 nav=[]
 if p>0: nav.append(InlineKeyboardButton("⬅️ السابق",callback_data="comppage_"+str(p-1)))
 if p<tp-1: nav.append(InlineKeyboardButton("التالي ➡️",callback_data="comppage_"+str(p+1)))
 if nav: kb.append(nav)
 kb.append([InlineKeyboardButton("🔙 انواع الشركات",callback_data="menu_companies")])
 kb.append([InlineKeyboardButton("🏠 القايمة الرئيسية",callback_data="menu_main")])
 tn=""
 if ct=="private": tn="الشركات الخاصة - "
 if ct=="gov": tn="الشركات الحكومية - "
 if ct=="foreign": tn="الشركات الاجنبية - "
 ti="بحث: "+q+" - " if q else ""
 tx=tn+ti+"اختار الشركة - صفحة "+str(p+1)+" من "+str(tp)+" - العدد "+str(len(lst))
 return InlineKeyboardMarkup(kb),tx
async def send_clist(up,ctx,p=0,q="",ct=""):
 uid=up.message.from_user.id if up.message else up.callback_query.from_user.id
 ctx.bot_data["cf_"+str(uid)]={"q":q,"ct":ct}
 kb,tx=ckb(p,q,ct)
 if up.message: await up.message.reply_text(tx,reply_markup=kb)
 else:
  try: await up.callback_query.message.edit_text(tx,reply_markup=kb)
  except: await up.callback_query.message.reply_text(tx,reply_markup=kb)
TL=[];TB={}
def ntc(s):
 if s is None: return ""
 s=str(s).strip().replace(" ","").replace(",","")
 if re.match(r"^\d{6}$",s): s=s[:4]+"."+s[4:]
 if re.match(r"^\d{8}$",s): s=s[:4]+"."+s[4:6]+"."+s[6:]
 return s
def ctt(v):
 if v is None: return ""
 s=str(v).strip()
 if s in ["","-","None"]: return ""
 if s in ["معفاه","معفاة"]: return "معفاه"
 t=s.replace("%","").strip()
 return t+"%" if t else ""
def tidx():
 fs=[]
 if os.path.isdir("tarifa"):
  for f in sorted(os.listdir("tarifa")):
   if f.lower().endswith(".xlsx"): fs.append(os.path.join("tarifa",f))
 for f in ["30285-1.xlsx","70490-2.xlsx"]:
  if os.path.exists(f) and f not in fs: fs.append(f)
 for fn in fs:
  try:
   wb=openpyxl.load_workbook(fn,read_only=True,data_only=True)
   ws=wb[wb.sheetnames[0]]
   for r in ws.iter_rows(values_only=True):
    try:
     if not r: continue
     if not any([str(x).strip() if x is not None else "" for x in r]): continue
     fi=str(r[0]).strip() if r[0] is not None else ""
     if "رمز" in fi or "الفصل" in fi: continue
     code="";desc="";tax="";ch="";mb=""
     if len(r)>=5 and r[2] is not None and "رمز" not in str(r[2]):
      code=ntc(r[2]);desc=cc(r[3]);tax=ctt(r[4]);ch=cc(r[0]);mb=cc(r[1])
     elif len(r)>=3 and r[0] is not None:
      code=ntc(r[0]);desc=cc(r[1]);tax=ctt(r[2])
      ch=code.split(".")[0] if "." in code else code[:2];mb=code
     else: continue
     if not code or code=="-": continue
     if not desc: continue
     it={"code":code,"desc":desc,"tax":tax,"ch":ch,"mb":mb}
     if code in TB:
      if not TB[code].get("tax") and tax: TB[code]=it
      continue
     TB[code]=it;TL.append(it)
    except: continue
   wb.close()
  except: continue
tidx()
def star(q):
 qq=(q or "").strip()
 if not qq: return []
 cq=ntc(qq)
 if cq in TB: return [TB[cq]]
 dg=re.sub(r"\D","",qq)
 if dg and len(dg)>=2:
  o=[t for t in TL if dg in re.sub(r"\D","",t["code"])]
  if o: return o[:10]
 nq=nar(qq)
 if len(nq)<2: return []
 o=[]
 for t in TL:
  if nq in nar(t["desc"]):
   o.append(t)
   if len(o)>=10: break
 return o
def ftar(t):
 tx="🧾 "+t["desc"]+"\n🔢 البند: "+t["code"]+"\n"
 if t.get("ch"): tx+="📂 الفصل: "+t["ch"]+"\n"
 if t.get("mb") and t["mb"]!=t["code"]: tx+="📁 البند الرئيسي: "+t["mb"]+"\n"
 tx+="💰 فئة الضريبة: "+(t["tax"] if t.get("tax") else "-")+"\n"
 return tx.strip()
def gtar(q=""):
 qq=nar(q);dg=re.sub(r"\D","",q or "")
 if not qq: return TL
 return [t for t in TL if qq in nar(t["desc"]) or (dg and len(dg)>=2 and dg in re.sub(r"\D","",t["code"]))]
def tkb(p=0,q=""):
 lst=gtar(q);per=10;st=p*per;en=st+per
 tp=max(1,(len(lst)+per-1)//per);kb=[]
 for i in range(st,min(en,len(lst))):
  t=lst[i]
  kb.append([InlineKeyboardButton(t["code"]+" - "+t["desc"][:45],callback_data="tarifsel_"+str(i))])
 nav=[]
 if p>0: nav.append(InlineKeyboardButton("⬅️ السابق",callback_data="tarifpage_"+str(p-1)))
 if p<tp-1: nav.append(InlineKeyboardButton("التالي ➡️",callback_data="tarifpage_"+str(p+1)))
 if nav: kb.append(nav)
 kb.append([InlineKeyboardButton("🏠 القايمة الرئيسية",callback_data="menu_main")])
 ti="بحث: "+q+" - " if q else ""
 return InlineKeyboardMarkup(kb),ti+"اختار البند - صفحة "+str(p+1)+" من "+str(tp)+" - العدد "+str(len(lst))
async def send_tlist(up,ctx,p=0,q=""):
 uid=up.message.from_user.id if up.message else up.callback_query.from_user.id
 ctx.bot_data["tf_"+str(uid)]={"q":q}
 kb,tx=tkb(p,q)
 if up.message: await up.message.reply_text(tx,reply_markup=kb)
 else:
  try: await up.callback_query.message.edit_text(tx,reply_markup=kb)
  except: await up.callback_query.message.reply_text(tx,reply_markup=kb)
SF2="ايميلات_وارقام_توكيلات.xlsx"
LC={};LN={};ROT=[];SLM=[]
def sidx():
 if not os.path.exists(SF2): return
 wb=openpyxl.load_workbook(SF2,read_only=True,data_only=True)
 if "EMAILS" in wb.sheetnames:
  ws=wb["EMAILS"];rows=list(ws.iter_rows(values_only=True));cur=""
  for r in rows[2:]:
   try:
    if not r: continue
    if r[0]: cur=str(r[0]).split("\n")[0].strip()
    if not cur: continue
    k=cur.upper()
    if k not in LN: LN[k]=cur
    nm=cc(r[1]) if len(r)>1 else "";po=cc(r[2]) if len(r)>2 else ""
    em=cc(r[3]) if len(r)>3 else "";nu=cc(r[4]) if len(r)>4 else ""
    bk=cc(r[5]) if len(r)>5 else "";fi=cc(r[6]) if len(r)>6 else ""
    vg=cc(r[7]) if len(r)>7 else "";dr=cc(r[8]) if len(r)>8 else ""
    inv=cc(r[9]) if len(r)>9 else "";pk=cc(r[10]) if len(r)>10 else ""
    if not any([nm,po,em,nu,bk,fi,vg,dr,inv,pk]): continue
    if k not in LC: LC[k]=[]
    LC[k].append({"line":cur,"name":nm,"pos":po,"email":em,"num":nu,"bk":bk,"fi":fi,"vg":vg,"dr":dr,"inv":inv,"pk":pk})
   except: continue
 if "SERVICE ROTATION" in wb.sheetnames:
  ws=wb["SERVICE ROTATION"];lp=""
  for r in ws.iter_rows(values_only=True):
   try:
    if not r: continue
    if str(r[0]).upper()=="POL": continue
    pol=cc(r[0]) or lp
    if pol: lp=pol
    pod=cc(r[1])
    if not pol or not pod: continue
    ROT.append({"pol":pol,"pod":pod,"line":cc(r[2]),"tt":cc(r[3]),"co":cc(r[4])})
   except: continue
 if "SLMC" in wb.sheetnames:
  ws=wb["SLMC"];lp=""
  for r in ws.iter_rows(values_only=True):
   try:
    if not r: continue
    vs=[cc(x) for x in r]
    if not any(vs): continue
    if "POL" in [v.upper() for v in vs]: continue
    pol="";pod="";ln="";tt="";co=""
    if len(vs)>=4 and vs[1] and vs[2]:
     pol=vs[1];pod=vs[2];ln=vs[3]
     tt=vs[4] if len(vs)>4 else "";co=vs[5] if len(vs)>5 else ""
    elif len(vs)>=3 and vs[0] and vs[1]:
     pol=vs[0];pod=vs[1];ln=vs[2]
     tt=vs[3] if len(vs)>3 else "";co=vs[4] if len(vs)>4 else ""
    elif len(vs)>=3 and vs[2]:
     pol=lp;pod=vs[2];ln=vs[3] if len(vs)>3 else ""
    elif len(vs)>=2 and vs[1]:
     pol=lp;pod=vs[1];ln=vs[2] if len(vs)>2 else ""
    if pol: lp=pol
    if not pol or not pod: continue
    SLM.append({"pol":pol,"pod":pod,"line":ln,"tt":tt,"co":co})
   except: continue
 wb.close()
 if os.path.exists(OF):
  try:
   with open(OF,"r",encoding="utf-8") as f: ov=json.load(f)
   for k,lst in ov.items():
    if k not in LN: LN[k]=k
    if k not in LC: LC[k]=[]
    LC[k].extend(lst)
  except: pass
sidx()
AM={}
AM["ميرسك"]="MAERSK";AM["هاباج"]="HAPAG";AM["كوسكو"]="COSCO"
AM["سي ام ايه"]="CMA";AM["ايفرجرين"]="EVERGREEN";AM["ام اس سي"]="MSC"
AM["سيجو"]="SEAGO";AM["زيم"]="ZIM";AM["اركاس"]="ARKAS"
AM["وان"]="ONE";AM["او او سي ال"]="OOCL";AM["واي ام ال"]="YML"
AM["بي اي ال"]="PIL";AM["تاروس"]="TARROS";AM["توركون"]="TURKON"
AM["ميدكون"]="MEDKON";AM["ميسينا"]="MESSINA";AM["هامبورج سود"]="HAMBURG SUD"
AM["ترانسمار"]="TRANSMAR";AM["جريمالدي"]="GRIMALDI";AM["هيونداي"]="HYUNDAI"
AM["دلتا"]="DELTA";AM["ام جي ال"]="MGL";AM["ادميرال"]="ADMIRAL"
SM={"م":"MAERSK","ه":"HAPAG","ك":"COSCO","س":"CMA","ا":"EVERGREEN"}
def lh():
 if not os.path.exists(SF): return {"lines":{},"ports":{}}
 try:
  with open(SF,"r",encoding="utf-8") as f: return json.load(f)
 except: return {"lines":{},"ports":{}}
def sh(h):
 try:
  with open(SF,"w",encoding="utf-8") as f: json.dump(h,f,ensure_ascii=False,indent=2)
 except: pass
def hl(l):
 h=lh();h["lines"][l]=h["lines"].get(l,0)+1;sh(h)
def hp(p):
 h=lh();h["ports"][p]=h["ports"].get(p,0)+1;sh(h)
def rl(q):
 qq=(q or "").strip()
 if not qq: return None
 if qq.upper() in LC: return qq.upper()
 if qq.lower() in AM: return AM[qq.lower()]
 if qq.strip() in SM: return SM[qq.strip()]
 ql=qq.lower()
 for k in LN:
  if ql in k.lower() or ql in LN[k].lower(): return k
 return None
def sugl(q):
 m=difflib.get_close_matches((q or "").strip().upper(),list(LN.keys()),n=1,cutoff=0.55)
 if m: return m[0]
 m2=difflib.get_close_matches((q or "").strip().lower(),list(AM.keys()),n=1,cutoff=0.55)
 if m2: return AM[m2[0]]
 return None
def arec(ctx,t):
 lst=ctx.user_data.get("recent",[]);t=(t or "").strip()
 if t and t not in lst: lst.insert(0,t)
 ctx.user_data["recent"]=lst[:5]
def mmk():
 return InlineKeyboardMarkup([
  [InlineKeyboardButton("🚢 توكيلات",callback_data="menu_lines"),InlineKeyboardButton("📦 GPC",callback_data="menu_gpc")],
  [InlineKeyboardButton("🧾 التعريفة الجمركية",callback_data="menu_tariff"),InlineKeyboardButton("📁 المستندات / تزيل",callback_data="menu_docs")],
  [InlineKeyboardButton("🚢 Rotation",callback_data="menu_rotation"),InlineKeyboardButton("🏢 ارقام تسجيل الشركات",callback_data="menu_companies")],
  [InlineKeyboardButton("🌍 Country",callback_data="menu_country"),InlineKeyboardButton("🏙️ Cities",callback_data="menu_cities")],
  [InlineKeyboardButton("💵 Currency",callback_data="menu_currency"),InlineKeyboardButton("📦 Package",callback_data="menu_package")],
  [InlineKeyboardButton("📏 UOM",callback_data="menu_uom"),InlineKeyboardButton("🕘 اخر بحث",callback_data="menu_recent")],
  [InlineKeyboardButton("❓ مساعدة / Help",callback_data="menu_help")]])
def lck(k):
 return InlineKeyboardMarkup([
  [InlineKeyboardButton("📘 Booking",callback_data="linecat_"+k+"_BOOKING"),InlineKeyboardButton("🧾 Invoice",callback_data="linecat_"+k+"_INVOICE")],
  [InlineKeyboardButton("📧 كل الايميلات",callback_data="linecat_"+k+"_ALL"),InlineKeyboardButton("📞 كل الارقام",callback_data="linecat_"+k+"_PHONE")],
  [InlineKeyboardButton("🏠 القايمة الرئيسية",callback_data="menu_main")]])
def flc(k,cat="ALL"):
 cs=LC.get(k,[]);dp=LN.get(k,k)
 if not cs: return "❌ ملقيتش توكيل اسمه: "+dp
 tx="🚢 "+dp+"\n"
 if cat=="BOOKING":
  tx+="📘 Booking:\n\n";c=0
  for x in cs:
   if x["bk"] or x["email"]:
    tx+="• "+(x["name"] or x["pos"] or "")+" - "+(x["email"] or "")+" "+(x["num"] or "")+"\nBooking: "+x["bk"]+"\n\n";c+=1
  if c==0: tx+="مفيش Booking مسجل\n"
  return tx.strip()
 if cat=="INVOICE":
  tx+="🧾 Invoice:\n\n"
  for x in cs:
   if x["inv"]: tx+="• "+(x["name"] or x["pos"] or "")+" - "+x["email"]+"\nInvoice: "+x["inv"]+"\n\n"
  return tx.strip() or ("مفيش Invoice لـ "+dp)
 if cat=="PHONE":
  tx+="📞 الارقام:\n\n"
  for x in cs:
   if x["num"]: tx+="• "+(x["name"] or x["pos"] or "")+": "+x["num"]+"\n"
  return tx.strip()
 tx+="📞 عدد جهات الاتصال: "+str(len(cs))+"\n\n"
 for c in cs[:12]:
  if c["name"]: tx+="👤 "+c["name"]+"\n"
  if c["pos"]: tx+="💼 "+c["pos"]+"\n"
  if c["email"]: tx+="📧 "+c["email"]+"\n"
  if c["num"]: tx+="📞 "+c["num"]+"\n"
  if c["bk"]: tx+="📘 Booking: "+c["bk"]+"\n"
  if c["inv"]: tx+="🧾 Invoice: "+c["inv"]+"\n"
  tx+="\n"
 if len(cs)>12: tx+="... و "+str(len(cs)-12)+" تانيين\n"
 return tx.strip()
def cpk(k):
 cs=LC.get(k,[]);kb=[]
 for i,c in enumerate(cs[:5]):
  if c["email"]: kb.append([InlineKeyboardButton("📋 نسخ "+c["email"][:30],callback_data="copy_"+k+"_"+str(i))])
 kb.append([InlineKeyboardButton("🏠 القايمة",callback_data="menu_main")])
 return InlineKeyboardMarkup(kb)
def srot(pol,pod):
 pl=pol.lower().strip();dl=pod.lower().strip();o=[]
 for r in ROT:
  if pl in r["pol"].lower() and dl in r["pod"].lower(): o.append(r)
 for r in SLM:
  if pl in r["pol"].lower() and dl in r["pod"].lower():
   if r not in o: o.append(r)
 return o[:10]
def frot(rs,pol,pod):
 if not rs: return "❌ ملقيتش خط من "+pol+" لـ "+pod
 tx="🚢 من "+pol+" لـ "+pod+" - لقيت "+str(len(rs))+":\n\n"
 for r in rs:
  tt=" - T.T: "+r["tt"] if r["tt"] else ""
  tx+="• "+r["line"]+" | "+r["pol"]+" -> "+r["pod"]+tt+" | "+r["co"]+"\n"
 return tx
def glines(q=""):
 lst=sorted(LN.keys(),key=lambda x: LN[x]);qq=(q or "").strip().lower()
 if not qq: return lst
 return [k for k in lst if qq in LN[k].lower() or qq in k.lower()]
def lkb(p=0,q=""):
 lst=glines(q);per=10;st=p*per;en=st+per
 tp=max(1,(len(lst)+per-1)//per);kb=[]
 for i in range(st,min(en,len(lst))):
  k=lst[i]
  kb.append([InlineKeyboardButton(LN[k][:55],callback_data="linesel_"+str(i))])
 nav=[]
 if p>0: nav.append(InlineKeyboardButton("⬅️ السابق",callback_data="linepage_"+str(p-1)))
 if p<tp-1: nav.append(InlineKeyboardButton("التالي ➡️",callback_data="linepage_"+str(p+1)))
 if nav: kb.append(nav)
 kb.append([InlineKeyboardButton("🏠 القايمة الرئيسية",callback_data="menu_main")])
 ti="بحث: "+q+" - " if q else ""
 return InlineKeyboardMarkup(kb),ti+"اختار التوكيل - صفحة "+str(p+1)+" من "+str(tp)+" - العدد "+str(len(lst))
async def send_llist(up,ctx,p=0,q=""):
 uid=up.message.from_user.id if up.message else up.callback_query.from_user.id
 ctx.bot_data["lf_"+str(uid)]={"q":q}
 kb,tx=lkb(p,q)
 if up.message: await up.message.reply_text(tx,reply_markup=kb)
 else:
  try: await up.callback_query.message.edit_text(tx,reply_markup=kb)
  except: await up.callback_query.message.reply_text(tx,reply_markup=kb)
def ldocs():
 if not os.path.exists(DF): return []
 try:
  with open(DF,"r",encoding="utf-8") as f: return json.load(f)
 except: return []
def sdocs(d):
 with open(DF,"w",encoding="utf-8") as f: json.dump(d,f,ensure_ascii=False,indent=2)
WEL="Welcome to Captain Emad Customs Bot. After the trial and subscription for 150 EGP, you can inquire about the customs item by code or by item name in Arabic or English. You can also search for GS1, and you can calculate the approximate customs."
HELP="مساعدة - تعريفة اكتب 0302.44 او خيول - توكيلات اكتب توكيلات او MAERSK - خط اكتب Alex Izmir - شركات اكتب شركات او رقم التسجيل - GPC اكتب gpc Airbrushes - Country اكتب Country Eg - Cities اكتب Cities Cai - Currency اكتب Currency US - Package اكتب Package B - UOM اكتب UOM K - مستندات اكتب docs او تزيل 1 - اخر بحث اكتب اخر - القايمة اكتب menu"
def lu():
 if not os.path.exists(UF): return {}
 try:
  with open(UF,"r",encoding="utf-8") as f: return json.load(f)
 except: return {}
def su(u):
 with open(UF,"w",encoding="utf-8") as f: json.dump(u,f,ensure_ascii=False,indent=2)
def gcode(us):
 while True:
  c="USR-"+"".join(random.choices(string.digits,k=4))
  if all(u.get("code")!=c for u in us.values()): return c
def ctxt(t):
 if not t: return ""
 t=cinv(t)
 for i,d in enumerate("٠١٢٣٤٥٦٧٨٩"): t=t.replace(d,str(i))
 return t.replace("أ","ا").replace("إ","ا").replace("آ","ا").replace("ة","ه").strip()
def isdl(t):
 return bool(re.search(r"(تزيل|تنزيل|تحميل|download|نزل)",ctxt(t)))
def exn(t):
 m=re.search(r"(\d+)",ctxt(t))
 return int(m.group(1)) if m else None
def norm(t): return re.sub(r"\s+"," ",(t or "").strip().lower())
def sgpc(q):
 qq=norm(q)
 if not qq: return None,None
 dg=re.sub(r"\D","",qq)
 if dg and dg in BC: return BI[BC[dg]],None
 if qq in BI: return BI[qq],None
 ms=[BI[k] for k in BI.keys() if qq in k]
 if len(ms)==1: return ms[0],None
 if len(ms)>1: return ("multiple",ms[:10]),None
 return None,None
def gattr(bt): return BA.get(bt.strip().lower(),[])
def fbrick(b,at):
 tx="📦 "+b["bt"]+"\n🔢 Brick Code: "+b["bc"]+"\n\n"
 tx+="📂 Segment: "+str(b["seg"])+" ("+str(b["segc"])+")\n"
 tx+="📁 Family: "+str(b["fam"])+" ("+str(b["famc"])+")\n"
 tx+="📄 Class: "+str(b["cls"])+" ("+str(b["clsc"])+")\n\n"
 if b.get("inc"): tx+="✅ Includes:\n"+str(b["inc"])+"\n\n"
 if b.get("exc"): tx+="⛔ Excludes:\n"+str(b["exc"])+"\n\n"
 if at:
  tx+="⚙️ Attributes:\n"
  for a in at:
   tx+="\n• "+str(a["t"])+" ("+str(a["c"])+")\n"
   if a.get("v"):
    vs=a["v"][:15];tx+="  "+ ", ".join(vs)
    if len(a["v"])>15: tx+=" ... (+"+str(len(a["v"])-15)+" more)"
    tx+="\n"
 return tx
def sref(kind,q):
 qq=q.strip();qu=qq.upper().strip()
 mp={"cities":(CB,CI),"country":(COB,COI),"currency":(CUB,CUI),"package":(PAB,PAI),"uom":(UOB,UOI)}
 if kind not in mp: return None,None
 by,ix=mp[kind]
 if qu in by: return by[qu],None
 ql=qq.lower();ms=[]
 for k,lst in ix.items():
  if ql in k: ms.extend(lst)
  if len(ms)>=10: break
 if not ms:
  for code,it in by.items():
   if qu in code: ms.append(it)
   if len(ms)>=10: break
 if len(ms)==1: return ms[0],None
 if len(ms)>1: return ("multiple",ms[:10]),None
 return None,None
def fref(kind,it):
 if kind=="cities":
  ex="\n🌍 Country: "+str(it.get("country","")) if it.get("country") else ""
  return "🏙️ "+it["name"]+"\n🔢 City Code: "+it["code"]+ex
 if kind=="country": return "🌍 "+it["name"]+"\n🔢 Country Code: "+it["code"]
 if kind=="currency": return "💵 "+it["name"]+"\n🔢 Currency Code: "+it["code"]
 if kind=="package": return "📦 "+it["name"]+"\n🔢 Package Code: "+it["code"]
 if kind=="uom": return "📏 "+it["name"]+"\n🔢 UOM Code: "+it["code"]
 return it["name"]+" - "+it["code"]
def saref(q):
 rs=[];qq=q.strip();qu=qq.upper().strip();ql=qq.lower()
 if qu in CB: rs.append(("cities",CB[qu]))
 if qu in COB: rs.append(("country",COB[qu]))
 if qu in CUB: rs.append(("currency",CUB[qu]))
 if qu in PAB: rs.append(("package",PAB[qu]))
 if qu in UOB: rs.append(("uom",UOB[qu]))
 if rs: return rs
 pairs=[("cities",CI),("country",COI),("currency",CUI),("package",PAI),("uom",UOI)]
 for kind,ix in pairs:
  for k,lst in ix.items():
   if ql in k:
    for it in lst: rs.append((kind,it))
    if len(rs)>=10: return rs
 return rs
def gsref(kind):
 mp={"cities":CB,"country":COB,"currency":CUB,"package":PAB,"uom":UOB}
 return sorted(mp.get(kind,{}).values(),key=lambda x: x["code"])
def gfref(kind,q=""):
 lst=gsref(kind);qq=(q or "").strip().lower()
 if not qq: return lst
 qu=qq.upper()
 return [it for it in lst if qq in it["name"].lower() or qu in it["code"]]
def rkb(kind,p=0,q=""):
 lst=gfref(kind,q);per=10;st=p*per;en=st+per
 tp=max(1,(len(lst)+per-1)//per);kb=[]
 for i in range(st,min(en,len(lst))):
  it=lst[i]
  kb.append([InlineKeyboardButton(it["name"][:40]+" - "+it["code"],callback_data="refsel_"+kind+"_"+str(i))])
 nav=[]
 if p>0: nav.append(InlineKeyboardButton("⬅️ السابق",callback_data="refpage_"+kind+"_"+str(p-1)))
 if p<tp-1: nav.append(InlineKeyboardButton("التالي ➡️",callback_data="refpage_"+kind+"_"+str(p+1)))
 if nav: kb.append(nav)
 kb.append([InlineKeyboardButton("🏠 القايمة الرئيسية",callback_data="menu_main")])
 ti="بحث: "+q+" - " if q else ""
 return InlineKeyboardMarkup(kb),ti+"اختار من "+kind+" - صفحة "+str(p+1)+" من "+str(tp)+" - العدد "+str(len(lst))
async def send_rlist(up,ctx,kind,p=0,q=""):
 uid=up.message.from_user.id if up.message else up.callback_query.from_user.id
 ctx.bot_data["rf_"+str(uid)]={"kind":kind,"q":q}
 kb,tx=rkb(kind,p,q)
 if up.message: await up.message.reply_text(tx,reply_markup=kb)
 else:
  try: await up.callback_query.message.edit_text(tx,reply_markup=kb)
  except: await up.callback_query.message.reply_text(tx,reply_markup=kb)
async def start(up:Update,ctx:ContextTypes.DEFAULT_TYPE):
 u=up.message.from_user;us=lu();uid=str(u.id);now=time.time()
 if uid in us and us[uid].get("status")=="rejected":
  last=us[uid].get("rejected_at",0);rem=int((5*60-(now-last))/60)+1
  if now-last<5*60:
   kb=InlineKeyboardMarkup([[InlineKeyboardButton("🔁 اطلب قبول تاني",callback_data="request_again")]])
   await up.message.reply_text("Your request was rejected. You can send again after "+str(rem)+" minutes.",reply_markup=kb);return
  us[uid]["status"]="pending";us[uid]["requested_at"]=now;su(us)
 if uid in us and us[uid].get("status")=="accepted":
  await up.message.reply_text(WEL+"\n\n✅ Your account is accepted.",reply_markup=mmk());return
 if uid not in us:
  code=gcode(us);us[uid]={"code":code,"name":u.first_name or "","username":u.username or "","id":u.id,"status":"pending","requested_at":now};su(us)
 else:
  code=us[uid]["code"];us[uid]["status"]="pending";us[uid]["requested_at"]=now;su(us)
 await up.message.reply_text(WEL+"\n\nYour Code: "+code+"\nWaiting for acceptance from Captain Emad.",reply_markup=mmk())
 if ADM:
  kb=[[InlineKeyboardButton("✅ قبول / Accept",callback_data="accept_"+uid),InlineKeyboardButton("❌ رفض / Reject",callback_data="reject_"+uid)]]
  try: await ctx.bot.send_message(chat_id=int(ADM),text="📦 New User Request\n\nCode: "+code+"\nName: "+str(u.first_name)+"\nUsername: @"+str(u.username)+"\nID: "+str(u.id),reply_markup=InlineKeyboardMarkup(kb))
  except Exception as e: print(e)
async def helpc(up:Update,ctx:ContextTypes.DEFAULT_TYPE):
 await up.message.reply_text(HELP,reply_markup=mmk())
async def menuc(up:Update,ctx:ContextTypes.DEFAULT_TYPE):
 await up.message.reply_text("🏠 القايمة الرئيسية - اختار:",reply_markup=mmk())
async def topc(up:Update,ctx:ContextTypes.DEFAULT_TYPE):
 if str(up.message.from_user.id)!=str(ADM): return
 h=lh();ls=sorted(h["lines"].items(),key=lambda x: x[1],reverse=True)[:10]
 ps=sorted(h["ports"].items(),key=lambda x: x[1],reverse=True)[:10]
 tx="📊 اكتر حاجة اتسأل عليها:\n\n🚢 خطوط:\n"
 for k,v in ls: tx+="• "+k+": "+str(v)+"\n"
 tx+="\n⚓ موانئ:\n"
 for k,v in ps: tx+="• "+k+": "+str(v)+"\n"
 await up.message.reply_text(tx or "لسه مفيش بحث")
async def btn(up:Update,ctx:ContextTypes.DEFAULT_TYPE):
 q=up.callback_query
 try: await q.answer()
 except: pass
 us=lu();d=q.data
 if d=="menu_main":
  await q.message.reply_text("🏠 القايمة الرئيسية - اختار:",reply_markup=mmk());return
 if d=="menu_lines":
  await send_llist(q,ctx,0,"");return
 if d=="menu_companies":
  await send_csub(q,ctx);return
 if d=="comp_type_private":
  await send_clist(q,ctx,0,"","private");return
 if d=="comp_type_gov":
  await send_clist(q,ctx,0,"","gov");return
 if d=="comp_type_foreign":
  await send_clist(q,ctx,0,"","foreign");return
 if d=="menu_tariff":
  await send_tlist(q,ctx,0,"");return
 if d=="menu_gpc":
  await q.message.reply_text("اكتب: gpc Airbrushes او اسم الصنف",reply_markup=mmk());return
 if d=="menu_country":
  await send_rlist(q,ctx,"country",0,"");return
 if d=="menu_cities":
  await send_rlist(q,ctx,"cities",0,"");return
 if d=="menu_currency":
  await send_rlist(q,ctx,"currency",0,"");return
 if d=="menu_package":
  await send_rlist(q,ctx,"package",0,"");return
 if d=="menu_uom":
  await send_rlist(q,ctx,"uom",0,"");return
 if d=="menu_rotation":
  await q.message.reply_text("اكتب: Alex Izmir",reply_markup=mmk());return
 if d=="menu_docs":
  ds=ldocs()
  if not ds:
   await q.message.reply_text("مفيش مستندات لسه",reply_markup=mmk());return
  tx="📁 المستندات:\n\n"
  for x in ds: tx+=str(x["num"])+" - "+x["name"]+"\n"
  tx+="\nاكتب: تزيل 1";await q.message.reply_text(tx,reply_markup=mmk());return
 if d=="menu_help":
  await q.message.reply_text(HELP,reply_markup=mmk());return
 if d=="menu_recent":
  lst=ctx.user_data.get("recent",[])
  if not lst:
   await q.message.reply_text("لسه مفيش بحث");return
  await q.message.reply_text("🕘 اخر 5:\n"+"\n".join(["• "+x for x in lst]));return
 if d=="request_again":
  uid=str(q.from_user.id);us2=lu()
  if uid in us2:
   us2[uid]["status"]="pending";us2[uid]["requested_at"]=time.time();su(us2)
   await q.message.reply_text("✅ طلبك اتبعت تاني لكابتن عماد، استنى القبول")
   if ADM: await ctx.bot.send_message(chat_id=int(ADM),text="🔁 طلب قبول تاني\nCode: "+us2[uid]["code"]+" - "+us2[uid]["name"]+" - ID: "+uid)
  return
 if d.startswith("tarifpage_"):
  p=int(d.split("_")[1]);f=ctx.bot_data.get("tf_"+str(q.from_user.id),{});qq=f.get("q","")
  kb,tx=tkb(p,qq)
  try: await q.message.edit_text(tx,reply_markup=kb)
  except: await q.message.reply_text(tx,reply_markup=kb)
  return
 if d.startswith("tar