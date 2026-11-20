import os
import json
import time
import random
import re
import traceback
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
BOT = os.getenv('BOT_TOKEN')
ADM = os.getenv('ADMIN_ID')
SL = chr(47)
os.makedirs('docs', exist_ok=True)
def LJ(f, d):
    try:
        return json.load(open(f, encoding='utf-8'))
    except:
        return d
def SJ(f, d):
    json.dump(d, open(f, 'w', encoding='utf-8'))
def CC(v):
    if v is None:
        return ''
    return str(v).strip()
def GV(r, i):
    if r is None:
        return ''
    if len(r) <= i:
        return ''
    if r[i] is None:
        return ''
    return r[i]
def NA(s):
    s = str(s or '').strip().lower()
    s = s.replace('أ', 'ا')
    s = s.replace('إ', 'ا')
    s = s.replace('آ', 'ا')
    s = s.replace('ة', 'ه')
    s = s.replace('ى', 'ي')
    return s
def B(t, d):
    return InlineKeyboardButton(t, callback_data=d)
def MK(r):
    return InlineKeyboardMarkup(r)
def MAIN_KB():
    r1 = [B('توكيلات', 'm_lines'), B('GPC', 'm_gpc')]
    r2 = [B('التعريفة', 'm_tariff'), B('مستندات', 'm_docs')]
    r3 = [B('Rotation', 'm_rot'), B('شركات', 'm_comp')]
    r4 = [B('Country', 'm_country'), B('Cities', 'm_cities')]
    r5 = [B('Currency', 'm_cur'), B('Package', 'm_pack')]
    r6 = [B('UOM', 'm_uom'), B('اخر بحث', 'm_recent')]
    r7 = [B('مساعدة', 'm_help')]
    return MK([r1, r2, r3, r4, r5, r6, r7])
def COMP_KB():
    r1 = [B('الخاصة', 'ct_pri')]
    r2 = [B('الحكومية', 'ct_gov')]
    r3 = [B('الاجنبية', 'ct_for')]
    r4 = [B('القايمة', 'm_main')]
    return MK([r1, r2, r3, r4])
def LCAT_KB(k):
    r1 = [B('Booking', 'lc_' + k + '_BK')]
    r1.append(B('Invoice', 'lc_' + k + '_IN'))
    r2 = [B('الايميلات', 'lc_' + k + '_AL')]
    r3 = [B('القايمة', 'm_main')]
    return MK([r1, r2, r3])
BI = {}
BC = {}
BA = {}
for fn in ['gs1xlsx.xlsx']:
    if not os.path.exists(fn):
        continue
    wb = openpyxl.load_workbook(fn, read_only=True)
    ws = wb[wb.sheetnames[0]]
    for r in ws.iter_rows(values_only=True):
        if not r:
            continue
        r = list(r) + [''] * 15
        r = tuple(r)
        bt = str(GV(r, 10) or '')
        if not bt:
            continue
        k = bt.strip().lower()
        if k not in BI:
            d = {}
            d['c'] = str(GV(r, 9) or '')
            d['t'] = bt
            d['s'] = GV(r, 1)
            d['f'] = GV(r, 4)
            d['cl'] = GV(r, 7)
            d['i'] = GV(r, 11) or ''
            d['e'] = GV(r, 12) or ''
            BI[k] = d
            BC[d['c']] = k
        if GV(r, 13) != '':
            ac = str(GV(r, 13))
            if k not in BA:
                BA[k] = []
            ex = None
            for a in BA[k]:
                if a['c'] == ac:
                    ex = a
            if ex is None:
                ex = {'c': ac, 't': GV(r, 14), 'v': []}
                BA[k].append(ex)
    wb.close()
REF = {}
for kind in ['cities', 'country', 'currency', 'package', 'uom']:
    REF[kind] = {}
if os.path.exists('cities.xlsx'):
    wb = openpyxl.load_workbook('cities.xlsx', read_only=True)
    for kind, sh in [('cities', 'City List'), ('country', 'Country List'), ('currency', 'Currency List'), ('package', 'Package Type list'), ('uom', 'UOM List')]:
        if sh not in wb.sheetnames:
            continue
        for r in wb[sh].iter_rows(values_only=True):
            if not r or not GV(r, 0) or not GV(r, 1):
                continue
            if kind == 'currency':
                c = str(GV(r, 1)).strip().upper()
                n = str(GV(r, 0)).strip()
            else:
                c = str(GV(r, 0)).strip().upper()
                n = str(GV(r, 1)).strip()
            if c in REF[kind]:
                continue
            REF[kind][c] = {'code': c, 'name': n}
    wb.close()
CL = []
CBR = {}
def AC(n, f, rg, en, ac, st, ad, ct):
    n = CC(n)
    rg = CC(rg).replace(' ', '')
    if not n or not rg:
        return
    if rg in CBR:
        return
    it = {}
    it['name'] = n
    it['file'] = CC(f)
    it['reg'] = rg
    it['ent'] = CC(en)
    it['act'] = CC(ac)
    it['st'] = CC(st)
    it['ad'] = CC(ad)
    it['ct'] = ct
    CL.append(it)
    CBR[rg] = it
F1 = 'ارقام_تسجيل_الشركات_111111.xlsx'
if os.path.exists(F1):
    wb = openpyxl.load_workbook(F1, read_only=True)
    ws = wb[wb.sheetnames[0]]
    for r in ws.iter_rows(values_only=True):
        if not r or not GV(r, 0):
            continue
        if 'الاسم' in str(GV(r, 0)):
            continue
        AC(GV(r, 0), CC(GV(r, 1)), CC(GV(r, 2)), CC(GV(r, 3)), CC(GV(r, 4)), CC(GV(r, 5)), CC(GV(r, 6)), 'private')
    wb.close()
F2 = 'ارقام_تسجيل_الشركات_2026.xlsx'
if os.path.exists(F2):
    wb = openpyxl.load_workbook(F2, read_only=True)
    if 'خاص' in wb.sheetnames:
        for r in wb['خاص'].iter_rows(values_only=True):
            if not r or not GV(r, 0) or not GV(r, 1):
                continue
            e = 'خاص - ' + CC(GV(r, 3))
            AC(GV(r, 1), '', GV(r, 0), e, '', CC(GV(r, 5)), CC(GV(r, 4)), 'private')
    if 'حكومي' in wb.sheetnames:
        for r in wb['حكومي'].iter_rows(values_only=True):
            if not r or not GV(r, 1) or not GV(r, 2):
                continue
            AC(GV(r, 2), '', GV(r, 1), 'حكومي', '', '', CC(GV(r, 3)), 'gov')
    if 'فرع اجنبي' in wb.sheetnames:
        for r in wb['فرع اجنبي'].iter_rows(values_only=True):
            if not r or not GV(r, 1) or not GV(r, 2):
                continue
            AC(GV(r, 2), '', GV(r, 1), 'فرع اجنبي', '', '', CC(GV(r, 3)), 'foreign')
    wb.close()
def FC(c):
    t = c['name'] + '\n'
    t = t + 'رقم التسجيل: ' + c['reg'] + '\n'
    if c['file']:
        t = t + 'رقم الملف: ' + c['file'] + '\n'
    if c['ent']:
        t = t + 'الكيان: ' + c['ent'] + '\n'
    if c['act']:
        t = t + 'النشاط: ' + c['act'] + '\n'
    if c['ad']:
        t = t + 'العنوان: ' + c['ad'] + '\n'
    return t.strip()
def GC(q, ct):
    out = [c for c in CL if not ct or c['ct'] == ct]
    nq = NA(q)
    if not nq:
        return out
    o = []
    for c in out:
        if nq in NA(c['name']):
            o.append(c)
        elif nq in NA(c['act']):
            o.append(c)
    return o
TL = []
TB = {}
def NTC(s):
    s = str(s or '').strip().replace(' ', '')
    if re.match(r'^\d{6}$', s):
        s = s[:4] + '.' + s[4:]
    if re.match(r'^\d{8}$', s):
        s = s[:4] + '.' + s[4:6] + '.' + s[6:]
    return s
def load_tar(fn):
    wb = openpyxl.load_workbook(fn, read_only=True)
    ws = wb[wb.sheetnames[0]]
    for r in ws.iter_rows(values_only=True):
        if not r:
            continue
        code = NTC(GV(r, 0))
        desc = CC(GV(r, 1))
        if not code or not desc:
            continue
        if code in TB:
            continue
        it = {}
        it['code'] = code
        it['desc'] = desc
        it['tax'] = CC(GV(r, 2))
        TB[code] = it
        TL.append(it)
    wb.close()
if os.path.isdir('tarifa'):
    for f in sorted(os.listdir('tarifa')):
        if f.lower().endswith('.xlsx'):
            load_tar(os.path.join('tarifa', f))
for fn in ['30285-1.xlsx', '70490-2.xlsx']:
    if os.path.exists(fn):
        load_tar(fn)
def ST(q):
    q = (q or '').strip()
    if not q:
        return []
    cq = NTC(q)
    if cq in TB:
        return [TB[cq]]
    nq = NA(q)
    out = []
    for t in TL:
        if nq in NA(t['desc']):
            out.append(t)
            if len(out) >= 10:
                break
    return out
def FT(t):
    x = t['desc'] + '\n'
    x = x + 'البند: ' + t['code'] + '\n'
    x = x + 'الضريبة: ' + t['tax']
    return x.strip()
LC = {}
LN = {}
ROT = []
F3 = 'ايميلات_وارقام_توكيلات.xlsx'
if os.path.exists(F3):
    wb = openpyxl.load_workbook(F3, read_only=True)
    if 'EMAILS' in wb.sheetnames:
        cur = ''
        rows = list(wb['EMAILS'].iter_rows(values_only=True))
        for r in rows[2:]:
            if not r:
                continue
            if GV(r, 0):
                cur = str(GV(r, 0)).split('\n')[0].strip()
            if not cur:
                continue
            k = cur.upper()
            if k not in LN:
                LN[k] = cur
            nm = CC(GV(r, 1))
            po = CC(GV(r, 2))
            em = CC(GV(r, 3))
            nu = CC(GV(r, 4))
            bk = CC(GV(r, 5))
            inv = CC(GV(r, 9))
            if not nm and not em and not nu:
                continue
            if k not in LC:
                LC[k] = []
            d = {}
            d['name'] = nm
            d['pos'] = po
            d['email'] = em
            d['num'] = nu
            d['book'] = bk
            d['inv'] = inv
            LC[k].append(d)
    if 'SERVICE ROTATION' in wb.sheetnames:
        for r in wb['SERVICE ROTATION'].iter_rows(values_only=True):
            if not r:
                continue
            pol = CC(GV(r, 0))
            pod = CC(GV(r, 1))
            if not pol or not pod:
                continue
            d = {}
            d['pol'] = pol
            d['pod'] = pod
            d['line'] = CC(GV(r, 2))
            d['tt'] = CC(GV(r, 3))
            d['co'] = CC(GV(r, 4))
            ROT.append(d)
    wb.close()
AM = {}
AM['ميرسك'] = 'MAERSK'
AM['هاباج'] = 'HAPAG'
AM['كوسكو'] = 'COSCO'
AM['ايفرجرين'] = 'EVERGREEN'
AM['ام اس سي'] = 'MSC'
AM['زيم'] = 'ZIM'
AM['اركاس'] = 'ARKAS'
AM['وان'] = 'ONE'
def RL(q):
    q = (q or '').strip()
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
        return 'ملقيتش: ' + dp
    t = dp + '\n'
    if cat == 'BK':
        t = t + 'Booking:\n'
        for x in cs:
            t = t + x['name'] + ' - ' + x['email'] + '\n'
            t = t + x['book'] + '\n'
        return t.strip()
    if cat == 'IN':
        t = t + 'Invoice:\n'
        for x in cs:
            if x['inv']:
                t = t + x['name'] + ' - ' + x['email'] + '\n'
                t = t + x['inv'] + '\n'
        return t.strip()
    t = t + 'العدد: ' + str(len(cs)) + '\n\n'
    for c in cs[:10]:
        if c['name']:
            t = t + c['name'] + '\n'
        if c['pos']:
            t = t + c['pos'] + '\n'
        if c['email']:
            t = t + c['email'] + '\n'
        if c['num']:
            t = t + c['num'] + '\n'
        if c['book']:
            t = t + 'Booking: ' + c['book'] + '\n'
        if c['inv']:
            t = t + 'Invoice: ' + c['inv'] + '\n'
        t = t + '\n'
    return t.strip()
def SROT(pol, pod):
    out = []
    for r in ROT:
        if pol.lower() in r['pol'].lower():
            if pod.lower() in r['pod'].lower():
                out.append(r)
    return out[:10]
def FROT(rs, pol, pod):
    if not rs:
        return 'ملقيتش خط من ' + pol + ' لـ ' + pod
    t = 'من ' + pol + ' لـ ' + pod + ':\n\n'
    for r in rs:
        t = t + r['line'] + ' | ' + r['pol'] + ' -> ' + r['pod']
        if r['tt']:
            t = t + ' - ' + r['tt']
        t = t + ' | ' + r['co'] + '\n'
    return t
WEL = 'Welcome to Captain Emad Customs Bot'
HELP = 'تعريفة: 0302.44 - توكيلات: MAERSK - خط: Alex Izmir - شركات: شركات - GPC: gpc Airbrushes - مستندات: تزيل 1 - القايمة: menu'
def LU():
    return LJ('users.json', {})
def SU(u):
    SJ('users.json', u)
def OK(uid):
    if str(uid) == str(ADM):
        return True
    u = LU()
    if str(uid) not in u:
        return False
    return u[str(uid)].get('status') == 'accepted'
def FB(b):
    t = b['t'] + '\n'
    t = t + 'Brick: ' + b['c'] + '\n'
    t = t + 'Segment: ' + str(b['s']) + '\n'
    t = t + 'Family: ' + str(b['f']) + '\n'
    t = t + 'Class: ' + str(b['cl']) + '\n'
    if b['i']:
        t = t + 'Includes: ' + str(b['i']) + '\n'
    if b['e']:
        t = t + 'Excludes: ' + str(b['e']) + '\n'
    at = BA.get(b['t'].strip().lower(), [])
    if at:
        t = t + 'Attributes:\n'
        for a in at:
            t = t + str(a['t']) + ' (' + str(a['c']) + ')\n'
            if a['v']:
                t = t + ', '.join(a['v'][:10]) + '\n'
    return t
def SGPC(q):
    qq = re.sub(r'\s+', ' ', (q or '').strip().lower())
    if not qq:
        return None
    dg = re.sub(r'\D', '', qq)
    if dg and dg in BC:
        return BI[BC[dg]]
    if qq in BI:
        return BI[qq]
    m = [BI[k] for k in BI if qq in k]
    if len(m) == 1:
        return m[0]
    if len(m) > 1:
        return ('m', m[:10])
    return None
def PG(items, page, pre, lab):
    per = 10
    st = page * per
    tp = max(1, (len(items) + per - 1) // per)
    rows = []
    for i in range(st, min(st + per, len(items))):
        cb = pre + 'sel_' + str(i)
        rows.append([B(lab(items[i]), cb)])
    nav = []
    if page > 0:
        nav.append(B('السابق', pre + 'page_' + str(page - 1)))
    if page < tp - 1:
        nav.append(B('التالي', pre + 'page_' + str(page + 1)))
    if nav:
        rows.append(nav)
    rows.append([B('القايمة', 'm_main')])
    txt = 'اختار صفحة ' + str(page + 1) + ' من ' + str(tp)
    return MK(rows), txt
async def SEND(up, ctx, kind, page, q):
    if up.message:
        uid = up.message.from_user.id
    else:
        uid = up.callback_query.from_user.id
    ctx.bot_data['fq'] = q
    ctx.bot_data['fk'] = kind
    if kind == 'comp':
        ct = ctx.bot_data.get('ct', '')
        items = GC(q, ct)
        kb, txt = PG(items, page, 'comp_', lambda c: c['name'][:40] + ' - ' + c['reg'])
    elif kind == 'tariff':
        items = TL
        if q:
            items = [t for t in TL if NA(q) in NA(t['desc']) or q in t['code']]
        kb, txt = PG(items, page, 'tariff_', lambda t: t['code'] + ' - ' + t['desc'][:35])
    elif kind == 'line':
        items = sorted(LN.keys(), key=lambda x: LN[x])
        if q:
            items = [k for k in items if q.lower() in LN[k].lower()]
        kb, txt = PG(items, page, 'line_', lambda k: LN[k][:40])
    else:
        items = sorted(REF[kind].values(), key=lambda x: x['code'])
        if q:
            items = [x for x in items if q.lower() in x['name'].lower() or q.upper() in x['code']]
        kb, txt = PG(items, page, 'ref_' + kind + '_', lambda x: x['name'][:35] + ' - ' + x['code'])
    ctx.bot_data['lst'] = items
    if up.message:
        await up.message.reply_text(txt, reply_markup=kb)
    else:
        try:
            await up.callback_query.message.edit_text(txt, reply_markup=kb)
        except:
            await up.callback_query.message.reply_text(txt, reply_markup=kb)
async def START(update, ctx):
    us = update.message.from_user
    u = LU()
    uid = str(us.id)
    now = time.time()
    if uid in u and u[uid].get('status') == 'accepted':
        await update.message.reply_text(WEL, reply_markup=MAIN_KB())
        return
    if uid not in u:
        code = 'USR-' + ''.join(random.choices('0123456789', k=4))
        d = {}
        d['code'] = code
        d['name'] = us.first_name or ''
        d['status'] = 'pending'
        d['requested_at'] = now
        u[uid] = d
        SU(u)
    else:
        code = u[uid]['code']
        u[uid]['status'] = 'pending'
        SU(u)
    msg = WEL + '\nCode: ' + code
    await update.message.reply_text(msg, reply_markup=MAIN_KB())
    if ADM:
        a = B('قبول', 'accept_' + uid)
        b = B('رفض', 'reject_' + uid)
        kb = MK([[a, b]])
        try:
            await ctx.bot.send_message(chat_id=int(ADM), text='New: ' + code, reply_markup=kb)
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
    if d == 'm_main':
        await q.message.reply_text('القايمة:', reply_markup=MAIN_KB())
        return
    if d == 'm_lines':
        await SEND(q, ctx, 'line', 0, '')
        return
    if d == 'm_comp':
        kb = COMP_KB()
        await q.message.reply_text('الشركات:', reply_markup=kb)
        return
    if d == 'ct_pri':
        ctx.bot_data['ct'] = 'private'
        await SEND(q, ctx, 'comp', 0, '')
        return
    if d == 'ct_gov':
        ctx.bot_data['ct'] = 'gov'
        await SEND(q, ctx, 'comp', 0, '')
        return
    if d == 'ct_for':
        ctx.bot_data['ct'] = 'foreign'
        await SEND(q, ctx, 'comp', 0, '')
        return
    if d == 'm_tariff':
        await SEND(q, ctx, 'tariff', 0, '')
        return
    if d == 'm_gpc':
        await q.message.reply_text('اكتب: gpc Airbrushes')
        return
    if d == 'm_rot':
        await q.message.reply_text('اكتب: Alex Izmir')
        return
    if d == 'm_country':
        await SEND(q, ctx, 'country', 0, '')
        return
    if d == 'm_cities':
        await SEND(q, ctx, 'cities', 0, '')
        return
    if d == 'm_cur':
        await q.message.reply_text('اكتب: currency USD')
        return
    if d == 'm_pack':
        await q.message.reply_text('اكتب: package Box')
        return
    if d == 'm_uom':
        await q.message.reply_text('اكتب: uom KG')
        return
    if d == 'm_help':
        await q.message.reply_text(HELP, reply_markup=MAIN_KB())
        return
    if d == 'm_recent':
        lst = ctx.user_data.get('recent', [])
        if not lst:
            await q.message.reply_text('لسه مفيش بحث')
            return
        t = 'اخر بحث:\n'
        for x in lst:
            t = t + x + '\n'
        await q.message.reply_text(t)
        return
    if d == 'm_docs':
        ds = LJ('docs_index.json', [])
        if not ds:
            await q