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
    x = x + 'البند: ' +