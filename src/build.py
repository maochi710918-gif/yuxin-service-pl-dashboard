# -*- coding: utf-8 -*-
"""服務部損益 Data Model builder: Excel -> Fact/Dim tables -> data.json + 資料模型.xlsx"""
import openpyxl, json, re, sys
import pandas as pd, os, shutil, tempfile
HERE = os.path.dirname(os.path.abspath(__file__))

SRC = sys.argv[1] if len(sys.argv) > 1 else r'C:\Users\ES\OneDrive\桌面\AI參照\2026損益表 的複本1001.xlsx'
_tmp = os.path.join(tempfile.gettempdir(), '_svc_pl_src.xlsx'); shutil.copyfile(SRC, _tmp)  # 避免 Excel 開啟中被鎖定
SRC_NAME = os.path.basename(SRC); SRC = _tmp
MONTH_SHEETS = ['2601', '2602', '2603', '2604', '2605', '2606', '2607', '2608']
NM = len(MONTH_SHEETS)
wb = openpyxl.load_workbook(SRC, data_only=True)
# 單位清單：直接讀 2601 表頭（新增單位會自動帶入）
_hdrs = [[str(h.value).strip() for h in wb[sh][1][2:] if h.value is not None and str(h.value).strip()] for sh in MONTH_SHEETS]
BR = [b for b in _hdrs[0] if all(b in h for h in _hdrs)]
SKIPPED_COLS = sorted({(b, MONTH_SHEETS[k]) for k, h in enumerate(_hdrs) for b in h if b not in BR})
HQ = '服務部本部'  # 服務部總表 − 各單位加總 (未分攤之收入/費用)
ENT = BR + [HQ]
TYPE = ['鈑噴中心' if '鈑噴' in b else '服務廠' for b in BR]
ws = wb['2026服務部'] if '2026服務部' in wb.sheetnames else wb['服務部']
ws5 = wb['2025服務部'] if '2025服務部' in wb.sheetnames else None
PRIOR_SHEETS = [f'25{m:02d}' for m in range(1, 13) if f'25{m:02d}' in wb.sheetnames] if ws5 else []
NP = len(PRIOR_SHEETS)
num = lambda x: float(x) if isinstance(x, (int, float)) else 0.0
code_s = lambda c: str(c).strip() if c is not None else ''

# 管理層級 (科目代號 -> level)；其餘 indent0 列 = 所屬管理層級 +1，indent1 列 = 上一群組 +1
MG = {'100': 1, '4108': 2, '4111': 2, '410804': 2, '110': 1, '5102': 2, '510204': 2, '11010': 2,
      '120': 0, '130': 1, '140': 1, '141': 2, '6117': 2, '150': 1, '510302': 2, '520302': 2,
      '160': 1, '16010': 2, '16020': 2, '510210': 1, '170': 1, '490': 1, '49010': 2, '4901010': 3,
      '4901020': 3, '4901030': 3, '49020': 2, '4902030': 3, '49030': 2, '500': 0, '510': 1,
      '51010': 2, '51020': 2, '5102020': 3, '520': 0, '90': 0}
SKIP = {'690', '714405', '480'}  # 週邊業務重複列 (=170)，只保留 170

# ---------- 1. 讀服務部總表 (科目樹 + 預算/去年) ----------
accts = []
stack = []  # (level, id)
start, end = 166, 1426
last_mg = None; last_grp = None; used_mg = set()
for i in range(start, end + 1):
    c = code_s(ws.cell(i, 5).value); name = ws.cell(i, 6).value
    if ws.row_dimensions[i].hidden and (c not in MG or c in used_mg):
        continue
    if not c or name is None or c in SKIP:
        continue
    months = [num(ws.cell(i, 7 + k).value) for k in range(NM)]
    b26 = num(ws.cell(i, 25).value); a25 = num(ws.cell(i, 27).value); a24 = num(ws.cell(i, 29).value)
    indent = ws.cell(i, 6).alignment.indent or 0
    mrow = ws.cell(i, 4).value  # 月表列號
    if c in MG and c not in used_mg:
        used_mg.add(c); lvl = MG[c]; last_mg = (lvl, len(accts)); last_grp = (lvl, len(accts))
    elif indent == 0:
        lvl = last_mg[0] + 1; last_grp = (lvl, len(accts))
    else:
        lvl = last_grp[0] + 1
    # parent = 最近一個 level < lvl 的列
    parent = None
    for j in range(len(accts) - 1, -1, -1):
        if accts[j]['lvl'] < lvl:
            parent = j; break
    if lvl == 0: parent = None
    accts.append(dict(id=len(accts), code=c, name=str(name).strip(), lvl=lvl, parent=parent,
                      mrow=mrow, svc=months, b26=b26, a25=a25, a24=a24, xrow=i))

# ---------- 2. 讀各月廠別矩陣 (Wide -> 值陣列) ----------
recon = []
for a in accts:
    a['v'] = [[0.0] * NM for _ in ENT]; a['blank'] = set()
for mi, sh in enumerate(MONTH_SHEETS):
    rows = list(wb[sh].iter_rows(values_only=True))
    hdr = [str(h).strip() if h else '' for h in rows[0]]
    col = {b: hdr.index(b) for b in BR}
    code_index = {}
    for ri, r in enumerate(rows):
        code_index.setdefault(code_s(r[0]), ri)
    for a in accts:
        ri = (a['mrow'] - 1) if isinstance(a['mrow'], int) else None
        if ri is None or ri >= len(rows) or code_s(rows[ri][0]) != a['code']:
            ri = code_index.get(a['code'])
        r = rows[ri] if ri is not None else None
        s = 0.0
        for bi, b in enumerate(BR):
            raw = r[col[b]] if r and col[b] < len(r) else None
            if raw is None: a['blank'].add((bi, mi))
            val = num(raw)
            a['v'][bi][mi] = val; s += val
        a['v'][len(BR)][mi] = a['svc'][mi] - s  # 本部 = 總表 − 各廠

# ---------- 2b. 2025 月別 × 單位（同期比較用）----------
def read_month_matrix(sheets, svc_ws, store):
    for a in accts:
        a[store] = [[0.0] * len(sheets) for _ in ENT]; a[store + '_blank'] = set()
    for mi, sh in enumerate(sheets):
        rows = list(wb[sh].iter_rows(values_only=True))
        hdr = [str(h).strip() if h else '' for h in rows[0]]
        col = {b: hdr.index(b) for b in BR if b in hdr}
        code_index = {}
        for ri, r in enumerate(rows): code_index.setdefault(code_s(r[0]), ri)
        for a in accts:
            ri = (a['mrow'] - 1) if isinstance(a['mrow'], int) else None
            if ri is None or ri >= len(rows) or code_s(rows[ri][0]) != a['code']: ri = code_index.get(a['code'])
            r = rows[ri] if ri is not None else None
            tot = 0.0
            for bi, b in enumerate(BR):
                raw = r[col[b]] if (r and b in col and col[b] < len(r)) else None
                if raw is None and b in col: a[store + '_blank'].add((bi, mi))
                v = num(raw); a[store][bi][mi] = v; tot += v
            svc = num(svc_ws.cell(a['xrow'], 7 + mi).value)
            a[store][len(BR)][mi] = svc - tot
if NP:
    read_month_matrix(PRIOR_SHEETS, ws5, 'py')
else:
    for a in accts: a['py'] = None

kpis = {}
KPI_ROWS = {'保修進廠台數': 1453, '有費台數(不含補償)': 1454, '補償台數': 1455, '服務廠人數': 1462,
            '服務廠直接人數': 1463, '服務廠間接人數': 1464}
for k, xr in KPI_ROWS.items():
    mrow = ws.cell(xr, 4).value
    svc = [num(ws.cell(xr, 7 + m).value) for m in range(NM)]
    vals = []
    for mi, sh in enumerate(MONTH_SHEETS):
        wsm = wb[sh]; hdr = [str(h.value).strip() if h.value else '' for h in wsm[1]]
        vals.append([num(wsm.cell(mrow, hdr.index(b) + 1).value) for b in BR])
    py = None
    if NP:
        pv = []
        for sh in PRIOR_SHEETS:
            wsm = wb[sh]; hdr = [str(h.value).strip() if h.value else '' for h in wsm[1]]
            pv.append([num(wsm.cell(mrow, hdr.index(b) + 1).value) if b in hdr else 0.0 for b in BR])
        py = dict(svc=[num(ws5.cell(xr, 7 + m).value) for m in range(NP)], br=[[pv[m][b] for m in range(NP)] for b in range(len(BR))])
    kpis[k] = dict(svc=svc, br=[[vals[m][b] for m in range(NM)] for b in range(len(BR))], py=py,
                   b26=num(ws.cell(xr, 25).value), a25=num(ws.cell(xr, 27).value))

# 廠別公式勾稽：空白小計列依公式補算並記錄；非空白但不符只警告、不調整
fills = []
def _a(code): return next(a for a in accts if a['code'] == code)
FORM = [('120', '零服毛利 = 營收 − 成本', lambda g: g('100') - g('110')),
        ('90', '零服業務貢獻 = 毛利 + 獎金 − 外促 − 內促 − 固薪', lambda g: g('120') + g('130') - g('140') - g('150') - g('160')),
        ('500', '營業利益 = 業務貢獻 + 週邊 − 費用', lambda g: g('90') + g('170') - g('490')),
        ('520', '稅前淨利 = 營業利益 + 業外淨收入', lambda g: g('500') + g('510'))]
for bi, b in enumerate(BR):
    for mi in range(NM):
        g = lambda code: _a(code)['v'][bi][mi]
        for code, label, f in FORM:
            a = _a(code); calc = f(g); book = a['v'][bi][mi]
            if (bi, mi) in a['blank'] and abs(calc) > 0.5:
                a['v'][bi][mi] = calc
                fills.append(dict(type='公式補算', item=f"{b}｜{code} {a['name']}（Excel 儲存格空白，依「{label}」補算）", month=mi + 1, book=0.0, detail=calc, diff=-calc))
            elif abs(book - calc) > 1 and (bi, mi) not in a['blank']:
                fills.append(dict(type='廠別公式差異', item=f"{b}｜{label}", month=mi + 1, book=book, detail=calc, diff=book - calc))
if NP:
    for bi, b in enumerate(BR):
        for mi in range(NP):
            g = lambda code: _a(code)['py'][bi][mi]
            for code, label, f in FORM:
                a = _a(code); calc = f(g); book = a['py'][bi][mi]
                if (bi, mi) in a['py_blank'] and abs(calc) > 0.5:
                    a['py'][bi][mi] = calc
                    fills.append(dict(type='公式補算(2025)', item=f"{b}｜{code} {a['name']}（2025/{mi+1} 儲存格空白，依「{label}」補算）", month=mi + 1, book=0.0, detail=calc, diff=-calc))
                elif abs(book - calc) > 1 and (bi, mi) not in a['py_blank']:
                    fills.append(dict(type='廠別公式差異(2025)', item=f"{b}｜{label}", month=mi + 1, book=book, detail=calc, diff=book - calc))
    for a in accts:
        for mi in range(NP):
            a['py'][len(BR)][mi] = num(ws5.cell(a['xrow'], 7 + mi).value) - sum(a['py'][bi][mi] for bi in range(len(BR)))
for a in accts:  # 補算後本部殘差同步
    for mi in range(NM):
        a['v'][len(BR)][mi] = a['svc'][mi] - sum(a['v'][bi][mi] for bi in range(len(BR)))

# 刪除全零科目
def nz(a):
    return any(abs(x) > 0.5 for x in a['svc']) or abs(a['b26']) > 0.5 or abs(a['a25']) > 0.5
keep = [a for a in accts if nz(a) or a['code'] in MG]
idmap = {a['id']: n for n, a in enumerate(keep)}
for a in keep:
    p = a['parent']
    while p is not None and p not in idmap:
        p = accts[p]['parent']
    a['parent'] = idmap.get(p) if p is not None else None
for n, a in enumerate(keep): a['id'] = n
accts = keep

# ---------- 3. 勾稽：群組 = 子項加總 ----------
children = {}
for a in accts:
    if a['parent'] is not None: children.setdefault(a['parent'], []).append(a['id'])
by_code = {}
for a in accts: by_code.setdefault(a['code'], a)
checks = list(fills) + [dict(type='欄位排除', item=f'{sh} 工作表「{b}」欄：不是 8 個月份都有的單位欄位，未納入（避免重複計算）', month=0, book=0, detail=0, diff=0) for b, sh in SKIPPED_COLS]
for pid, ch in children.items():
    p = accts[pid]
    if p['code'] in ('120', '500', '520', '90', '510'): continue
    for mi in range(NM):
        s = sum(accts[c]['svc'][mi] for c in ch)
        d = p['svc'][mi] - s
        if abs(d) > 1:
            checks.append(dict(type='科目加總', item=f"{p['code']} {p['name']}", month=mi + 1,
                               book=p['svc'][mi], detail=s, diff=d))

def sv(code, mi): return by_code[code]['svc'][mi]
for mi in range(NM):
    f = [('零服營收 = 零件+工資+外販', '100', sv('4108', mi) + sv('4111', mi) + sv('410804', mi)),
         ('零服毛利 = 營收 − 成本', '120', sv('100', mi) - sv('110', mi)),
         ('營業利益 = 毛利+獎金−外促−內促−固薪−跌價+週邊−費用', '500',
          sv('120', mi) + sv('130', mi) - sv('140', mi) - sv('150', mi) - sv('160', mi)
          - (sv('510210', mi) if '510210' in by_code else 0) + sv('170', mi) - sv('490', mi)),
         ('稅前淨利 = 營業利益 + 業外淨收入', '520', sv('500', mi) + sv('510', mi))]
    for label, c, calc in f:
        d = by_code[c]['svc'][mi] - calc
        checks.append(dict(type='損益公式', item=label, month=mi + 1, book=by_code[c]['svc'][mi], detail=calc, diff=d))

# 2025 月加總 vs 2026 表「25年實績」
if NP == 12:
    for code in ['100', '120', '90', '490', '500']:
        if code in by_code:
            a = by_code[code]; sm = sum(num(ws5.cell(a['xrow'], 7 + m).value) for m in range(12))
            checks.append(dict(type='2025資料', item=f"{code} {a['name']}：2025 月別合計 vs 2026 表「25年實績」", month=0, book=a['a25'], detail=sm, diff=a['a25'] - sm))
# 總表合計欄 vs 月加總
for a in accts:
    tot = num(ws.cell(a['xrow'], 19).value); s = sum(a['svc'])
    if abs(tot - s) > 1:
        checks.append(dict(type='YTD合計', item=f"{a['code']} {a['name']}", month=0, book=tot, detail=s, diff=tot - s))

# ---------- 4. 費用明細（總帳傳票明細，1002 版起）+ 摘要智慧分類 ----------
e = pd.read_excel(SRC, sheet_name='2601-08費用表')
RULES = [
    (r'遞延收益沖', '會員遞延收益沖銷', None),
    (r'費用部門調整', '會計調整', '部門間調整'),
    (r'延長保固會員服務費', '會員／延保', '延長保固會員服務費'),
    (r'五夠美(護照|會員)', '會員／延保', '贈送五夠美會員'),
    (r'感謝影片|影片', '製作物', '影片製作'),
    (r'開學季|加碼贈點', '促銷活動', '季節活動加碼贈點'),
    (r'小小汽車達人', '促銷活動', '親子活動'),
    (r'業代促進', '促銷活動', '業代促進專案'),
    (r'贈\d+元', '優惠券', '現金折抵'),
    (r'^\s*B\d{6,}', '優惠券', '個別工單折抵（工單號）'),
    (r'電訪', '促銷活動', '電訪進廠促進／延保電訪'),
    (r'DC', '客戶招攬DC服務', None),
    (r'圖資更新', '系統服務', '圖資更新'),
    (r'DMS抵用券', '優惠券', 'DMS抵用券找補'),
    (r'YK護照', '會員／延保', 'YK護照抵用券'),
    (r'^各項優惠券$', '優惠券', '各項優惠券(系統)'),
    (r'抵用券', '優惠券', '個別維修／鈑噴抵用券'),
    (r'印製|貼紙', '製作物', '印刷／貼紙'),
    (r'贈品|雨刷|隨行杯|精品|鈑噴活動贈禮', '贈品', None),
    (r'線上登[入錄]', '數位會員', '線上登入禮'),
    (r'UIO', '數位會員', 'UIO 活動'),
    (r'試乘', '促銷活動', '試乘禮'),
    (r'美容', '促銷活動', '美容專案'),
    (r'定保|春檢', '促銷活動', '定保招攬折抵'),
    (r'折', '優惠券', '現金折抵'),
]
PROMO = {'61179910', '61179918'}
def classify(s):
    s = str(s)
    for pat, big, sub in RULES:
        if re.search(pat, s):
            if big == '會員遞延收益沖銷':
                sub = '裕信VIP護照' if 'VIP' in s else ('五夠美／十再美' if '五夠美' in s else '五夠省／十再省')
            if big == '客戶招攬DC服務':
                sub = next((k for k in ['定保招攬', '保養萬元', '撤廠通知', '電話禮儀'] if k in s), 'DC 服務')
            if sub is None:
                sub = '台數活動贈品' if '台數' in s else ('鈑噴活動贈禮' if '鈑噴' in s else '贈品／精品')
            return big, sub
    return '其他', '未分類'
NOT_VENDOR = {'預估', '發票', '合約', '調整', '迴轉', '沖回', '補登', '更正'}
def vendor(s):
    s = str(s).strip()
    m = re.match(r'^([一-鿿A-Za-z0-9（）()]{2,30}?(?:股份有限公司|有限公司|公司|銀行))', s)
    if m: return m.group(1)
    m = re.search(r'-([一-鿿A-Za-z]{2,8})$', s)
    if m and m.group(1) not in NOT_VENDOR: return m.group(1)
    return ''
# 成本中心 → 損益表單位（16 廠依「單位」名稱；其餘依勾稽結果指定）
CC_MAP = {'20S00': '新莊鈑噴', '28S00': '撫遠鈑噴', '00S10': HQ, '21S00': HQ}
def unit_of(cc, u):
    if cc in CC_MAP and CC_MAP[cc] in ENT: return CC_MAP[cc]
    u = str(u) if isinstance(u, str) else ''
    if u.endswith('廠') and u[:-1] in BR: return u[:-1]
    return None
ALIAS = {'714405': '170', '510321': '51030106', '520321': '52030106'}  # 傳票科目 → 損益表科目（名稱不同、內容相同）
STR = {}; SL = []
def si(x):
    x = '' if x is None or (isinstance(x, float) and pd.isna(x)) else str(x)
    if x not in STR: STR[x] = len(SL); SL.append(x)
    return STR[x]
import datetime as _dt
def _day(dt):
    if dt is None or (isinstance(dt, float) and pd.isna(dt)): return 0
    if isinstance(dt, (int, float)): return (_dt.datetime(1899, 12, 30) + _dt.timedelta(days=float(dt))).day
    try: return 0 if pd.isna(dt) else dt.day
    except Exception: return 0
def process_gl(e, year, field, nmonths, tag):
    """e：傳票明細 DataFrame；field：'v'（2026）或 'py'（2025）；回傳精簡列並寫入勾稽"""
    e = e.copy()
    e['code0'] = e['會計科目'].astype(str); e['code'] = e['code0'].map(lambda c: ALIAS.get(c, c)); e['m'] = e['月份'].astype(str).str.replace('月', '').astype(int)
    e['ent'] = [unit_of(c, u) for c, u in zip(e['成本中心'], e['單位'])]
    e['ent'] = e['ent'].where(e['ent'].apply(lambda x: isinstance(x, str)), None)
    e['outu'] = [('' if isinstance(ent, str) else (u if isinstance(u, str) else f'成本中心 {c}')) for ent, c, u in zip(e['ent'], e['成本中心'], e['單位'])]
    sign = {}
    for code, g in e[e['ent'].notna() & (e['ent'] != HQ)].groupby('code'):
        if code not in by_code or not by_code[code].get(field): continue
        a = by_code[code]; glv = g['金額'].sum()
        plv = sum(a[field][ENT.index(u)][m] for u in BR for m in range(nmonths))
        sign[code] = -1 if abs(plv + glv) < abs(plv - glv) else 1
    gl = e[e['ent'].notna()].groupby(['code', 'ent', 'm'])['金額'].sum()
    match = tot = 0; bad = {}
    for code in sorted(set(e['code']) & set(by_code)):
        a = by_code[code]; sg = sign.get(code, 1)
        if not a.get(field): continue
        for u in BR:
            bi = ENT.index(u)
            for mi in range(nmonths):
                pv = a[field][bi][mi]; gv = sg * gl.get((code, u, mi + 1), 0)
                if abs(pv) < 0.5 and abs(gv) < 0.5: continue
                tot += 1
                if abs(pv - gv) <= 1: match += 1
                else: bad.setdefault((code, u), []).append((mi + 1, pv, gv))
    checks.append(dict(type=f'費用明細{tag}', item=f'{year}｜各單位×科目×月：損益表 vs 傳票明細 吻合 {match}/{tot}（{match/max(tot,1):.1%}）', month=0, book=tot, detail=match, diff=tot - match))
    for (code, u), lst in sorted(bad.items(), key=lambda kv: -sum(abs(p - g) for _, p, g in kv[1])):
        a = by_code[code]
        checks.append(dict(type=f'費用明細差異{tag}', item=f"{year}｜{code} {a['name']}｜{u}（{'、'.join(str(m) + '月' for m, _, _ in lst)}不符）", month=(lst[0][0] if len(lst) == 1 else 0),
                           book=sum(p for _, p, _ in lst), detail=sum(g for _, _, g in lst), diff=sum(p - g for _, p, g in lst)))
    codes = set(e['code'])
    for a in accts:
        if a['id'] in children or a['code'] in codes or not a.get(field): continue
        t = sum(sum(r[:nmonths]) for r in a[field])
        if abs(t) > 1: checks.append(dict(type=f'明細缺科目{tag}', item=f"{year}｜{a['code']} {a['name']}（損益表有金額，傳票明細無此科目）", month=0, book=t, detail=0, diff=t))
    for (code, nm), g in e[e['ent'].notna()].groupby(['code', '會計科目中文']):
        if code not in by_code and abs(g['金額'].sum()) > 1:
            checks.append(dict(type=f'損益表無此科目{tag}', item=f"{year}｜{code} {nm}（傳票明細有金額，損益表無此科目）", month=0, book=0, detail=g['金額'].sum(), diff=-g['金額'].sum()))
    for u, g in e[e['ent'].isna()].groupby('outu'):
        checks.append(dict(type=f'損益表範圍外{tag}', item=f"{year}｜{u}（{len(g)} 筆，不在服務部損益表各單位欄位中，未納入分析）", month=0, book=0, detail=g['金額'].sum(), diff=0))
    dc = int((e['借方金額'] - e['貸方金額']).abs().sub(e['金額'].abs()).abs().sum())
    checks.append(dict(type=f'借貸{tag}', item=f'{year}｜|借方 − 貸方| = |金額|（{len(e):,} 筆）', month=0, book=int(e['金額'].abs().sum()), detail=int((e['借方金額'] - e['貸方金額']).abs().sum()), diff=dc))
    rows = []
    for i, r in enumerate(e.itertuples(index=False)):
        code = r.code; s = r.摘要
        c1, c2 = classify(s) if code in PROMO else ('', '')
        ent = ENT.index(r.ent) if isinstance(r.ent, str) else -1
        rows.append([i + 2, si(code), si(str(r.會計科目中文) + ('' if r.code0 == code else f'（傳票科目 {r.code0}）')), ent, si(r.outu), int(r.m), _day(r.傳票日期),
                     int(r.借方金額), int(r.貸方金額), int(sign.get(code, 1) * r.金額), si(s), si(vendor(s)), si(c1), si(c2), int(r.傳票號碼), si(r.成本中心), year])
    return rows
exp = process_gl(e, 2026, 'v', NM, '')
if '2025費用表' in wb.sheetnames and NP:
    e25 = pd.read_excel(SRC, sheet_name='2025費用表')
    exp += process_gl(e25, 2025, 'py', NP, '(2025)')
EXP_FIELDS = ['i', 'code', 'acn', 'e', 'ou', 'm', 'd', 'dr', 'cr', 'a', 's', 'v', 'c1', 'c2', 'vno', 'cc', 'y']

# ---------- 5. 輸出 ----------
def r0(x): return round(x)
out = dict(
    meta=dict(src=SRC_NAME, months=NM, pmonths=NP, year=2026, ent=ENT, br=BR, hq=HQ, type=TYPE,
              note='去年同期＝2025 月別×單位×科目實際數；費用明細含 2025、2026 總帳傳票。'),
    acc=[dict(id=a['id'], c=a['code'], n=a['name'], l=a['lvl'], p=a['parent'], x=a['xrow'],
              v=[[r0(x) for x in row] for row in a['v']], py=([[r0(x) for x in row] for row in a['py']] if a.get('py') else None),
              b26=r0(a['b26']), a25=r0(a['a25']), a24=r0(a['a24'])) for a in accts],
    kpi={k: dict(svc=v['svc'], br=v['br'], b26=v['b26'], a25=v['a25'], py=v['py']) for k, v in kpis.items()},
    exp=dict(f=EXP_FIELDS, s=SL, rows=exp, alias=ALIAS), checks=checks)
json.dump(out, open(os.path.join(HERE, 'data.json'), 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'), default=lambda o: o.item() if hasattr(o, 'item') else str(o))

# Excel 資料模型
pl = []
for a in accts:
    path = []; p = a['parent']
    while p is not None: path.insert(0, accts[p]['name']); p = accts[p]['parent']
    for bi, b in enumerate(ENT):
        for mi in range(NM):
            if abs(a['v'][bi][mi]) > 0.5:
                pl.append(dict(年度=2026, 月份=mi + 1, 廠別=b, 科目代碼=a['code'], 層級=a['lvl'],
                               大科目=path[0] if len(path) > 0 else a['name'], 中科目=path[1] if len(path) > 1 else '',
                               小科目=path[2] if len(path) > 2 else '', 原始科目名稱=a['name'],
                               是否明細=('否' if a['id'] in children else '是'), 金額=a['v'][bi][mi],
                               資料來源=f"{MONTH_SHEETS[mi]}!" + ('總表差額' if b == HQ else b)))
with pd.ExcelWriter(os.path.join(HERE, '資料模型_服務部損益.xlsx')) as w:
    pd.DataFrame([dict(資料表=n, 說明=d) for n, d in [
        ('Fact_PL', '損益事實表：年度×月份×廠別×科目（Long Format）。服務部本部 = 服務部總表 − 各單位加總'),
        ('Fact_Expense', '費用明細表：2025 全年＋2026 1–8月總帳傳票明細（全部會計科目）＋ 單位對應 + 廠商 + 廣促 AI 摘要分類'),
        ('Dim_Account', '會計科目主檔：代碼、層級、上層科目、26預算、25/24實績（年度，服務部層級）'),
        ('Dim_Branch', '廠別主檔'), ('Dim_Date', '日期表'), ('勾稽', '財務數字勾稽結果'),
        ('限制', '2025 無月別/廠別資料；2609–2612 工作表為空白')]]).to_excel(w, sheet_name='資料字典', index=False)
    pd.DataFrame(pl).to_excel(w, sheet_name='Fact_PL', index=False)
    pd.DataFrame([dict(年度=r[16], 月份=r[5], 傳票號碼=r[14], 傳票日=r[6], 會計科目=SL[r[1]], 科目名稱=SL[r[2]], 成本中心=SL[r[15]],
                       損益表單位=(ENT[r[3]] if r[3] >= 0 else ''), 範圍外單位=SL[r[4]], 借方=r[7], 貸方=r[8], 淨額_損益方向=r[9],
                       原始摘要=SL[r[10]], 廠商對象=SL[r[11]], AI分類=SL[r[12]], AI子分類=SL[r[13]], 資料來源=f"2601-08費用表!列{r[0]}") for r in exp]).to_excel(w, sheet_name='Fact_Expense', index=False)
    pd.DataFrame([dict(科目代碼=a['code'], 科目名稱=a['name'], 層級=a['lvl'],
                       上層科目=accts[a['parent']]['name'] if a['parent'] is not None else '',
                       YTD實績=sum(a['svc']), 預算26年=a['b26'], 實績25年=a['a25'], 實績24年=a['a24'], 總表列號=a['xrow']) for a in accts]).to_excel(w, sheet_name='Dim_Account', index=False)
    pd.DataFrame([dict(廠別=b, 類型=(TYPE[i] if i < len(BR) else '本部/未分攤')) for i, b in enumerate(ENT)]).to_excel(w, sheet_name='Dim_Branch', index=False)
    pd.DataFrame([dict(年度=2026, 月份=m, 季度=(m - 1) // 3 + 1, 年月=f'2026-{m:02d}', 是否有資料='是' if m <= NM else '否') for m in range(1, 13)]).to_excel(w, sheet_name='Dim_Date', index=False)
    pd.DataFrame(checks).to_excel(w, sheet_name='勾稽', index=False)

bad = [c for c in checks if abs(c['diff']) > 1]
print('accounts', len(accts), 'exp', len(exp), 'checks', len(checks), 'bad', len(bad))
for c in bad[:40]: print(c)
import collections
print('exp strings', len(SL)); print([c for c in checks if c['type'] in ('費用明細', '欄位排除', '損益表範圍外', '借貸')])
