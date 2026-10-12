# «Aloqabank» ATB: prudensial shakllardan (crs031C/L/O) MBII-milliy ko'rsatkichlarini hisoblash – panel31.py bilan bir xil usul
import glob, json, openpyxl, re, sys
R = '/home/user/tarjima/manbalar/bank_hisobotlari/Aloqabank-31.12.2020-31.12.2025'
WANT = ('balance', 'баланс', 'economic', 'иқт', 'asset classification', 'loan classification', 'bank capital calc', 'liquidity analysis', 'income',
        'calculation of risk weighted as')
def sheets(path):
    out = {}
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    for ws in wb.worksheets:
        if any(w in ws.title.strip().lower() for w in WANT):
            out[ws.title.strip()] = [list(r) for r in ws.iter_rows(max_row=160, max_col=26, values_only=True)]
    return out
P = {}
for y in range(2020, 2026):
    m = {}
    for f in sorted(glob.glob(f'{R}/{y}/crs031*')):
        m.update(sheets(f))
    P[('AL', y)] = m
src = open('panel31.py').read()
ns = {}
exec(src.split('OUT={}')[0].split("P=pickle.load")[0] + src.split('O=pickle.load(open(SP+\'/mk1619/d.pkl\',\'rb\'))')[1].split('OUT={}')[0], ns)
ns['P'] = P; ns['O'] = {}
get = ns['get']
OUT = {}; prev = None
for y in range(2020, 2026):
    r = get('AL', y)
    if prev and r['_NP'] is not None:
        r['ROA'] = r['_NP'] / ((r['_TA'] + prev['_TA']) / 2) * 100
        r['ROE'] = r['_NP'] / ((r['_EQ'] + prev['_EQ']) / 2) * 100
    else:
        r['ROA'] = r['_NP'] / r['_TA'] * 100 if r['_NP'] else None; r['ROE'] = r['_NP'] / r['_EQ'] * 100 if r['_NP'] else None
    OUT[f'AL{y}'] = r; prev = r
    print(y, {k: (round(v, 2) if isinstance(v, float) and abs(v) < 1e5 else (round(v / 1e6, 1) if isinstance(v, (int, float)) and v else v)) for k, v in r.items()})
json.dump(OUT, open('aloqa_milliy.json', 'w'), indent=1)
