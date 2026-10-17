# «Aloqabank» ATB: milliy va MHXS hisobotlari qiyosi (2.3-§ metodikasi), 2021–2025
import glob, json, openpyxl
R = '/home/user/tarjima/manbalar/bank_hisobotlari/Aloqabank-31.12.2020-31.12.2025'
IF = json.load(open('/home/user/tarjima/manbalar/mhxs_banklar/ALOQA.json'))['yillar']
NAT = json.load(open('aloqa_milliy.json'))
def num(v):
    try: return float(v)
    except Exception: return None
def capform(y):
    f = glob.glob(f'{R}/{y}/crs031C*')[0]
    ws = openpyxl.load_workbook(f, data_only=True, read_only=True).worksheets[0]
    rows = [list(r) for r in ws.iter_rows(max_row=90, max_col=4, values_only=True)]
    out = {}
    for r in rows:
        k = str(r[0]).strip() if r[0] is not None else ''
        v = num(r[2])
        if k in ('I', 'II', 'III', 'IV', 'V', 'VI', '1.7.', '3.1.', '1.6.') and v is not None and k not in out:
            out[k] = v / 1e6
    return out
OUT = {}
for y in range(2021, 2026):
    c = capform(y); f = IF[str(y)]; n = NAT[f'AL{y}']
    ded = c.get('1.7.', 0.0); at1 = c.get('II', 0.0); t2 = c.get('IV', 0.0); np_n = c.get('3.1.', 0.0); rwa = c['VI']
    np_i = f['NP']; eq_i = f['EQ']; nci = f.get('NCI') or 0.0
    cet1_i = eq_i - nci - max(np_i, 0) - at1 - ded          # imtiyozli aksiyalar (AT1) MHXS kapitalidan ayiriladi
    if np_i < 0 and False: pass
    t2_i = t2 - np_n + max(np_i, 0)
    car_i = (cet1_i + at1 + t2_i) / rwa * 100
    s3n = n['S3']; s3i = f['S3'] / f['S_TOTAL'] * 100; s2i = f['S2'] / f['S_TOTAL'] * 100
    OUT[y] = dict(TA_n=n['_TA'] / 1e6, TA_i=f['TA'], EQ_n=n['_EQ'] / 1e6, EQ_i=eq_i, NP_n=n['_NP'] / 1e6, NP_i=np_i,
                  ROE_n=n['ROE'], CET1_n=c['I'] / rwa * 100, CAR_n=c['V'] / rwa * 100, CET1_i=cet1_i / rwa * 100, CAR_i=car_i,
                  NPL_n=s3n, S3_i=s3i, S2_n=n['S2'], S2_i=s2i, KYKR=s3i / s3n, RES_n=None, ECL_i=f['ECL'], L_i=f['LOANS_GROSS'], L_n=n['_L'] / 1e6,
                  RWA=rwa, DED=ded, AT1=at1)
for y, o in OUT.items():
    print(y, {k: round(v, 2) for k, v in o.items() if isinstance(v, float)})
# ROE MHXS
for y in range(2021, 2026):
    p = IF[str(y - 1)]; f = IF[str(y)]
    OUT[y]['ROE_i'] = f['NP'] / ((f['EQ'] + p['EQ']) / 2) * 100
    OUT[y]['ROA_i'] = f['NP'] / ((f['TA'] + p['TA']) / 2) * 100
print({y: (round(OUT[y]['ROE_n'], 1), round(OUT[y]['ROE_i'], 1)) for y in OUT})
print('5 yillik foyda: milliy', round(sum(o['NP_n'] for o in OUT.values()), 1), 'MHXS', round(sum(o['NP_i'] for o in OUT.values()), 1))
json.dump(OUT, open('aloqa23.json', 'w'), indent=1)
