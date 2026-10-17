# 3.1-§: MBII 2021–2025-yillar, uch bank (SQB, MKB, Aloqabank), milliy va MHXS variantlari – mbii31_3bank.py bilan yagona ta'riflar
import json, sys
sys.path.insert(0, '/home/user/tarjima/manbalar/hisob_kitob')
import mbii as M, mbii32 as I32, mbii_milliy as MM
R = '/home/user/tarjima/manbalar/hisob_kitob/'
alm = json.load(open(R + 'aloqa_milliy.json'))
al = json.load(open(R + 'aloqa_mbii.json'))
ALQ_CAP = dict(CET1=[12.30, 9.63, 11.53, 11.65, 9.76], CAR=[14.37, 12.54, 15.07, 14.02, 13.11])  # 2.3.5-jadval
OUT = {}
for y in range(2021, 2026):
    X = {}
    X['SQB milliy'] = MM.ind('SQ', y)
    X['MKB milliy'] = MM.ind('MK', y)
    a = dict(al[str(y)]['ind_milliy']); a['LEV'] = alm[f'AL{y}']['_EQ'] / alm[f'AL{y}']['_TA'] * 100
    X['ALQ milliy'] = a
    X['SQB MHXS'] = I32.ind('SQB')[y]
    X['MKB MHXS'] = I32.ind('MKB')[y]
    x = dict(I32.ind('ALOQA')[y]); i = y - 2021; x['CET1'], x['CAR'] = ALQ_CAP['CET1'][i], ALQ_CAP['CAR'][i]
    X['ALQ MHXS'] = x
    for k, v in X.items():
        m, bs = M.mbii(v, M.W['taklif'])
        OUT.setdefault(k, {})[y] = dict(mbii=m, sinf=M.cls(m), bloklar=bs)
for k, v in OUT.items():
    print(k, [round(v[y]['mbii'], 3) for y in range(2021, 2026)], [v[y]['sinf'] for y in range(2021, 2026)])
json.dump(OUT, open(R + 'mbii31_dyn3.json', 'w'), indent=1, ensure_ascii=False, default=float)
