# 3.1-§: MBII 2025-yil, uch bank (SQB, MKB, Aloqabank), milliy va MHXS variantlari – yagona ta'riflar
import json, sys
sys.path.insert(0, '/home/user/tarjima/manbalar/hisob_kitob')
import mbii as M, mbii32 as I32, mbii_milliy as MM
R = '/home/user/tarjima/manbalar/hisob_kitob/'
alm = json.load(open(R + 'aloqa_milliy.json'))['AL2025']
al = json.load(open(R + 'aloqa_mbii.json'))['2025']
X = {}
X['SQB milliy'] = MM.ind('SQ', 2025)
X['MKB milliy'] = MM.ind('MK', 2025)
a = dict(al['ind_milliy']); a['LEV'] = alm['_EQ'] / alm['_TA'] * 100   # boshqa banklar kabi xususiy kapital / aktivlar
X['ALQ milliy'] = a
X['SQB MHXS'] = I32.ind('SQB')[2025]
X['MKB MHXS'] = I32.ind('MKB')[2025]
x = I32.ind('ALOQA')[2025]; x['CET1'], x['CAR'] = 9.76, 13.11            # 2.3.5-jadval: MB metodikasi bo'yicha qayta hisoblangan
X['ALQ MHXS'] = x
OUT = {}
for k, v in X.items():
    m, bs = M.mbii(v, M.W['taklif'])
    sc = {n: M.score(v[n], dr, an) for _, _, n, dr, an in M.IND if v.get(n) is not None}
    OUT[k] = dict(ind=v, ball=sc, bloklar=bs, mbii=m, sinf=M.cls(m))
    print(k, round(m, 3), M.cls(m), {b: round(s, 2) for b, s in bs.items()})
json.dump(OUT, open(R + 'mbii31_3bank.json', 'w'), indent=1, ensure_ascii=False, default=float)
