# 4.2-§: bank darajasidagi prognoz – ikki usul:
#  P (proporsional, dissertatsiyadagi asosiy usul): tizim modeli og‘ishi bankka 2025-yil IV choragidagi ulushiga proporsional taqsimlanadi;
#  B (bank paneli, 2-model): S3_it = a_i + 0,684·S3_i,t−1 − 0,153·YaIM_t – har bir bank uchun bevosita, og‘ish barcha banklarga bir xil (foiz punkt).
# Har ikkala usulda ROA – 3-model, boshqa ko‘rsatkichlar 2025-yil IV choragi darajasida, MBII-8 va 4.1-§ zonalari.
import json, numpy as np, pandas as pd
import mbii as M

O = json.load(open('fc42.json'))
p = pd.read_csv('panel33.csv'); d = p[p.tanlov].copy().sort_values(['bank', 't'])


def w(x, a=.01):
    lo, hi = x.quantile([a, 1 - a]); return x.clip(lo, hi)


for c in ['S3', 'S3_l1', 'ROAq']:
    d[c] = w(d[c])
D = pd.read_csv('tizim42.csv', index_col=0)
S25, G25 = D.NPL.loc['2025_IV'], D.YaIM.loc['2025_IV']
SYS = O['tizim_NPL_yo‘li']                                   # 12 chorak: 2026-I … 2028-IV
SC = list(SYS)
GDP = {s: [O['ssenariy'][s][str(y)] for y in (2026, 2027, 2028) for _ in range(4)] for s in SC}
RHO_B, BG_B = O['panel_NPL']['b']['S3_l1'], O['panel_NPL']['b']['YaIM_osish']
BR = O['panel_ROA']['b']['S3']
W8 = {b: v for b, v in M.W['taklif'].items() if b != 'V'}


def mb(v):
    v = {k: (None if (x is None or x != x) else x) for k, x in v.items()}
    if sum(x is not None for x in v.values()) < 6: return np.nan
    return M.mbii(v, W8)[0]


def zona(m):
    return 'yashil' if m >= 3.25 else 'sariq' if m >= 2.75 else 'qizil'


def proj(r, s3):
    s3 = max(s3, 0.0)
    roa = r.ROA + BR * (s3 - r.S3); roe = roa * r.Aktiv / r.Kapital
    cov = min(r.RES / r.NPL, 1.0) if (r.RES == r.RES and r.NPL > 0) else 0.5
    npln = r.NPLnet + (s3 - r.S3) * r.Kredit / r.Kapital * (1 - cov) if r.NPLnet == r.NPLnet else np.nan
    return s3, roa, mb({'LEV': r.LEV, 'S3': s3, 'NPLnet': npln, 'HLA': r.HLA_r, 'LTD': r.LTD, 'ROA': roa, 'ROE': roe, 'CI': r.CI})


def bank_dev(s):
    dev, out = 0.0, []
    for g in GDP[s]:
        dev = RHO_B * dev + BG_B * (g - G25); out.append(dev)
    return out


b25 = d[d.davr == '2025_IV'].set_index('bank')
BD = {s: bank_dev(s) for s in SC}
rows = []
for s in SC:
    for q in range(12):
        for b, r in b25.iterrows():
            for usul, s3n in (('P', r.S3 * SYS[s][q] / S25), ('B', r.S3 + BD[s][q])):
                s3, roa, m = proj(r, s3n)
                rows.append(dict(usul=usul, ssenariy=s, q=q, yil=2026 + q // 4, bank=b, Aktiv=r.Aktiv, DAVLAT=r.DAVLAT, S3=s3, ROA=roa, MBII=m))
F = pd.DataFrame(rows); F['zona'] = F.MBII.apply(zona)


def agg(x):
    wv = lambda c: np.average(x[c], weights=x.Aktiv)
    return pd.Series({'S3_w': wv('S3'), 'ROA_w': wv('ROA'), 'qizil': int((x.zona == 'qizil').sum()), 'sariq': int((x.zona == 'sariq').sum()),
                      'yashil': int((x.zona == 'yashil').sum()), 'qizil_aktiv': x.loc[x.zona == 'qizil', 'Aktiv'].sum() / x.Aktiv.sum() * 100,
                      'qizil_banklar': ', '.join(sorted(x.loc[x.zona == 'qizil', 'bank']))})


YE = F[F.q % 4 == 3]                                          # yil oxiri
T = YE.groupby(['usul', 'ssenariy', 'yil']).apply(agg)
PEAK = F.groupby(['usul', 'ssenariy', 'bank']).S3.max()
BANKS = ['O‘zsanoatqurilishbank', 'Mikrokreditbank', 'Aloqa bank']
HIST = {b: d[d.bank == b].sort_values('t')[['davr', 'S3']].values.tolist() for b in BANKS}
PATH = {(u, s, b): F[(F.usul == u) & (F.ssenariy == s) & (F.bank == b)].sort_values('q').S3.round(4).tolist() for u in 'PB' for s in SC for b in BANKS}

if __name__ == '__main__':
    pd.set_option('display.width', 250)
    print(T.drop(columns='qizil_banklar').round(2).to_string())
    print(T.qizil_banklar.to_string())
    for b in BANKS:
        print(b, {(u, s): round(PEAK[(u, s, b)], 2) for u in 'PB' for s in SC})
    print('bank dev keskin', [round(x, 2) for x in BD['Keskin salbiy']])
    print('tizim keskin', SYS['Keskin salbiy'])
