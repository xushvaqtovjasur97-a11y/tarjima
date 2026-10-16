# 3.2-§: sinf chegaralarini MBII-13 bo'yicha qo'shimcha tekshirish (5 bank, 2021–2024, hodisa keyingi yilda)
import pandas as pd, json, numpy as np
import mbii as M
p = pd.read_csv('panel33.csv').sort_values(['bank', 't']); g = p.groupby('bank')
p['lossq'] = -p.NPq / g.Kapital.shift(1) * 100; p['EV'] = (p.lossq >= 5).astype(int)
ev = p.groupby(['bank', 'y']).EV.max()
V = json.load(open('vazn32.json'))['bank_mbii']['Ekspert']
nm = {'SQB': 'O‘zsanoatqurilishbank', 'MKB': 'Mikrokreditbank', 'BRB': 'Biznesni rivojlantirish banki', 'XALQ': 'Xalq banki', 'IPOTEKA': 'Ipoteka-bank'}
rows = []
for k, v in V.items():
    b, y = k.split('|'); y = int(y)
    if y <= 2024:
        rows.append((b, y, v, M.cls(v), int(ev.get((nm[b], y + 1), 0))))
df = pd.DataFrame(rows, columns=['bank', 'y', 'mbii13', 'sinf', 'ev_next'])
out = df.groupby('sinf').ev_next.agg(['count', 'sum']).to_dict('index')
print(df.to_string()); print(out)
json.dump({'kuzatuvlar': df.to_dict('records'), 'sinflar': out}, open('sinf_mbii13.json', 'w'), ensure_ascii=False, indent=1, default=int)
