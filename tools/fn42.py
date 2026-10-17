import re, json
SRC='/home/user/tarjima/paragraflar/4.2-paragraf.src.md'
DST='/home/user/tarjima/paragraflar/4.2-paragraf.md'
R={}
for f in ['fn11.py','fn13x.py','fn22.py','fn12n.py','fn23.py','fn31.py','fn32.py','fn33.py','fn41.py']:
    code=open('/home/user/tarjima/tools/'+f).read(); code=code[:code.find('\ndef ref')]
    if f not in ('fn11.py','fn13x.py','fn22.py'): code=code[code.find('\nR.update({'):]
    ns={'R':{}}; exec(code,ns); R.update(ns['R'])
R.update({
'CIH07':('Čihák M. Introduction to Applied Stress Testing: IMF Working Paper WP/07/59. – Washington, DC: International Monetary Fund, 2007.',''),
'SV':('Sorge M., Virolainen K. A comparative analysis of macro stress-testing methodologies with application to Finland // Journal of Financial Stability. – 2006. – Vol. 2, No. 2. – P. 113–151.',''),
'BCBSST':('Basel Committee on Banking Supervision. Stress testing principles. – Basel: BIS, October 2018.',''),
'BUL20':('O‘zbekiston Respublikasi Markaziy banki. Statistik byulleten: 2021-yil. – Toshkent, 2022. – 5.1.8a, 5.1.8b-jadvallar; O‘zbekiston Respublikasi Markaziy banki. Statistik byulleten: 2026-yil mart. – Toshkent, 2026. – 1.1, 5.1-jadvallar. – https://cbu.uz/uz/statistics/buleten/',''),
})
def ref(k,p):
  t,pp=R[k]
  return t if p=='-' else f'{t} – {pp} {p}.'
def J(*xs): return '; '.join(x.rstrip('.') for x in xs)+'.'
SRCS='Tadqiqot natijasida muallif tomonidan quyidagi manbalar asosida ishlab chiqildi: '
CALC='Tadqiqot natijasida muallif tomonidan quyidagi manbalar ma’lumotlari asosida hisoblandi: '
DATA=(ref('BUL20','-'),ref('CBUBS','-'),ref('SNS','-'))
SAME='Tadqiqot natijasida muallif tomonidan 4.2.2-jadvalda ko‘rsatilgan manbalar ma’lumotlari asosida hisoblandi.'
A={'T1':SRCS+J(ref('BUL20','-')),'T2':CALC+J(*DATA),'T3':SAME,'F1':SAME,'T4':SAME,'F2':SAME,
'T5':SAME,'T6':CALC+J(*DATA,ref('SQBIFRS','2–3, 28–31'),ref('MKIFRS25','11–12')),'F3':'Tadqiqot natijasida muallif tomonidan 4.2.6-jadval ma’lumotlari asosida ishlab chiqildi.'}
O=json.load(open('/home/user/tarjima/manbalar/hisob_kitob/fc42.json'))
def n(x,d): return f'{x:.{d}f}'.replace('.',',').replace('-','−')
def st(p): return '\\*\\*\\*' if p<0.01 else '\\*\\*' if p<0.05 else '\\*' if p<0.1 else ''
def c(m,v,d=3): return f"{n(m['b'][v],d)}{st(m['p'][v])} ({n(m['se'][v],d)})"
def T2():
  s,pn,pr=O['tizim_model'],O['panel_NPL'],O['panel_ROA']
  L=['| O‘zgaruvchi | 1-model: tizim, muammoli kreditlar ulushi | 2-model: bank paneli, muammoli kreditlar ulushi | 3-model: bank paneli, ROA |','|---|---|---|---|']
  L.append(f"| Muammoli kreditlar ulushi (t−1) | {c(s,'NPL_l1')} | {c(pn,'S3_l1')} | – |")
  L.append(f"| YaIM o‘sishi (t−4) | {c(s,'G')} | – | – |")
  L.append(f"| YaIM o‘sishi (t) | – | {c(pn,'YaIM_osish')} | – |")
  L.append(f"| Muammoli kreditlar ulushi (t) | – | – | {c(pr,'S3')} |")
  L.append(f"| Konstanta | {c(s,'const')} | {c(pn,'c')} | {c(pr,'c')} |")
  L.append(f"| Bank effektlari | – | ha | ha |")
  L.append(f"| Davr | 2021-y. I chorak – 2025-y. IV chorak (YaIM laglari 2020-yildan) | 2021-y. IV chorak – 2025-y. IV chorak | 2021-y. IV chorak – 2025-y. IV chorak |")
  L.append(f"| Kuzatuvlar soni | {s['N']} | {pn['N']} | {pr['N']} |")
  L.append(f"| R² | {n(s['R2'],2)} | {n(pn['R2w'],2)} (bank ichidagi) | {n(pr['R2w'],2)} (bank ichidagi) |")
  L.append(f"| YaIM o‘sishining 1 foiz punktga barqaror pasayishining uzoq muddatli ta’siri, foiz punkt | {n(-s['LR'],2)} | {n(-pn['LR'],2)} | – |")
  return '\n'.join(L)
def T3():
  L=['| Prognoz qilingan yil | Haqiqiy qiymat (I–IV chorak), % | Taklif etilgan usul bo‘yicha prognoz, % | O‘rtacha xato: taklif etilgan usul | O‘rtacha xato: o‘zgarishsiz holat | O‘rtacha xato: o‘rtachaga qaytuvchi model |','|---|---|---|---|---|---|']
  for b in O['backtest_tizim']:
    y=int(b['davr'])+1
    L.append(f"| {y} | {'; '.join(n(x,2) for x in b['haqiqiy'])} | {'; '.join(n(x,2) for x in b['bizning'])} | {n(b['MAE_bizning'],2)} | {n(b['MAE_naiv'],2)} | {n(b['MAE_qaytuvchi'],2)} |")
  return '\n'.join(L)
def T4():
  T=O['tizim']; cols=[('2025','2025')]+[(s,y) for s in ('Asosiy','Ijobiy','Salbiy','Keskin salbiy') for y in ('2027','2028')]
  hd=['2025 (haqiqiy)']+[f'{s} {y}' for s,y in cols[1:]]
  L=['| Ko‘rsatkich | '+' | '.join(hd)+' |','|---|'+'---|'*len(hd)]
  def row(t,k,d,suf=''):
    return f'| {t} | '+' | '.join(n(T[f'{s}|{y}'][k],d)+suf for s,y in cols)+' |'
  L.append(row('Muammoli kreditlar ulushi, %','S3_w',2)); L.append(row('ROA, %','ROA_w',2)); L.append(row('MBII medianasi, ball','MBII_med',2))
  L.append('| Banklar soni: yashil / sariq / qizil zona | '+' | '.join(f"{int(T[f'{s}|{y}']['yashil'])} / {int(T[f'{s}|{y}']['sariq'])} / {int(T[f'{s}|{y}']['qizil'])}" for s,y in cols)+' |')
  L.append(row('Qizil zonadagi banklar aktivlari ulushi, %','qizil_aktiv',1))
  return '\n'.join(L)
def T5g():
  G=O['guruh']; cols=[('2025','2025')]+[(s,'2027') for s in ('Asosiy','Ijobiy','Salbiy','Keskin salbiy')]
  L=['| Ko‘rsatkich | Guruh | 2025 (haqiqiy) | Asosiy | Ijobiy | Salbiy | Keskin salbiy |','|---|---|---|---|---|---|---|']
  for k,t_,dd in (('S3_w','Muammoli kreditlar ulushi, %',2),('ROA_w','ROA, %',2),('MBII_med','MBII medianasi, ball',2),('qizil','Qizil zonadagi banklar soni',0)):
    for g,gn in (('1','Davlat banklari'),('0','Boshqa banklar')):
      L.append(f'| {t_} | {gn} | '+' | '.join(n(G[f"{s}|{y}|{g}"][k],dd) for s,y in cols)+' |')
  return '\n'.join(L)
def T5():
  B=O['banklar']; cols=[('2025',None)]+[(s,'2027') for s in ('Asosiy','Ijobiy','Salbiy','Keskin salbiy')]
  L=['| Bank va hisobot | 2025 (haqiqiy) | Asosiy 2027 | Ijobiy 2027 | Salbiy 2027 | Keskin salbiy 2027 |','|---|---|---|---|---|---|']
  def z(m): return 'yashil' if m>=3.25 else 'sariq' if m>=2.75 else 'qizil'
  for code,nm in (('SQB','«O‘zsanoatqurilishbank» ATB'),('MKB','«Mikrokreditbank» ATB')):
    for var,vn,s3 in (('milliy','milliy hisobot','S3_milliy'),('MHXS','MHXS hisoboti','S3_MHXS')):
      cells=[]
      for s,y in cols:
        e=B[code]['2025'] if s=='2025' else B[code][f'{s}|{y}']
        cells.append(f"{n(e[var],2)} ({z(e[var])}; {n(e[s3],1)} %)")
      L.append(f'| {nm}, {vn} | '+' | '.join(cells)+' |')
  return '\n'.join(L)
txt=open(SRC).read()
for k,f in {'T2':T2,'T3':T3,'T4':T4,'T5':T5g,'T6':T5}.items(): txt=txt.replace('{{'+k+'}}',f())
notes=[]
def sub(m):
  k,v=m.group(1),m.group(2)
  notes.append(A[v] if k=='AUTHOR' else ref(k,v))
  return f'[^{len(notes)}]'
txt=re.sub(r'%([A-Z][A-Z0-9]*):([^%]+)%',sub,txt)
txt=txt.rstrip()+'\n\n'+'\n'.join(f'[^{i+1}]: {x}' for i,x in enumerate(notes))+'\n'
open(DST,'w').write(txt); print(len(notes),'notes')
