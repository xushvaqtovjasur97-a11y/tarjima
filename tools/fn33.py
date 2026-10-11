import re, json
SRC='/home/user/tarjima/paragraflar/3.3-paragraf.src.md'
DST='/home/user/tarjima/paragraflar/3.3-paragraf.md'
R={}
for f in ['fn11.py','fn13x.py','fn22.py','fn12n.py','fn23.py','fn31.py','fn32.py']:
    code=open('/home/user/tarjima/tools/'+f).read(); code=code[:code.find('\ndef ref')]
    if f in ('fn12n.py','fn23.py','fn31.py','fn32.py'): code=code[code.find('\nR.update({'):]
    ns={'R':{}}; exec(code,ns); R.update(ns['R'])
R.update({
'CBUBS':('O‘zbekiston Respublikasi Markaziy banki. Banklar statistikasi: tijorat banklarining asosiy ko‘rsatkichlari va muammoli kreditlari to‘g‘risidagi oylik ma’lumotlar, 2021-yil aprel – 2026-yil yanvar. – https://cbu.uz/uz/statistics/bankstats/',''),
'BULM':('O‘zbekiston Respublikasi Markaziy banki. Statistik byulleten: 2026-yil mart. – Toshkent, 2026. – 1.1, 1.2, 3.1, 3.7, 5.1-jadvallar. – https://cbu.uz/uz/statistics/buleten/',''),
'ATH':('Athanasoglou P.P., Brissimis S.N., Delis M.D. Bank-specific, industry-specific and macroeconomic determinants of bank profitability // Journal of International Financial Markets, Institutions and Money. – 2008. – Vol. 18, No. 2. – P. 121–136.',''),
'INS':('Iannotta G., Nocera G., Sironi A. Ownership structure, risk and performance in the European banking industry // Journal of Banking & Finance. – 2007. – Vol. 31, No. 7. – P. 2127–2149.',''),
'CORN':('Cornett M.M., Guo L., Khaksari S., Tehranian H. The impact of state ownership on performance differences in privately-owned versus state-owned banks: An international comparison // Journal of Financial Intermediation. – 2010. – Vol. 19, No. 1. – P. 74–94.',''),
'FNW':('Foos D., Norden L., Weber M. Loan growth and riskiness of banks // Journal of Banking & Finance. – 2010. – Vol. 34, No. 12. – P. 2929–2940.',''),
'LVM':('Louzis D.P., Vouldis A.T., Metaxas V.L. Macroeconomic and bank-specific determinants of non-performing loans in Greece: A comparative study of mortgage, business and consumer loan portfolios // Journal of Banking & Finance. – 2012. – Vol. 36, No. 4. – P. 1012–1027.',''),
'BJP':('Beck R., Jakubík P., Piloiu A. Key Determinants of Non-performing Loans: New Evidence from a Global Sample // Open Economies Review. – 2015. – Vol. 26, No. 3. – P. 525–550.',''),
'NK':('Nkusu M. Nonperforming Loans and Macrofinancial Vulnerabilities in Advanced Economies: IMF Working Paper WP/11/161. – Washington, DC: International Monetary Fund, 2011.',''),
'DKH':('Demirgüç-Kunt A., Huizinga H. Bank activity and funding strategies: The impact on risk and returns // Journal of Financial Economics. – 2010. – Vol. 98, No. 3. – P. 626–650.',''),
'LL':('Laeven L., Levine R. Bank governance, regulation and risk taking // Journal of Financial Economics. – 2009. – Vol. 93, No. 2. – P. 259–275.',''),
'LS':('Lepetit L., Strobel F. Bank insolvency risk and time-varying Z-score measures // Journal of International Financial Markets, Institutions and Money. – 2013. – Vol. 25. – P. 73–87.',''),
'HAUS':('Hausman J.A. Specification Tests in Econometrics // Econometrica. – 1978. – Vol. 46, No. 6. – P. 1251–1271.',''),
'BP':('Breusch T.S., Pagan A.R. The Lagrange Multiplier Test and its Applications to Model Specification in Econometrics // Review of Economic Studies. – 1980. – Vol. 47, No. 1. – P. 239–253.',''),
'AHS':('Anderson T.W., Hsiao C. Estimation of Dynamic Models with Error Components // Journal of the American Statistical Association. – 1981. – Vol. 76, No. 375. – P. 598–606.',''),
'NIC':('Nickell S. Biases in Dynamic Models with Fixed Effects // Econometrica. – 1981. – Vol. 49, No. 6. – P. 1417–1426.',''),
'ABND':('Arellano M., Bond S. Some Tests of Specification for Panel Data: Monte Carlo Evidence and an Application to Employment Equations // Review of Economic Studies. – 1991. – Vol. 58, No. 2. – P. 277–297.',''),
'WOOL':('Wooldridge J.M. Econometric Analysis of Cross Section and Panel Data. – 2nd ed. – Cambridge, MA: MIT Press, 2010.','P.'),
'BALT':('Baltagi B.H. Econometric Analysis of Panel Data. – 6th ed. – Cham: Springer, 2021.','P.'),
'SY':('Stock J.H., Yogo M. Testing for Weak Instruments in Linear IV Regression // Identification and Inference for Econometric Models: Essays in Honor of Thomas Rothenberg / ed. by D.W.K. Andrews, J.H. Stock. – Cambridge: Cambridge University Press, 2005. – P. 80–108.',''),
})
def ref(k,p):
  t,pp=R[k]
  return t if p=='-' else f'{t} – {pp} {p}.'
def J(*xs): return '; '.join(x.rstrip('.') for x in xs)+'.'
CALC='Tadqiqot natijasida muallif tomonidan quyidagi manbalar ma’lumotlari asosida hisoblandi: '
SRCS='Tadqiqot natijasida muallif tomonidan quyidagi manbalar asosida ishlab chiqildi: '
DATA=(ref('CBUBS','-'),ref('SNS','-'),ref('BULM','-'))
SAME='Tadqiqot natijasida muallif tomonidan 3.3.2-jadvalda ko‘rsatilgan manbalar ma’lumotlari asosida hisoblandi.'
A={
'T1':SRCS+J(ref('ATH','-'),ref('INS','-'),ref('CORN','-'),ref('FNW','-'),ref('DKH','-'),ref('LVM','-'),ref('BJP','-')),
'T2':CALC+J(*DATA),
'F1':SAME,'F2':SAME,'F3':SAME,'T3':SAME,'T4':SAME,
'T5':SAME,'T6':SAME,
'T7':'Tadqiqot natijasida muallif tomonidan 3.3.3–3.3.6-jadvallar ma’lumotlari asosida ishlab chiqildi.',
'IPO':'«Ipoteka-bank» ATB aksiyalarining nazorat paketi 2023-yilda OTP Bank guruhiga sotilganligi sababli bank 2023-yil III choragidan boshlab xususiy bank sifatida qaraldi. Tanlanmadagi boshqa davlat banklari: «O‘zmilliybank», «O‘zsanoatqurilishbank», «Agrobank», «Asaka bank», «Xalq banki», «Mikrokreditbank», «Biznesni rivojlantirish banki», «Aloqabank», «Turonbank» ATBlar.',
'DATA':'Markaziy bank ma’lumotlari bilan «SNS» bazasi bank nomlari bo‘yicha birlashtirildi. «SNS» bazasidagi tekshiruvda aniqlangan xato qiymatlar («Mikrokreditbank» ATB ning 2021-yil IV va 2023-yil IV choraklaridagi, «Asaka bank» ATB ning 2022-yil I va 2023-yil I choraklaridagi, «Invest Finance bank» ATB ning 2022-yil IV choragidagi, «Tenge bank» ATB ning 2023-yil II choragidagi nolga teng kredit zaxiralari, «O‘zmilliybank» ATB ning 2022-yil IV choragidagi barcha qiymatlari) tahlildan chiqarildi. Manbalar: '+J(*DATA[:2]),
}
# --- regressiya jadvallari model33.json dan
O=json.load(open('/home/user/tarjima/manbalar/hisob_kitob/model33.json'))
def n(x,d): return f'{x:.{d}f}'.replace('.',',').replace('-','−')
def st(p): return '***' if p<0.01 else '**' if p<0.05 else '*' if p<0.1 else ''
def cell(m,v,d=3):
  if v not in O[m]['b']: return '–'
  b,s,p=O[m]['b'][v],O[m]['se'][v],O[m]['p'][v]
  return f'{n(b,d)}{st(p)} ({n(s,d)})'
NM={'lnTA':'ln(aktivlar)','KR_osish_l4':'Kreditlar o‘sishi (t−4), %','DEP_A':'Depozitlar / aktivlar, %','NIM':'Sof foiz marjasi, %','DAVLAT':'Davlat banki (dummy)',
    'S3_l1':'Muammoli kreditlar ulushi (t−1), %','KR_osish_l8':'Kreditlar o‘sishi (t−8), %','YaIM_osish':'YaIM o‘sishi, %','Asosiy_stavka':'Asosiy stavka, %','USD_osish':'AQSh dollari kursi o‘zgarishi, %','Inflyatsiya':'Inflyatsiya, %'}
def T3():
  ms=['MBII Pooled','MBII RE','MBII TWFE','MBII Anderson-Hsiao']
  L=['| O‘zgaruvchi | Pooled OLS | RE | FE (bank va chorak) | Dinamik (Anderson–Hsiao) |','|---|---|---|---|---|']
  L.append('| MBII (t−1) | – | – | – | '+cell(ms[3],'D_MBII_l1')+' |')
  for v in ['lnTA','KR_osish_l4','DEP_A','NIM','DAVLAT']:
    L.append(f'| {NM[v]} | '+' | '.join(cell(m,('D_'+v) if 'Anderson' in m else v,4 if v=='KR_osish_l4' else 3) for m in ms)+' |')
  L.append('| Chorak effektlari | ha | ha | ha | ha |')
  L.append('| Kuzatuvlar soni | '+' | '.join(str(O[m]['N']) for m in ms)+' |')
  L.append(f"| R² | {n(O[ms[0]]['R2'],2)} | {n(O[ms[1]]['R2'],2)} (umumiy) | {n(O[ms[2]]['R2'],2)} (bank ichidagi) | – |")
  h=O['MBII Hausman']; bp=O['MBII BP_LM']
  L.append(f"| Testlar | – | Breusch–Pagan LM = {n(bp['LM'],1)} (p < 0,001) | Hausman χ²({h['df']}) = {n(h['H'],2)} (p = {n(h['p'],3)}) | 1-bosqich F = {n(O[ms[3]]['F1'],1)} |")
  return '\n'.join(L)
def T4():
  a,b=O['MBII davlat_umumiy'],O['MBII davlat_funding']
  L=['| Ko‘rsatkich | 1-model: funding nazoratisiz | 2-model: depozitlar va foiz marjasi nazorati bilan |','|---|---|---|']
  for v in ['DAVLAT','lnTA','DEP_A','NIM']:
    L.append(f'| {NM[v]} | '+cell('MBII davlat_umumiy',v)+' | '+cell('MBII davlat_funding',v,4 if v=='DEP_A' else 3)+' |')
  L.append(f"| Kuzatuvlar soni | {a['N']} | {b['N']} |")
  L.append(f"| R² | {n(a['R2'],2)} | {n(b['R2'],2)} |")
  da,db=a['b']['DAVLAT'],b['b']['DAVLAT']
  L.append(f"| Davlat banki ta’sirining funding modeli orqali izohlanadigan qismi | – | {n(da-db,2)} ball ({n((da-db)/da*100,0)} %) |")
  return '\n'.join(L)
def T5():
  ms=['lnZ Pooled','lnZ RE','lnZ TWFE']
  L=['| O‘zgaruvchi | Pooled OLS | RE | FE (bank va chorak) |','|---|---|---|---|']
  for v in ['lnTA','KR_osish_l4','DEP_A','NIM','DAVLAT']:
    L.append(f'| {NM[v]} | '+' | '.join(cell(m,v,4 if v=='KR_osish_l4' else 3) for m in ms)+' |')
  L.append('| Kuzatuvlar soni | '+' | '.join(str(O[m]['N']) for m in ms)+' |')
  h=O['lnZ Hausman']; L.append(f"| Hausman testi | – | χ²({h['df']}) = {n(h['H'],2)} (p = {n(h['p'],3)}) | – |")
  return '\n'.join(L)
def T6():
  ms=['S3 FE','S3 Pooled']
  L=['| O‘zgaruvchi | FE (bank effektlari) | Pooled OLS |','|---|---|---|']
  for v in ['S3_l1','KR_osish_l4','KR_osish_l8','lnTA','YaIM_osish','Asosiy_stavka','USD_osish','Inflyatsiya','DAVLAT']:
    L.append(f'| {NM[v]} | '+' | '.join(cell(m,v,4 if 'KR' in v else 3) for m in ms)+' |')
  L.append('| Kuzatuvlar soni | '+' | '.join(str(O[m]['N']) for m in ms)+' |')
  L.append(f"| R² | {n(O[ms[0]]['R2'],2)} (bank ichidagi) | {n(O[ms[1]]['R2'],2)} |")
  return '\n'.join(L)
def T2():
  D=O['deskriptiv']
  rows=[('MBII','Panel integral indeksi (MBII), ball',2),('bA','Aktivlar sifati bloki, ball',2),('bL','Likvidlik bloki, ball',2),('bD','Daromadlilik bloki, ball',2),('lnZ','ln(Z-score)',2),
   ('LEV','Kapital / aktivlar, %',1),('S3','Muammoli kreditlar ulushi, %',1),('HLA_r','Yuqori likvid aktivlar / aktivlar, %',1),('LTD','Kreditlar / depozitlar, %',1),('ROA','ROA, %',1),('ROE','ROE, %',1),
   ('lnTA','ln(aktivlar, mlrd so‘m)',2),('KR_osish','Kreditlar o‘sishi (yillik), %',1),('DEP_A','Depozitlar / aktivlar, %',1),('NIM','Sof foiz marjasi, %',1)]
  L=['| O‘zgaruvchi | Kuzatuvlar | O‘rtacha | Standart chetlanish | Minimum | Mediana | Maksimum |','|---|---|---|---|---|---|---|']
  for k,t,d in rows:
    x=D[k]; L.append(f"| {t} | {int(x['count'])} | {n(x['mean'],d)} | {n(x['std'],d)} | {n(x['min'],d)} | {n(x['50%'],d)} | {n(x['max'],d)} |")
  return '\n'.join(L)
txt=open(SRC).read()
for k,f in {'T2':T2,'T3':T3,'T4':T4,'T5':T5,'T6':T6}.items(): txt=txt.replace('{{'+k+'}}',f())
notes=[]
def sub(m):
  k,v=m.group(1),m.group(2)
  notes.append(A[v] if k=='AUTHOR' else ref(k,v))
  return f'[^{len(notes)}]'
txt=re.sub(r'%([A-Z][A-Z0-9]*):([^%]+)%',sub,txt)
txt=txt.rstrip()+'\n\n'+'\n'.join(f'[^{i+1}]: {n_}' for i,n_ in enumerate(notes))+'\n'
open(DST,'w').write(txt); print(len(notes),'notes')
