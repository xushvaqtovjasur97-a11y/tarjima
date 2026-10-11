import re, json
SRC='/home/user/tarjima/paragraflar/4.1-paragraf.src.md'
DST='/home/user/tarjima/paragraflar/4.1-paragraf.md'
R={}
for f in ['fn11.py','fn13x.py','fn22.py','fn12n.py','fn23.py','fn31.py','fn32.py','fn33.py']:
    code=open('/home/user/tarjima/tools/'+f).read(); code=code[:code.find('\ndef ref')]
    if f in ('fn12n.py','fn23.py','fn31.py','fn32.py','fn33.py'): code=code[code.find('\nR.update({'):]
    ns={'R':{}}; exec(code,ns); R.update(ns['R'])
R.update({
'KR':('Kaminsky G.L., Reinhart C.M. The Twin Crises: The Causes of Banking and Balance-of-Payments Problems // American Economic Review. – 1999. – Vol. 89, No. 3. – P. 473–500.',''),
'KLR':('Kaminsky G., Lizondo S., Reinhart C.M. Leading Indicators of Currency Crises // IMF Staff Papers. – 1998. – Vol. 45, No. 1. – P. 1–48.',''),
'DKD':('Demirgüç-Kunt A., Detragiache E. The Determinants of Banking Crises in Developing and Developed Countries // IMF Staff Papers. – 1998. – Vol. 45, No. 1. – P. 81–109.',''),
'DAVK':('Davis E.P., Karim D. Comparing early warning systems for banking crises // Journal of Financial Stability. – 2008. – Vol. 4, No. 2. – P. 89–120.',''),
'CG':('Cole R.A., Gunther J.W. Predicting Bank Failures: A Comparison of On- and Off-Site Monitoring Systems // Journal of Financial Services Research. – 1998. – Vol. 13, No. 2. – P. 103–117.',''),
'BETZ':('Betz F., Oprică S., Peltonen T.A., Sarlin P. Predicting distress in European banks // Journal of Banking & Finance. – 2014. – Vol. 45. – P. 225–241.',''),
'SVB':('Sahajwala R., Van den Bergh P. Supervisory risk assessment and early warning systems: Basel Committee on Banking Supervision Working Papers No. 4. – Basel: BIS, December 2000.',''),
'DELONG':('DeLong E.R., DeLong D.M., Clarke-Pearson D.L. Comparing the Areas under Two or More Correlated Receiver Operating Characteristic Curves: A Nonparametric Approach // Biometrics. – 1988. – Vol. 44, No. 3. – P. 837–845.',''),
'FAW':('Fawcett T. An introduction to ROC analysis // Pattern Recognition Letters. – 2006. – Vol. 27, No. 8. – P. 861–874.',''),
'HOS':('Hosmer D.W., Lemeshow S., Sturdivant R.X. Applied Logistic Regression. – 3rd ed. – Hoboken, NJ: John Wiley & Sons, 2013.',''),
'ALT':('Altman E.I. Financial Ratios, Discriminant Analysis and the Prediction of Corporate Bankruptcy // Journal of Finance. – 1968. – Vol. 23, No. 4. – P. 589–609.',''),
})
def ref(k,p):
  t,pp=R[k]
  return t if p=='-' else f'{t} – {pp} {p}.'
def J(*xs): return '; '.join(x.rstrip('.') for x in xs)+'.'
SRCS='Tadqiqot natijasida muallif tomonidan quyidagi manbalar asosida ishlab chiqildi: '
CALC='Tadqiqot natijasida muallif tomonidan quyidagi manbalar ma’lumotlari asosida hisoblandi: '
DATA=(ref('CBUBS','-'),ref('SNS','-'))
SAME='Tadqiqot natijasida muallif tomonidan 4.1.2-jadvalda ko‘rsatilgan manbalar ma’lumotlari asosida hisoblandi.'
A={'T1':SRCS+J(ref('SVB','-'),ref('CG','-'),ref('LM','45, 107–114'),ref('KR','-'),ref('DKD','-'),ref('CP','83–95'),ref('BETZ','-')),
'T2':CALC+J(*DATA),'T3':SAME,'F1':SAME,'T4':SAME,'F2':SAME,'T5':SAME,'F3':SAME,
'T6':'Tadqiqot natijasida muallif tomonidan 4.1.3–4.1.5-jadvallar ma’lumotlari asosida ishlab chiqildi.'}
O=json.load(open('/home/user/tarjima/manbalar/hisob_kitob/ews41.json'))
def n(x,d): return f'{x:.{d}f}'.replace('.',',').replace('-','−')
def st(p): return '\\*\\*\\*' if p<0.01 else '\\*\\*' if p<0.05 else '\\*' if p<0.1 else ''
Q={'I':'I','II':'II','III':'III','IV':'IV'}
def dv(s): y,q=s.split('_'); return f'{y}-y. {q} chorak'
def T2():
  L=['| Bank | Chorak | Chorak zararining kapitalga nisbati, % |','|---|---|---|']
  for b,dd,l in O['hodisalar']: L.append(f'| «{b}» ATB | {dv(dd)} | {n(l,1)} |')
  return '\n'.join(L)
def T3():
  L=['| Ko‘rsatkich | Kuzatuvlar soni | Shundan hodisadan oldingi | AUC | 95 % ishonch oralig‘i |','|---|---|---|---|---|']
  for k,N,pos,a,lo,hi in sorted(O['AUC'],key=lambda r:-r[3]): L.append(f'| {k} | {N} | {pos} | {n(a,3)} | {n(lo,2)} – {n(hi,2)} |')
  return '\n'.join(L)
def T4():
  m=O['logit']; ks=['1: MBII','2: ln(Z)','3: MBII + ln(Z)','4: Bloklar','5: MBII + davlat']
  V=[('const','Konstanta'),('MBII','MBII'),('lnZ','ln(Z-score)'),('bA','Aktivlar sifati bloki'),('bL','Likvidlik bloki'),('bD','Daromadlilik bloki'),('DAVLAT','Davlat banki (dummy)')]
  L=['| O‘zgaruvchi | 1-model | 2-model | 3-model | 4-model | 5-model |','|---|---|---|---|---|---|']
  for v,t in V:
    L.append(f'| {t} | '+' | '.join(f"{n(m[k]['b'][v],2)}{st(m[k]['p'][v])} ({n(m[k]['se'][v],2)})" if v in m[k]['b'] else '–' for k in ks)+' |')
  L.append('| Kuzatuvlar soni | '+' | '.join(str(m[k]['N']) for k in ks)+' |')
  L.append('| Psevdo-R² | '+' | '.join(n(m[k]['pseudoR2'],2) for k in ks)+' |')
  L.append('| AUC | '+' | '.join(n(m[k]['AUC'],3) for k in ks)+' |')
  return '\n'.join(L)
def T5():
  L=['| MBII chegarasi | To‘g‘ri signal | O‘tkazib yuborilgan hodisa | Noto‘g‘ri signal | I tur xato, % | II tur xato, % | Shovqin/signal nisbati |','|---|---|---|---|---|---|---|']
  for r in sorted(O['signal'],key=lambda r:r['th']):
    L.append(f"| {n(r['th'],2)} | {r['TP']} | {r['FN']} | {r['FP']} | {n(r['I_tur'],1)} | {n(r['II_tur'],1)} | {n(r['NSR'],2)} |")
  return '\n'.join(L)
txt=open(SRC).read()
for k,f in {'T2':T2,'T3':T3,'T4':T4,'T5':T5}.items(): txt=txt.replace('{{'+k+'}}',f())
notes=[]
def sub(m):
  k,v=m.group(1),m.group(2)
  notes.append(A[v] if k=='AUTHOR' else ref(k,v))
  return f'[^{len(notes)}]'
txt=re.sub(r'%([A-Z][A-Z0-9]*):([^%]+)%',sub,txt)
txt=txt.rstrip()+'\n\n'+'\n'.join(f'[^{i+1}]: {x}' for i,x in enumerate(notes))+'\n'
open(DST,'w').write(txt); print(len(notes),'notes')
