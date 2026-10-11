import re
SRC='/home/user/tarjima/paragraflar/4.3-paragraf.src.md'
DST='/home/user/tarjima/paragraflar/4.3-paragraf.md'
R={}
for f in ['fn11.py','fn13x.py','fn22.py','fn12n.py','fn23.py','fn31.py','fn32.py','fn33.py','fn41.py','fn42.py']:
    code=open('/home/user/tarjima/tools/'+f).read(); code=code[:code.find('\ndef ref')]
    if f not in ('fn11.py','fn13x.py','fn22.py'): code=code[code.find('\nR.update({'):]
    ns={'R':{}}; exec(code,ns); R.update(ns['R'])
R.update({
'IMFP':('International Monetary Fund. Republic of Uzbekistan: Financial Sector Assessment Program – Financial System Stability Assessment. IMF Country Report No. 25/145. – Washington, DC: IMF, 2025.','P.'),
'RM':('The Central Bank of the Republic of Uzbekistan. «Roadmap» aimed at further improving the bank regulation and supervision system in 2025–2028. – Tashkent, 17 November 2025.',''),
'PCA':('Federal Deposit Insurance Corporation Improvement Act of 1991, Section 131: Prompt Corrective Action (12 U.S.C. § 1831o). – Washington, DC: U.S. Congress, 1991.',''),
'SREP':('European Banking Authority. Guidelines on common procedures and methodologies for the supervisory review and evaluation process (SREP) and supervisory stress testing (EBA/GL/2022/03). – Paris: EBA, 18 March 2022.',''),
})
def ref(k,p):
  t,pp=R[k]
  return t if p=='-' else f'{t} – {pp} {p}.'
def J(*xs): return '; '.join(x.rstrip('.') for x in xs)+'.'
SRCS='Tadqiqot natijasida muallif tomonidan quyidagi manbalar asosida ishlab chiqildi: '
A={'T1':SRCS+J(ref('IMFP','1, 8, 20, 24, 28–29'),ref('RM','-'),'dissertatsiyaning 2.3, 3.1–3.3, 4.1, 4.2-paragraflari natijalari'),
'F1':'Tadqiqot natijasida muallif tomonidan ishlab chiqildi.',
'T3':'Tadqiqot natijasida muallif tomonidan quyidagi manbalar ma’lumotlari asosida hisoblandi: '+J(ref('CBUBS','-'),ref('SNS','-'),ref('SQBIFRS','2–3, 28–31'),ref('MKIFRS25','11–12'),ref('BRBIFRS','8–9, 33, 69'),ref('XALQIFRS','1–2, 12, 56'),ref('IPOIFRS','1–2')),
'T4':SRCS+J(ref('IMFP','8, 28'),ref('IMF227','-'),ref('RM','-'),'dissertatsiyaning 2.1, 3.1–3.3, 4.1, 4.2-paragraflari natijalari'),
'T2':SRCS+J(ref('RM','-'),'dissertatsiyaning 4.1, 4.2-paragraflari natijalari')}
txt=open(SRC).read(); notes=[]
def sub(m):
  k,v=m.group(1),m.group(2)
  notes.append(A[v] if k=='AUTHOR' else ref(k,v))
  return f'[^{len(notes)}]'
txt=re.sub(r'%([A-Z][A-Z0-9]*):([^%]+)%',sub,txt)
txt=txt.rstrip()+'\n\n'+'\n'.join(f'[^{i+1}]: {x}' for i,x in enumerate(notes))+'\n'
open(DST,'w').write(txt); print(len(notes),'notes')
