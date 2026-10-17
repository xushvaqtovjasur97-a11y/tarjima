import re
SRC='/home/user/tarjima/paragraflar/3.1-paragraf.src.md'
DST='/home/user/tarjima/paragraflar/3.1-paragraf.md'
R={}
for f in ['fn11.py','fn13x.py','fn22.py','fn12n.py','fn23.py']:
    code=open('/home/user/tarjima/tools/'+f).read(); code=code[:code.find('\ndef ref')]
    if f in ('fn12n.py','fn23.py'): code=code[code.find('\nR.update({'):]
    ns={'R':{}}; exec(code,ns); R.update(ns['R'])
R.update({
'N3697':('Banklar kapitalining monandligiga qo‘yiladigan talablar to‘g‘risidagi nizom (O‘zbekiston Respublikasi Markaziy banki Boshqaruvining 02.10.2025 dagi 21/6-son qarori, ro‘yxat raqami 3697, 28.10.2025) // https://lex.uz/docs/7793613',''),
'B3GRF':('Basel Committee on Banking Supervision. Basel III: A global regulatory framework for more resilient banks and banking systems. – Basel: Bank for International Settlements, December 2010 (rev. June 2011).',''),
'DSIB':('Basel Committee on Banking Supervision. A framework for dealing with domestic systemically important banks. – Basel: Bank for International Settlements, October 2012.',''),
})
def ref(k,p):
  t,pp=R[k]
  return t if p=='-' else f'{t} – {pp} {p}.'
def J(*xs): return '; '.join(x.rstrip('.') for x in xs)+'.'
DEV='Tadqiqot natijasida muallif tomonidan ishlab chiqildi.'
SRCS='Tadqiqot natijasida muallif tomonidan quyidagi manbalar asosida ishlab chiqildi: '
CALC='Tadqiqot natijasida muallif tomonidan quyidagi manbalar ma’lumotlari asosida hisoblandi: '
A={
'T1':'Tadqiqot natijasida muallif tomonidan dissertatsiyaning 2.1–2.3-paragraflari natijalari asosida ishlab chiqildi.',
'F1':DEV,
'T2':SRCS+J(ref('B3GRF','-'),ref('DSIB','-'),ref('B3LCR','-'),ref('B3NSFR','-'),ref('IFRS9','-'),ref('IMFFSI','-'),ref('N3697','-'),ref('N2709','-'),ref('LM','135, 145–147'),ref('LAV','52–53'),ref('TAR','228–230')),
'F2':'Tadqiqot natijasida muallif tomonidan 3.1.2-jadval asosida ishlab chiqildi.',
'T3':SRCS+J(ref('KS','213'),ref('LM','68'),ref('IZZI','76, 85')),
'T4':SRCS+J(ref('CAR','193'),ref('LM','91–92')),
'T5':CALC+J(ref('SQBREP','-'),ref('MKREP','-'),ref('SQBIFRS','2–3, 28–31, 37–38, 74'),ref('MKIFRS25','11–12, 48')),
'F3':'Tadqiqot natijasida muallif tomonidan 3.1.5-jadval ma’lumotlari asosida ishlab chiqildi.',
'F4':'Tadqiqot natijasida muallif tomonidan 3.1.3 va 3.1.5-jadvallar ma’lumotlari asosida hisoblandi.',
'T6':SRCS+J(ref('LM','91–92, 135–154'),ref('CAR','193'),ref('IMFFSI','-'),ref('KS','208–213')),
}
txt=open(SRC).read(); notes=[]
def sub(m):
  k,v=m.group(1),m.group(2)
  notes.append(A[v] if k=='AUTHOR' else ref(k,v))
  return f'[^{len(notes)}]'
txt=re.sub(r'%([A-Z][A-Z0-9]*):([^%]+)%',sub,txt)
txt=txt.rstrip()+'\n\n'+'\n'.join(f'[^{i+1}]: {n}' for i,n in enumerate(notes))+'\n'
open(DST,'w').write(txt); print(len(notes),'notes')
