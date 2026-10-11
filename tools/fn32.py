import re
SRC='/home/user/tarjima/paragraflar/3.2-paragraf.src.md'
DST='/home/user/tarjima/paragraflar/3.2-paragraf.md'
R={}
for f in ['fn11.py','fn13x.py','fn22.py','fn12n.py','fn23.py','fn31.py']:
    code=open('/home/user/tarjima/tools/'+f).read(); code=code[:code.find('\ndef ref')]
    if f in ('fn12n.py','fn23.py','fn31.py'): code=code[code.find('\nR.update({'):]
    ns={'R':{}}; exec(code,ns); R.update(ns['R'])
R.update({
'BRBIFRS':('Joint-Stock Commercial Bank "Business Development Bank". Consolidated Financial Statements for the year ended 31 December 2025 and Independent Auditors’ Report. – Tashkent: KPMG Audit, 2026. – https://brb.uz/bank-haqida/moliyaviy-korsatkichlar','P.'),
'XALQIFRS':('Joint-stock commercial "Xalq Bank" of the Republic of Uzbekistan. Consolidated financial statements and independent auditor’s report for the year ended 31 December 2025. – Tashkent: Ernst & Young, 2026. – https://xb.uz/page/auditorlik-xulosalari-2','P.'),
'XALQIFRS22':('Joint-Stock Commercial "Xalq Bank". Consolidated financial statements and independent auditor’s report for the year ended 31 December 2022. – Tashkent: Ernst & Young, 2023.','P.'),
'IPOIFRS':('Joint Stock Commercial Mortgage Bank "Ipoteka Bank" and its subsidiaries. Consolidated Financial Statements and Independent Auditor’s Report for the year ended 31 December 2025. – Tashkent: Ernst & Young, 2026. – https://www.ipotekabank.uz/en/about/investors/','P.'),
'IPOIFRS23':('Joint Stock Commercial Mortgage Bank "Ipoteka Bank" and its subsidiaries. Consolidated Financial Statements and Independent Auditors’ Report for the year ended 31 December 2023. – Tashkent: Ernst & Young, 2024.','P.'),
'SNS':('«SNS» reyting agentligi. O‘zbekiston tijorat banklarining asosiy ko‘rsatkichlari bazasi (Banks_2025_IV, Banks_2026_II): 36 bank, 2021-yil IV chorak – 2026-yil II chorak.',''),
})
def ref(k,p):
  t,pp=R[k]
  return t if p=='-' else f'{t} – {pp} {p}.'
def J(*xs): return '; '.join(x.rstrip('.') for x in xs)+'.'
CALC='Tadqiqot natijasida muallif tomonidan quyidagi manbalar ma’lumotlari asosida hisoblandi: '
SRCS='Tadqiqot natijasida muallif tomonidan quyidagi manbalar asosida ishlab chiqildi: '
IF=(ref('SQBIFRS','2–3, 28–31, 74'),ref('MKIFRS25','11–12'),ref('BRBIFRS','8–9, 33, 69'),ref('XALQIFRS','1–2, 12, 56'),ref('IPOIFRS','1–2'))
A={
'T1':SRCS+J(ref('IMFFSI','-'),ref('N2709','-'),ref('SNS','-')),
'F1':CALC+J(*IF,ref('SNS','-')),
'T2':CALC+J(*IF,ref('SNS','-')),
'F2':'Tadqiqot natijasida muallif tomonidan 3.2.2-jadval ma’lumotlari asosida ishlab chiqildi.',
'T3':'Tadqiqot natijasida muallif tomonidan 3.2.1-rasm ma’lumotlari asosida ishlab chiqildi.',
'INJ':'Tadqiqot natijasida muallif tomonidan hisoblandi: kapitalga kiritilgan mablag‘lar = xususiy kapitalning yillik o‘zgarishi − yil sof foydasi (boshqa umumlashgan daromad va dividendlar ham shu farqqa kiradi, shu sababli baho taxminiy). Manbalar: '+J(*IF),
'F3':CALC+J(*IF),
'F4':CALC+J(ref('SQBREP','-'),ref('MKREP','-'),ref('SQBIFRS','2–3, 28–31, 74'),ref('MKIFRS25','11–12')),
'T4':CALC+J(ref('SNS','-')),
'F5':CALC+J(ref('SNS','-')),
}
txt=open(SRC).read(); notes=[]
def sub(m):
  k,v=m.group(1),m.group(2)
  notes.append(A[v] if k=='AUTHOR' else ref(k,v))
  return f'[^{len(notes)}]'
txt=re.sub(r'%([A-Z][A-Z0-9]*):([^%]+)%',sub,txt)
txt=txt.rstrip()+'\n\n'+'\n'.join(f'[^{i+1}]: {n}' for i,n in enumerate(notes))+'\n'
open(DST,'w').write(txt); print(len(notes),'notes')
