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
'CBUBS':('O‘zbekiston Respublikasi Markaziy banki. Banklar statistikasi: tijorat banklarining asosiy ko‘rsatkichlari va muammoli kreditlari to‘g‘risidagi oylik ma’lumotlar, 2021-yil aprel – 2026-yil yanvar. – https://cbu.uz/uz/statistics/bankstats/',''),
'OECDCI':('OECD, JRC. Handbook on Constructing Composite Indicators: Methodology and User Guide. – Paris: OECD Publishing, 2008.','P.'),
'SHAN':('Shannon C.E. A Mathematical Theory of Communication // Bell System Technical Journal. – 1948. – Vol. 27, No. 3. – P. 379–423.',''),
'SST':('Saisana M., Saltelli A., Tarantola S. Uncertainty and sensitivity analysis techniques as tools for the quality assessment of composite indicators // Journal of the Royal Statistical Society: Series A. – 2005. – Vol. 168, No. 2. – P. 307–323.',''),
'IL':('Illing M., Liu Y. Measuring financial stress in a developed country: An application to Canada // Journal of Financial Stability. – 2006. – Vol. 2, No. 3. – P. 243–265.',''),
'SNS':('«SNS» reyting agentligi. O‘zbekiston tijorat banklarining asosiy ko‘rsatkichlari bazasi (Banks_2025_IV, Banks_2026_II): 36 bank, 2021-yil IV chorak – 2026-yil II chorak.',''),
})
def ref(k,p):
  t,pp=R[k]
  return t if p=='-' else f'{t} – {pp} {p}.'
def J(*xs): return '; '.join(x.rstrip('.') for x in xs)+'.'
CALC='Tadqiqot natijasida muallif tomonidan quyidagi manbalar ma’lumotlari asosida hisoblandi: '
SRCS='Tadqiqot natijasida muallif tomonidan quyidagi manbalar asosida ishlab chiqildi: '
IF=(ref('SQBIFRS','2–3, 28–31, 74'),ref('MKIFRS25','11–12'),ref('BRBIFRS','8–9, 33, 69'),ref('XALQIFRS','1–2, 12, 56'),ref('IPOIFRS','1–2'),ref('SNS','-'))
PAN=(ref('CBUBS','-'),)
ALL=CALC+J(*IF,*PAN)
A={'T1':SRCS+J(ref('OECDCI','89–103'),ref('SHAN','-'),ref('SST','-'),ref('IL','-'),ref('KARM','89–91')),
'T2':ALL,'F1':'Tadqiqot natijasida muallif tomonidan 3.2.2-jadval ma’lumotlari asosida ishlab chiqildi.','T3':ALL,'F2':ALL,'T4':ALL,'F3':ALL,'INJ':'Kapitalga kiritilgan mablag‘lar xususiy kapitalning yillik o‘zgarishi va yil sof foydasi o‘rtasidagi farq sifatida baholandi. Manbalar: '+J(*IF[:5])}
txt=open(SRC).read(); notes=[]
def sub(m):
  k,v=m.group(1),m.group(2)
  notes.append(A[v] if k=='AUTHOR' else ref(k,v))
  return f'[^{len(notes)}]'
txt=re.sub(r'%([A-Z][A-Z0-9]*):([^%]+)%',sub,txt)
txt=txt.rstrip()+'\n\n'+'\n'.join(f'[^{i+1}]: {n}' for i,n in enumerate(notes))+'\n'
open(DST,'w').write(txt); print(len(notes),'notes')
