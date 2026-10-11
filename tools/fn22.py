import re
SRC='/home/user/tarjima/paragraflar/2.2-paragraf.src.md'
DST='/home/user/tarjima/paragraflar/2.2-paragraf.md'
LM='Оценка финансовой устойчивости кредитной организации: учебник / под ред. О.И. Лаврушина, И.Д. Мамоновой. – М.: КНОРУС, 2011.'
LAV='Устойчивость банковской системы и развитие банковской политики: монография / под ред. О.И. Лаврушина. – М.: КНОРУС, 2014.'
BR='Bank Risk, Governance and Regulation / ed. by E. Beccalli, F. Poli. – Basingstoke: Palgrave Macmillan, 2015.'
R={
'LM':(LM,'С.'),'LAV':(LAV,'С.'),
'HEN':('Hendrickson J.M. Regulation and Instability in U.S. Commercial Banking: A History of Crises. – Basingstoke: Palgrave Macmillan, 2011.','P.'),
'TIAN':('Tian W. Regulatory Capital Requirement in Basel III // Commercial Banking Risk Management: Regulation in the Wake of the Financial Crisis / ed. by W. Tian. – New York: Palgrave Macmillan, 2017.','P.'),
'IMF145':('International Monetary Fund. Republic of Uzbekistan: Financial Sector Assessment Program – Financial System Stability Assessment. IMF Country Report No. 25/145. – Washington, DC: IMF, 2025.',''),
'IMF227':('International Monetary Fund. Republic of Uzbekistan: Financial Sector Assessment Program – Detailed Assessment of Observance – Basel Core Principles for Effective Banking Supervision. IMF Country Report No. 25/227. – Washington, DC: IMF, 2025.',''),
'WBFSA':('World Bank. Republic of Uzbekistan: Financial Sector Assessment. – Washington, DC: World Bank, 2025.',''),
'WBPAD':('World Bank. Uzbekistan Financial Sector Reform Project: Project Appraisal Document. Report No. PAD4468. – Washington, DC: World Bank, 2022.',''),
'CBU22':('O‘zbekiston Respublikasi Markaziy banki. Moliyaviy barqarorlik sharhi: 2022-yil. – Toshkent, 2023.',''),
'CBU23':('The Central Bank of the Republic of Uzbekistan. Financial Stability Report for 2023. – Tashkent, 2024.',''),
'CBU25H1':('The Central Bank of the Republic of Uzbekistan. Financial Stability Report for the first half of 2025. – Tashkent, 2025.',''),
'CBU25':('The Central Bank of the Republic of Uzbekistan. Financial Stability Report for 2025. – Tashkent, 2026.',''),
'GET':('Key trends in Uzbekistan’s banking sector // German Economic Team. – https://www.german-economic-team.com/en/newsletter/key-trends-in-uzbekistans-banking-sector/',''),
'KILDE':('Non-Bank Financial Institutions & Private Credit in Uzbekistan 2026 // Kilde. – https://www.kilde.sg/post/non-bank-financial-institutions-and-private-credit-investment-in-uzbekistan-2026-update',''),
'UZD':('World Bank. Republic of Uzbekistan: Financial Sector Assessment. – Washington, DC: World Bank, 2025; Private bank assets in Uzbekistan rise 29% in a year // UzDaily.uz. – 2026. – https://www.uzdaily.uz/en/private-bank-assets-in-uzbekistan-rise-29-in-a-year/',''),
'SQBREP':('«O‘zsanoatqurilishbank» ATB ning Markaziy bankka taqdim etilgan prudensial hisobot shakllari (crs003C, crs003L, crs003O): bank balansi, daromadlar, regulyativ kapital hisob-kitobi, aktivlarni tasniflash va prudensial me’yorlar jadvallari, 2021–2025-yillar 30-dekabr holatiga.',''),
'MKREP':('«Mikrokreditbank» ATB ning Markaziy bankka taqdim etilgan prudensial hisobot shakllari (crs005C, crs005L, crs005O): bank balansi, daromadlar, regulyativ kapital hisob-kitobi, aktivlarni tasniflash va prudensial me’yorlar jadvallari, 2021–2025-yillar 30-dekabr holatiga.',''),
'SQBIFRS':('Joint Stock Commercial Bank «Uzbek Industrial and Construction Bank» and its subsidiaries. Consolidated Financial Statements for the year ended 31 December 2025 and Independent Auditor’s Report. – Tashkent, 2026.','P.'),
}
def ref(k,p):
  t,pp=R[k]
  return t if p=='-' else f'{t} – {pp} {p}.'
def J(*xs): return '; '.join(x[:-1] for x in xs)+'.'
DEV='Tadqiqot natijasida muallif tomonidan ishlab chiqildi.'
SRCS='Tadqiqot natijasida muallif tomonidan quyidagi manbalar asosida ishlab chiqildi: '
SYS='Tadqiqot natijasida muallif tomonidan quyidagi manbalar ma’lumotlari asosida ishlab chiqildi: '
BANK=SYS+J(ref('SQBREP','-'),ref('MKREP','-'))
A={
'F1':SYS+J(ref('WBPAD','-'),ref('UZD','-')),
'T1':SYS+J(ref('IMF145','-'),ref('CBU22','-'),ref('CBU23','-'),ref('CBU25H1','-'),ref('CBU25','-'),ref('WBFSA','-'),ref('GET','-')),
'F2':'Tadqiqot natijasida muallif tomonidan 2.2.1-jadval ma’lumotlari asosida ishlab chiqildi.',
'F3':SYS+J(ref('IMF145','-'),ref('CBU22','-'),ref('CBU23','-'),ref('WBFSA','-')),
'F4':'Tadqiqot natijasida muallif tomonidan 2.2.1-jadval ma’lumotlari asosida ishlab chiqildi.',
'T2':'Tadqiqot natijasida muallif tomonidan quyidagi manbalar ma’lumotlari asosida hisoblandi: '+J(ref('SQBREP','-'),ref('MKREP','-')),
'F5':BANK,'F6':BANK,'F7':BANK,
'T3':SYS+J(ref('IMF145','-'),ref('CBU25','-'),ref('WBFSA','-'),ref('SQBREP','-'),ref('MKREP','-'),ref('SQBIFRS','28–31')),
}
txt=open(SRC).read(); notes=[]
def sub(m):
  k,v=m.group(1),m.group(2)
  notes.append(A[v] if k=='AUTHOR' else ref(k,v))
  return f'[^{len(notes)}]'
txt=re.sub(r'%([A-Z][A-Z0-9]*):([^%]+)%',sub,txt)
txt=txt.rstrip()+'\n\n'+'\n'.join(f'[^{i+1}]: {n}' for i,n in enumerate(notes))+'\n'
open(DST,'w').write(txt); print(len(notes),'notes')
