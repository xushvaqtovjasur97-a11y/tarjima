import re
SRC='/home/user/tarjima/paragraflar/2.3-paragraf.src.md'
DST='/home/user/tarjima/paragraflar/2.3-paragraf.md'
R={}
for f in ['fn11.py','fn13x.py','fn22.py','fn12n.py']:
    code=open('/home/user/tarjima/tools/'+f).read(); code=code[:code.find('\ndef ref')]
    if f=='fn12n.py': code=code[code.find('R.update({'):]
    ns={'R':{}}; exec(code,ns); R.update(ns['R'])
R.update({
'PQ4611':('O‘zbekiston Respublikasi Prezidentining 2020-yil 24-fevraldagi PQ-4611-son «Moliyaviy hisobotning xalqaro standartlariga o‘tish bo‘yicha qo‘shimcha chora-tadbirlar to‘g‘risida»gi qarori // Qonunchilik ma’lumotlari milliy bazasi, 25.02.2020-y.',''),
'ZRU580':('O‘zbekiston Respublikasining 2019-yil 5-noyabrdagi O‘RQ-580-son «Banklar va bank faoliyati to‘g‘risida»gi Qonuni // Qonunchilik ma’lumotlari milliy bazasi, 06.11.2019-y.',''),
'N2696':('Tijorat banklarida aktivlar sifatini tasniflash va aktivlar bo‘yicha ehtimoliy yo‘qotishlarni qoplash uchun zaxiralarni shakllantirish hamda ulardan foydalanish tartibi to‘g‘risidagi nizom (O‘zbekiston Respublikasi Markaziy banki Boshqaruvining qarori, ro‘yxat raqami 2696, 14.07.2015) // https://lex.uz',''),
'N2709':('O‘zbekiston Respublikasi Markaziy banki Boshqaruvining 2015-yil 22-iyuldagi 19/14-son qarori bilan tasdiqlangan «Tijorat banklarining likvidliligini boshqarishga qo‘yiladigan talablar to‘g‘risida»gi nizom (ro‘yxat raqami 2709) // O‘zbekiston Respublikasi qonun hujjatlari to‘plami. – 2015. – 32 (688)-son. – 437-modda.',''),
'BCBS15':('Basel Committee on Banking Supervision. Guidance on credit risk and accounting for expected credit losses. – Basel: Bank for International Settlements, December 2015.',''),
'BCBS17':('Basel Committee on Banking Supervision. Regulatory treatment of accounting provisions – interim approach and transitional arrangements. – Basel: Bank for International Settlements, March 2017.',''),
'HL':('Huizinga H., Laeven L. Bank valuation and accounting discretion during a financial crisis // Journal of Financial Economics. – 2012. – Vol. 106, No. 3. – P. 614–634.',''),
'BHOL':('Bholat D., Lastra R.M., Markose S.M., Miglionico A., Sen K. Non-performing loans at the dawn of IFRS 9: regulatory and accounting treatment of asset quality // Journal of Banking Regulation. – 2018. – Vol. 19, No. 1. – P. 33–54.',''),
'ABAD':('Abad J., Suarez J. Assessing the cyclical implications of IFRS 9 – a recursive model // European Systemic Risk Board Occasional Paper Series. – 2017. – No. 12. – 48 p.',''),
'MKIFRS21':('Акционерно-коммерческий банк «Микрокредитбанк». Финансовая отчетность за год, закончившийся 31 декабря 2021 года, и заключение независимых аудиторов. – Ташкент, 2022.','С.'),
'MKIFRS23':('Акционерно-коммерческий банк «Микрокредитбанк». Финансовая отчетность за год, закончившийся 31 декабря 2023 года, и заключение независимых аудиторов. – Ташкент, 2024.','С.'),
'MKIFRS25':('Акционерно-коммерческий банк «Микрокредитбанк». Консолидированная финансовая отчетность за год, закончившийся 31 декабря 2025 года, и заключение независимых аудиторов. – Ташкент, 2026.','С.'),
})
def ref(k,p):
  t,pp=R[k]
  return t if p=='-' else f'{t} – {pp} {p}.'
def J(*xs): return '; '.join(x.rstrip('.') for x in xs)+'.'
SRCS='Tadqiqot natijasida muallif tomonidan quyidagi manbalar asosida ishlab chiqildi: '
SYS='Tadqiqot natijasida muallif tomonidan quyidagi manbalar ma’lumotlari asosida ishlab chiqildi: '
CALC='Tadqiqot natijasida muallif tomonidan quyidagi manbalar ma’lumotlari asosida hisoblandi: '
SQ=(ref('SQBREP','-'),ref('SQBIFRS21','8–9'),ref('SQBIFRS23','2–3, 39–40'),ref('SQBIFRS','2–3, 28–31, 37–38'))
MK=(ref('MKREP','-'),ref('MKIFRS21','6–7'),ref('MKIFRS23','7–8'),ref('MKIFRS25','11–12'),ref('MKIFRS','-'))
A={
'T1':SRCS+J(ref('N2696','-'),ref('IFRS9','-'),ref('SQBREP','-'),ref('SQBIFRS','37–38, 74'),ref('MKIFRS25','11–12')),
'F1':'Tadqiqot natijasida muallif tomonidan ishlab chiqildi.',
'T2':CALC+J(*SQ),
'F2':SYS+J(ref('SQBREP','-'),ref('SQBIFRS21','9'),ref('SQBIFRS23','3'),ref('SQBIFRS','3')),
'F3':SYS+J(*SQ),
'F4':SYS+J(ref('SQBIFRS','37–38')),
'F5':CALC+J(*(SQ+MK)),
'T3':CALC+J(*MK),
'F6':SYS+J(ref('MKREP','-'),ref('MKIFRS21','7'),ref('MKIFRS23','8'),ref('MKIFRS25','12')),
'T4':CALC+J(ref('SQBREP','-'),ref('SQBIFRS','2–3, 28–31, 74'),ref('N2709','-'),ref('B3NSFR','-'),ref('B3LCR','-')),
'F7':'Tadqiqot natijasida muallif tomonidan 2.3.4-jadval ma’lumotlari asosida ishlab chiqildi.',
}
txt=open(SRC).read(); notes=[]
def sub(m):
  k,v=m.group(1),m.group(2)
  notes.append(A[v] if k=='AUTHOR' else ref(k,v))
  return f'[^{len(notes)}]'
txt=re.sub(r'%([A-Z][A-Z0-9]*):([^%]+)%',sub,txt)
txt=txt.rstrip()+'\n\n'+'\n'.join(f'[^{i+1}]: {n}' for i,n in enumerate(notes))+'\n'
open(DST,'w').write(txt); print(len(notes),'notes')
