import re
SRC='/home/user/tarjima/paragraflar/Kirish.src.md'
DST='/home/user/tarjima/paragraflar/Kirish.md'
R={}
for f in ['fn11.py','fn13x.py','fn22.py']:
    code=open('/home/user/tarjima/tools/'+f).read(); code=code[:code.find('def ref')]
    ns={}; exec(code,ns); R.update(ns['R'])
R.update({
'IMFFSI':('International Monetary Fund. Financial Soundness Indicators: Compilation Guide. – Washington, DC: IMF, 2019.',''),
'PF5992':('O‘zbekiston Respublikasi Prezidentining 2020-yil 12-maydagi PF-5992-son «2020–2025-yillarga mo‘ljallangan O‘zbekiston Respublikasining bank tizimini isloh qilish strategiyasi to‘g‘risida»gi Farmoni // Qonunchilik ma’lumotlari milliy bazasi, 13.05.2020-y.',''),
'N3697':('Banklar kapitalining monandligiga qo‘yiladigan talablar to‘g‘risidagi nizom (O‘zbekiston Respublikasi Markaziy banki Boshqaruvining 02.10.2025 dagi 21/6-son qarori, ro‘yxat raqami 3697, 28.10.2025) // https://lex.uz/docs/7793613',''),
'ORQ1070':('O‘zbekiston Respublikasining 2025-yil 23-iyundagi O‘RQ-1070-son «Banklarni sanatsiya qilish va tugatish to‘g‘risida»gi Qonuni // Qonunchilik ma’lumotlari milliy bazasi.',''),
'B3LCR':('Basel Committee on Banking Supervision. Basel III: The Liquidity Coverage Ratio and liquidity risk monitoring tools. – Basel: BIS, January 2013.',''),
'B3NSFR':('Basel Committee on Banking Supervision. Basel III: The Net Stable Funding Ratio. – Basel: BIS, October 2014.',''),
'B3FIN':('Basel Committee on Banking Supervision. Basel III: Finalising post-crisis reforms. – Basel: BIS, December 2017.',''),
'LV':('Laeven L., Valencia F. Systemic Banking Crises Database II // IMF Economic Review. – 2020. – Vol. 68, No. 2. – P. 307–361.',''),
'FET':('Фетисов Г.Г. Устойчивость банковской системы и методология ее оценки. – М.: Экономика, 2003. – 394 с.',''),
})
def ref(k,p):
  t,pp=R[k]
  return t if p=='-' else f'{t} – {pp} {p}.'
def J(*xs): return '; '.join(x.rstrip('.') for x in xs)+'.'
SRCS='Tadqiqot natijasida muallif tomonidan quyidagi manbalar ma’lumotlari asosida hisoblandi: '
A={
'BANKS':SRCS+J(ref('SQBREP','-'),ref('MKREP','-'),ref('SQBIFRS','28–31'),ref('MKIFRS','-')),
'SITES':'Dissertatsiya mavzusi bo‘yicha xorijiy ilmiy-tadqiqotlar sharhi quyidagi manbalar asosida ishlab chiqilgan: https://www.bis.org; https://www.imf.org; https://www.worldbank.org; https://www.ifrs.org; https://www.federalreserve.gov; https://www.ecb.europa.eu; https://www.eba.europa.eu; https://www.bankofengland.co.uk; https://www.cbr.ru; https://www.nfra.gov.cn; https://tsue.uz va boshqa manbalar.',
'LOCAL':J(
 'Sattarov O.B. O‘zbekiston Respublikasi bank tizimi barqarorligini ta’minlash metodologiyasini takomillashtirish: iqtisodiyot fanlari doktori (DSc) dissertatsiyasi avtoreferati. – Toshkent, 2018. – 67 b.',
 'Xolmamatov F.K. O‘zbekiston Respublikasining bank tizimining barqarorligini oshirish yo‘llari: iqtisodiyot fanlari doktori (DSc) dissertatsiyasi avtoreferati. – Toshkent: TMI, 2023. – 91 b.',
 ref('SHAR','-'),
 'Xolov N.Q. Tijorat banklari barqarorligini mustahkamlash metodologiyasini takomillashtirish: iqtisodiyot fanlari doktori (DSc) dissertatsiyasi. – Toshkent: TDIU, 2025.',
 'Madaminov B.A. Tijorat banklarida risk-menejment tizimini takomillashtirish asosida moliyaviy barqarorlikni ta’minlash metodologiyasi: iqtisodiyot fanlari doktori (DSc) dissertatsiyasi avtoreferati. – Toshkent, 2026. – 78 b.',
 'Abdullayev A.Ya. Tijorat banklarining likvidlilik darajasini oshirish metodologiyasini takomillashtirish: iqtisodiyot fanlari doktori (DSc) dissertatsiyasi avtoreferati. – Toshkent, 2018. – 74 b.',
 'Bozorov R.X. Tijorat banklari moliyaviy barqarorligini ta’minlashni takomillashtirish: iqtisodiyot fanlari bo‘yicha falsafa doktori (PhD) dissertatsiyasi avtoreferati. – Toshkent, 2020. – 52 b.',
 'Togayev S.S. Tijorat banklarining moliyaviy barqarorligini ta’minlash yo‘llari: iqtisodiyot fanlari bo‘yicha falsafa doktori (PhD) dissertatsiyasi. – Toshkent, 2022.',
 'Djalilov G.K. Banklarning moliyaviy barqarorligini ta’minlashning ustuvor yo‘nalishlari: iqtisodiyot fanlari bo‘yicha falsafa doktori (PhD) dissertatsiyasi avtoreferati. – Toshkent, 2026. – 63 b.',
 'Sadikov I.G. Iqtisodiyotni transformatsiyalash jarayonida tijorat banklarining moliyaviy barqarorligini ta’minlash yo‘llari: iqtisodiyot fanlari bo‘yicha falsafa doktori (PhD) dissertatsiyasi avtoreferati. – Toshkent, 2025. – 57 b.',
 'Abdishukurov Sh.Sh. Tijorat banklarining moliyaviy barqarorligini ta’minlash yo‘nalishlari: iqtisodiyot fanlari bo‘yicha falsafa doktori (PhD) dissertatsiyasi avtoreferati. – Toshkent, 2026. – 54 b.',
 'Karabayev N.A. Tijorat banklari faoliyatini reyting tizimi asosida baholashni takomillashtirish: iqtisodiyot fanlari bo‘yicha falsafa doktori (PhD) dissertatsiyasi avtoreferati. – Toshkent, 2018. – 54 b.',
 'Yodgorov M.N. Tijorat banklari kapitali barqarorligini ta’minlash yo‘llari: iqtisodiyot fanlari bo‘yicha falsafa doktori (PhD) dissertatsiyasi avtoreferati. – Toshkent, 2026. – 69 b.',
 'Sattorova N.G‘. Tijorat banklari likvidliligini optimallashtirish yo‘llari: iqtisodiyot fanlari bo‘yicha falsafa doktori (PhD) dissertatsiyasi avtoreferati. – Toshkent, 2026. – 69 b.',
 'Ismatov N.R. Bank kredit portfeli riskini baholash va boshqarishning metodologik asoslarini takomillashtirish: iqtisodiyot fanlari doktori (DSc) dissertatsiyasi avtoreferati. – Toshkent, 2026. – 73 b.',
 'Mahmudaliyeva Yo.A. Tijorat banklarining moliyaviy barqarorligi va uni baholash: iqtisod fanlari nomzodi dissertatsiyasi avtoreferati. – Toshkent, 2001. – 22 b.',
 ref('KS','-')),
}
txt=open(SRC).read(); notes=[]
def sub(m):
  k,v=m.group(1),m.group(2)
  notes.append(A[v] if k=='AUTHOR' else ref(k,v))
  return f'[^{len(notes)}]'
txt=re.sub(r'%([A-Z][A-Z0-9]*):([^%]+)%',sub,txt)
txt=txt.rstrip()+'\n\n'+'\n'.join(f'[^{i+1}]: {n}' for i,n in enumerate(notes))+'\n'
open(DST,'w').write(txt); print(len(notes),'notes')
