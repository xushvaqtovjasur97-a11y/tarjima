"""1.3.2-rasm: AQSh PCA chegaralarini amaldagi 12 CFR § 324.403 ga moslashtirish (rasm – oddiy rangda, matn va manbalar – qizil)."""
import re, sys
sys.path.insert(0, '/home/user/tarjima/tools')
from docx.oxml.ns import qn
import kamlib as K
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
CFR = 'Code of Federal Regulations. Title 12, § 324.403: Capital measures and capital category definitions (2015-yil 1-yanvardan amaldagi tahriri). – https://www.ecfr.gov/current/title-12/section-324.403'
D = K.Doc(F)
cap = D.para('1.3.2-rasm. AQShda tezkor tuzatuvchi choralar tizimining kapital chegaralari')
an = D.para('Rasmdan ko‘rinib turibdiki, tizim kapital ko‘rsatkichini “bajarildi – bajarilmadi”')
fid = [r.get(qn('w:id')) for r in cap._p.iter(qn('w:footnoteReference'))][0]
drw = cap._p.getprevious()
REP = {'Jami kapital ≥ 10%; 1-darajali ≥ 6%; leverej ≥ 5%': 'Jami kapital ≥ 10%; 1-darajali ≥ 8%; CET1 ≥ 6,5%; leverej ≥ 5%',
       'Jami kapital ≥ 8%; 1-darajali ≥ 4%; leverej ≥ 4%': 'Jami kapital ≥ 8%; 1-darajali ≥ 6%; CET1 ≥ 4,5%; leverej ≥ 4%',
       'Jami kapital < 8%; 1-darajali < 4%; leverej < 4%': 'Jami kapital < 8%; 1-darajali < 6%; CET1 < 4,5%; leverej < 4%',
       'Jami kapital < 6%; 1-darajali < 3%; leverej < 3%': 'Jami kapital < 6%; 1-darajali < 4%; CET1 < 3%; leverej < 3%'}
n = 0
for t in drw.iter():
    if t.tag.endswith('}t') and t.text in REP:
        t.text = REP[t.text]; n += 1
assert n == 8, n
D.insert(D.fn_para(fid), 'P. 193 (Table 9.3)', '; ' + CFR, tag='132')
D.insert(an, 'balki besh pog‘onali shkala bilan baholaydi va har bir pog‘onani aniq nazorat oqibatlari bilan bog‘laydi.',
         ' 2015-yil 1-yanvardan Bazel III talablariga muvofiq tizimga asosiy kapital (CET1) mezoni qo‘shilgan va 1-darajali kapital chegaralari 2 foiz punktga oshirilgan[[FN:' + CFR + '.]].', tag='132')
D.save(F)

D = K.Doc(F)
a = [p for p in D.paras() if re.match(r'^\d+\. ', p.text) and 'Čihák M. Introduction to Applied Stress Testing' in p.text][-1]
D.new_para_after(a, '0. ' + CFR + '.', tag='132')
start = False; n = 0
for p in D.paras():
    if p.text.strip() == 'FOYDALANILGAN ADABIYOTLAR RO‘YXATI':
        start = True; continue
    if start and p.text.strip() == 'ILOVALAR':
        break
    if start:
        m = re.match(r'^(\d+)\. ', p.text)
        if m:
            n += 1
            r0 = D._runs(p)[0].find(qn('w:t'))
            if r0.text.startswith(m.group(1) + '.'):
                r0.text = str(n) + r0.text[len(m.group(1)):]
            elif r0.text == m.group(1):
                r0.text = str(n)
hp = D.para('Dissertatsiyaning tuzilishi va hajmi.')
D.replace(hp, re.search(r'ro‘yxati \d+ ta manbani', hp.text).group(0), f'ro‘yxati {n} ta manbani', tag='132')
D.save(F)
print('adabiyotlar:', n, pardoz(F))
