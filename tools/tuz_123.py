"""1.2.3-jadval: NSFR da 1-daraja likvid aktivlar koeffitsiyenti (0 → 0–5) va amaldagi Basel Framework manbasi (qizil)."""
import copy, sys
sys.path.insert(0, '/home/user/tarjima/tools')
from docx.oxml.ns import qn
import kamlib as K
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
cap = D.para('Bazel III likvidlik me’yorlarida qo‘llanadigan asosiy koeffitsiyentlar')
fid = [r.get(qn('w:id')) for r in cap._p.iter(qn('w:footnoteReference'))][0]
t = D.table('Moddalar | Koeffitsiyent')
rows = [i for i, r in enumerate(t.rows) if r.cells[0].text.strip() == '1-daraja likvid aktivlar']
assert len(rows) == 1 and t.rows[rows[0]].cells[1].text.strip() == '0'
ri = rows[0]
D.set_cell(t, ri, 1, '0–5', tag='123')
# bo‘sh “Mazmuni” katagi – shablon sifatida 1-ustun run’idan foydalanamiz
src = D._runs(t.rows[ri].cells[1].paragraphs[0])[0]
p = t.rows[ri].cells[2].paragraphs[0]
for r_ in list(p._p.iter(qn('w:r'))):
    r_.getparent().remove(r_)
for e in D._build(copy.deepcopy(src), 'Naqd pul va markaziy bankdagi zaxiralar – 0, boshqa 1-daraja aktivlar (masalan, davlat qimmatli qog‘ozlari) – 5'):
    p._p.append(e)
# jadval manbasi
D.insert(D.fn_para(fid), 'С. 69, 93–94', '; Basel Committee on Banking Supervision. The Basel Framework: LCR – Liquidity Coverage Ratio; NSF – Net Stable Funding Ratio (amaldagi tahriri). – Basel: BIS. – https://www.bis.org/basel_framework/', tag='123')
D.save(F)
print(pardoz(F))

# adabiyotlar ro‘yxatiga (II bo‘lim, alifbo tartibida “Stress testing principles” dan keyin) va qayta raqamlash
import re
D = K.Doc(F)
a = [p for p in D.paras() if re.match(r'^\d+\. ', p.text) and 'Basel Committee on Banking Supervision. Stress testing principles' in p.text][-1]
D.new_para_after(a, '0. Basel Committee on Banking Supervision. The Basel Framework: LCR – Liquidity Coverage Ratio; NSF – Net Stable Funding Ratio (amaldagi tahriri). – Basel: BIS. – https://www.bis.org/basel_framework/.', tag='123')
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
D.replace(hp, re.search(r'ro‘yxati \d+ ta manbani', hp.text).group(0), f'ro‘yxati {n} ta manbani', tag='123')
D.save(F)
print('adabiyotlar:', n, pardoz(F))
