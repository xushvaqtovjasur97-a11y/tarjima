"""NSFR/LCR tuzatishlaridan keyingi sahifalash: 2.3.4-jadval LCR qatoridagi izoh jadval ostiga (yulduzcha bilan),
“Aloqabank” kirish xatboshisi 2.3.7-rasm sarlavhasidan keyinga qaytariladi."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
import kamlib as K

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)


def tx(el):
    return ''.join(x.text or '' for x in el.iter(qn('w:t')))


body = D.d.element.body
els = list(body)
tbl = next(e for e in els if e.tag == qn('w:tbl') and 'Likvidlikni qoplash koeffitsiyenti, %' in tx(e) and '222,9' in tx(e))
# 1) LCR qatori – tafovut katagi
row = [r for r in tbl.findall(qn('w:tr')) if 'Likvidlikni qoplash koeffitsiyenti, %' in tx(r)][0]
cell = row.findall(qn('w:tc'))[-1]
old = '−68,8 p.p. (banklardagi mablag‘lar yuqori likvid aktivlardan chiqarilganda)'
assert tx(cell) == old, tx(cell)
ts = list(cell.iter(qn('w:t')))
ts[0].text = '−68,8 p.p.*'
for t in ts[1:]:
    t.text = ''
# 2) jadval izohi
izoh = tbl.getnext()
assert tx(izoh).startswith('Izoh: kapital monandligi'), tx(izoh)
P = Paragraph(izoh, None)
o = 'NSFR va LCR – Bazel III metodikasi bo‘yicha.'
D.replace(P, o, 'NSFR va LCR – Bazel III metodikasi bo‘yicha; * – boshqa banklardagi mablag‘lar yuqori likvid aktivlardan '
                'chiqarilganda, 30 kunlik sof chiqim o‘zgarishsiz (quyi chegara).')
# 3) “Aloqabank” kirish xatboshisi – 2.3.7-rasm sarlavhasidan keyin
els = list(body)
intro = next(e for e in els if tx(e).startswith('Ikki bank misolida olingan natijalar umumiy xususiyatga egami'))
cap = next(e for e in els if tx(e).startswith('2.3.7-rasm. “O‘zsanoatqurilishbank” ATB ko‘rsatkichlarining'))
if els.index(intro) < els.index(cap):
    cap.addnext(intro)
D.save(F)
print('ok')
