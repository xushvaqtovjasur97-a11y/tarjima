"""Namuna: 1.1.1-jadval bir sahifada – jadvaldan keyingi matn qismini jadvaldan oldinga olib chiqish (ko‘k)."""
import sys, copy
sys.path.insert(0, '/home/user/tarjima/tools')
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
import kamlib as K
K.RED = '0070C0'
SRC, OUT, SPLIT = sys.argv[1], sys.argv[2], sys.argv[3]   # SPLIT – jadvaldan keyin davom etadigan gap boshi
D = K.Doc(SRC)
body = D.d.element.body
def tx(e): return ''.join(x.text or '' for x in e.iter(qn('w:t')))
els = list(body)
head = next(e for e in els if tx(e).strip() == '1.1.1-jadval')
title = head.getnext(); tbl = title.getnext(); blank = tbl.getnext()
assert tbl.tag == qn('w:tbl') and tx(blank) == ''
p1 = blank.getnext(); p2 = p1.getnext()
assert tx(p1).startswith('Jadval ma’lumotlari shuni ko‘rsatadiki') and tx(p2).startswith('Rossiya iqtisodiy maktabining')
P1 = Paragraph(p1, None)
D.replace(P1, 'Jadval ma’lumotlari shuni ko‘rsatadiki, ta’riflardagi', 'Bu ta’riflardagi')
head.addprevious(p1)
# ikkinchi xatboshi: SPLIT gacha bo‘lgan qismi jadvaldan oldin, qolgani jadvaldan keyin
P2 = Paragraph(p2, None)
i, txt = D._text_index(P2, SPLIT)
first = D._split_at(P2, i)
p2b = copy.deepcopy(p2)
# p2 da first dan boshlab hamma run/izohni o‘chirish; p2b da first gacha bo‘lganlarini
kids = [c for c in p2 if c.tag != qn('w:pPr')]
k = kids.index(first)
for c in kids[k:]: p2.remove(c)
kidsb = [c for c in p2b if c.tag != qn('w:pPr')]
for c in kidsb[:k]: p2b.remove(c)
# p2 oxiridagi bo‘sh joy
last = [r for r in p2.iter(qn('w:t'))][-1]; last.text = last.text.rstrip()
head.addprevious(p2)
blank.addnext(p2b)
D.insert(Paragraph(p2b, None), '', '') if False else None
P2b = Paragraph(p2b, None)
D._normalize(P2b)
r0 = D._runs(P2b)[0]
for e in reversed(D._build(r0, 'Xususan, ')):
    r0.addprevious(e)
# jadval bir sahifada yaxlit
for e in (head, title):
    Paragraph(e, None).paragraph_format.keep_with_next = True
rows = tbl.findall(qn('w:tr'))
for tr in rows[:-1]:
    for p in tr.iter(qn('w:p')):
        Paragraph(p, None).paragraph_format.keep_with_next = True
D.save(OUT)
print('ok', tx(p2)[-80:], '||', tx(p2b)[:80])
