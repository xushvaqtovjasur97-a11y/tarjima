"""Qo‘lda: (1) 3.1.1-rasm (kirish gapi, rasm, sarlavha, tahlil) 3.1.1-jadvaldan oldinga – jadval to‘liq sahifa,
oldingi sahifadagi bo‘sh joyni rasm to‘ldiradi; (2) 2.1.2-jadval (bir necha sahifali) qatorlari sahifada bo‘linishi mumkin."""
import sys, re
from docx import Document
from docx.oxml.ns import qn
F = sys.argv[1]
d = Document(F)
els = list(d.element.body)
def tx(e): return ''.join(t.text or '' for t in e.iter(qn('w:t'))).strip()
h = next(e for e in els if tx(e) == '3.1.1-jadval')
cap = next(e for e in els if tx(e).startswith('3.1.1-rasm. '))
dr = cap.getprevious()
intro = next(e for e in els if tx(e).startswith('Bu vazifalar metodikaning hisoblash algoritmini belgilaydi'))
anal = next(e for e in els if tx(e).startswith('3.1.1-rasmdan ko‘rinib turibdiki'))
if els.index(intro) > els.index(h):
    for x in (intro, dr, cap, anal):
        h.addprevious(x)
    print('3.1.1-rasm jadvaldan oldinga o‘tkazildi')
t = next(e for e in els if tx(e) == '2.1.2-jadval')
tbl = t.getnext().getnext()
assert tbl.tag == qn('w:tbl')
n = 0
for tr in tbl.findall(qn('w:tr')):
    trpr = tr.find(qn('w:trPr'))
    if trpr is not None:
        for cs in trpr.findall(qn('w:cantSplit')):
            trpr.remove(cs); n += 1
print('2.1.2-jadval: cantSplit olib tashlandi', n)
d.save(F)
