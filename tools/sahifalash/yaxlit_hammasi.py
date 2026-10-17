"""Barcha jadval va rasmlarni bir sahifada yaxlit saqlash (keepNext zanjiri):
jadval – nomer, nomi, oxirgi qatordan boshqa barcha qatorlar (Izoh bo‘lsa oxirgi qator ham);
rasm – rasm xatboshisidan sarlavhagacha bo‘lgan barcha elementlar (sonlar jadvali bilan), Izoh bo‘lsa sarlavha ham.
Foydalanish: python3 yaxlit_hammasi.py <docx> [--skip Table5,Table13 …] – skip: UNO nomlari (bir sahifaga sig‘maydiganlar)."""
import sys, re
from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

F = sys.argv[1]
SKIP = set(sys.argv[sys.argv.index('--skip') + 1].split(',')) if '--skip' in sys.argv else set()
d = Document(F)
body = d.element.body
els = list(body)
def tx(e): return ''.join(t.text or '' for t in e.iter(qn('w:t'))).strip()
def kn(p, v=True): Paragraph(p, None).paragraph_format.keep_with_next = v
def has_draw(e): return e.find('.//' + qn('w:drawing')) is not None or e.find('.//' + qn('w:pict')) is not None
IZOH = re.compile(r'^(Izoh|Manba|\*)')
tblno = 0; n_t = n_r = 0
fig_tbls = set()
# rasmlar
for i, e in enumerate(els):
    if e.tag == qn('w:p') and re.match(r'^\d\.\d\.\d+-rasm\.', tx(e)):
        j = i - 1
        chain = []
        while j >= 0 and j > i - 6:
            x = els[j]
            if x.tag == qn('w:tbl'):
                chain.append(x); fig_tbls.add(x); j -= 1; continue
            if x.tag == qn('w:p') and (has_draw(x) or tx(x) == ''):
                chain.append(x)
                if has_draw(x): break
                j -= 1; continue
            break
        if not any(x.tag == qn('w:p') and has_draw(x) for x in chain):
            continue
        for x in chain:
            if x.tag == qn('w:tbl'):
                for p in x.iter(qn('w:p')): kn(p)
            else:
                kn(x)
        nx = e.getnext()
        if nx is not None and nx.tag == qn('w:p') and IZOH.match(tx(nx)):
            kn(e)
        n_r += 1
# jadvallar (UNO nomlari hujjat tartibida: Table1, Table2 …)
for e in els:
    if e.tag != qn('w:tbl'):
        continue
    tblno += 1
    name = f'Table{tblno}'
    if e in fig_tbls:
        continue
    if name in SKIP:          # bir sahifaga sig‘maydi: bo‘linadi, faqat nomi birinchi qator bilan
        for tr in e.findall(qn('w:tr')):
            for p in tr.iter(qn('w:p')): kn(p, False)
        pv = e.getprevious()
        for _ in range(2):
            if pv is not None and pv.tag == qn('w:p') and 0 < len(tx(pv)) < 260: kn(pv); pv = pv.getprevious()
        continue
    rows = e.findall(qn('w:tr'))
    for tr in rows[:-1]:
        for p in tr.iter(qn('w:p')): kn(p)
    nx = e.getnext()
    if nx is not None and nx.tag == qn('w:p') and IZOH.match(tx(nx)):
        for p in rows[-1].iter(qn('w:p')): kn(p)
    # nomi va raqami
    pv = e.getprevious()
    for _ in range(2):
        if pv is not None and pv.tag == qn('w:p') and 0 < len(tx(pv)) < 260 and not has_draw(pv):
            kn(pv); pv = pv.getprevious()
    n_t += 1
d.save(F)
print('jadval', n_t, 'rasm', n_r, 'jami jadval', tblno)
