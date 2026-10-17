"""Qo‘lda: jadvaldan keyin bevosita kelgan rasmni (tahlili jadvaldan oldin turgan) jadvaldan oldinga o‘tkazish."""
import sys, re
from docx import Document
from docx.oxml.ns import qn
F = sys.argv[1]
d = Document(F)
els = list(d.element.body)
def tx(e): return ''.join(t.text or '' for t in e.iter(qn('w:t'))).strip()
EXT = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent'
def swap(jad, rasm, scale=1.0):
    els = list(d.element.body)
    h = next(e for e in els if tx(e) == jad)
    cap = next(e for e in els if tx(e).startswith(rasm + '. '))
    dr = cap.getprevious()
    assert dr.find('.//' + qn('w:drawing')) is not None
    if scale != 1.0:
        for ext in dr.iter(EXT):
            ext.set('cx', str(int(int(ext.get('cx')) * scale))); ext.set('cy', str(int(int(ext.get('cy')) * scale)))
        for ext in dr.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}ext'):
            ext.set('cx', str(int(int(ext.get('cx')) * scale))); ext.set('cy', str(int(int(ext.get('cy')) * scale)))
    h.addprevious(dr); h.addprevious(cap)
    print('ko‘chirildi:', rasm, '→', jad, 'dan oldin')
swap('2.2.1-jadval', '2.2.2-rasm', 0.92)
swap('2.3.2-jadval', '2.3.2-rasm')
d.save(F)
