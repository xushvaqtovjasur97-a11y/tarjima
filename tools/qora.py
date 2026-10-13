"""Matndagi qizil (FF0000) va ko‘k (0070C0) ranglarni olib tashlash (qora/avtomatik). Chizma va diagrammalar ichidagi ranglar o‘zgarmaydi."""
import collections, sys
sys.path.insert(0, '/home/user/tarjima/tools')
import docx
from docx.oxml.ns import qn
from lxml import etree
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
COLS = {'FF0000', '0070C0'}
d = docx.Document(F)


def in_drawing(e):
    a = e.getparent()
    while a is not None:
        if a.tag in (qn('w:drawing'), qn('w:pict'), '{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent'):
            return True
        a = a.getparent()
    return False


def clean(root):
    n = 0; other = collections.Counter()
    for c in list(root.iter(qn('w:color'))):
        v = (c.get(qn('w:val')) or '').upper()
        if v in COLS and not in_drawing(c):
            c.getparent().remove(c); n += 1
        elif v not in ('', 'AUTO', '000000') and not in_drawing(c):
            other[v] += 1
    return n, other


n1, o1 = clean(d.element.body)
for part in [r.target_part for r in d.part.rels.values() if r.reltype.endswith('/footnotes')]:
    x = etree.fromstring(part.blob)
    n2, o2 = clean(x)
    part._blob = etree.tostring(x, xml_declaration=True, encoding='UTF-8', standalone=True)
d.save(F)
print('matn:', n1, 'snoskalar:', n2, '| qolgan boshqa ranglar:', dict(o1 + o2))
print(pardoz(F))
