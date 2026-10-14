"""2.3.7-rasm: tafovut zanjiri tartibi, bloklar rangi, vergulli yorliqlar; 366-izohda 2.3.2-jadval (qizil)."""
import re, sys
sys.path.insert(0, '/home/user/tarjima/tools')
from docx.oxml.ns import qn
import kamlib as K
import wordart as W
import namuna_237r as N
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
cap = [p for p in D.d.paragraphs if p.text.startswith('2.3.7-rasm.')][0]
ch = cap._p.getprevious()
rid = [e.get(qn('r:id')) for e in ch.iter('{http://schemas.openxmlformats.org/drawingml/2006/chart}chart')][0]
part = D.d.part.related_parts[rid]
xrel = [r for r in part.rels.values() if r.reltype.endswith('/package')][0]
xml, rows = N.chart()
part._blob = re.sub(r'(<c:externalData r:id=")[^"]+(")', r'\g<1>' + xrel.rId + r'\g<2>', xml).encode('utf-8')
xrel.target_part._blob = W._xlsx(rows)
H = int(9.5 * 360000)
for ext in ch.iter('{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent'):
    ext.set('cy', str(H))
for ext in ch.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}ext'):
    if ext.getparent().tag.endswith('}xfrm'):
        ext.set('cy', str(H))
# rasm nomi ostidagi izoh
fid = [r.get(qn('w:id')) for r in cap._p.iter(qn('w:footnoteReference'))][0]
D.replace(D.fn_para(fid=fid), N.FN_OLD, N.FN_NEW, tag='237r')
D.save(F)
print(part.partname, fid, pardoz(F))
