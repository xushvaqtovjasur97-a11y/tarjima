"""2.3.5-rasm: chiziqli diagramma (yorliqlar ustma-ust tushmaydi); matn (qizil)."""
import re, sys
sys.path.insert(0, '/home/user/tarjima/tools')
from docx.oxml.ns import qn
import kamlib as K
import wordart as W
import namuna_235 as N
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
cap = [p for p in D.d.paragraphs if p.text.startswith('2.3.5-rasm.')][0]
txt = [p for p in D.d.paragraphs if N.OLD in p.text][0]
ch = cap._p.getprevious()
rid = [e.get(qn('r:id')) for e in ch.iter('{http://schemas.openxmlformats.org/drawingml/2006/chart}chart')][0]
part = D.d.part.related_parts[rid]
xrel = [r for r in part.rels.values() if r.reltype.endswith('/package')][0]
xml, rows = N.chart()
part._blob = re.sub(r'(<c:externalData r:id=")[^"]+(")', r'\g<1>' + xrel.rId + r'\g<2>', xml).encode('utf-8')
xrel.target_part._blob = W._xlsx(rows)
for ext in ch.iter('{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent'):
    ext.set('cx', str(int(16.0 * 360000))); ext.set('cy', str(int(8.5 * 360000)))
for ext in ch.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}ext'):
    if ext.getparent().tag.endswith('}xfrm'):
        ext.set('cx', str(int(16.0 * 360000))); ext.set('cy', str(int(8.5 * 360000)))
del ch
D.replace(txt, N.OLD, N.NEW, tag='235')
D.save(F)
print(part.partname, pardoz(F))
