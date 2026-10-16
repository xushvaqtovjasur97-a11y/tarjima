"""3.2.2-rasm: yorliqsiz chiziqlar, sinf chegaralari, ostida sonlar jadvali va izoh (qizil);
3.2-§: sinf chegaralari teng oraliqlar asosida, MBII-13 bo‘yicha qo‘shimcha tekshiruv (qizil)."""
import re, sys
sys.path.insert(0, '/home/user/tarjima/tools')
from docx.oxml.ns import qn
import kamlib as K
import wordart as W
import namuna_322r as R
import namuna_sinf13 as S
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
# 1) matn
D.replace(D.para(S.OLD1), S.OLD1, S.NEW1, tag='sinf')
D.insert(D.para(S.ANC2), S.ANC2, S.INS2, tag='sinf')
# 2) 3.2.2-rasm
cap = [p for p in D.d.paragraphs if p.text.startswith('3.2.2-rasm. Tanlangan banklarda')][0]
ch = cap._p.getprevious()
rid = [e.get(qn('r:id')) for e in ch.iter('{http://schemas.openxmlformats.org/drawingml/2006/chart}chart')][0]
part = D.d.part.related_parts[rid]
xrel = [r for r in part.rels.values() if r.reltype.endswith('/package')][0]
xml, rows = R.chart()
part._blob = re.sub(r'(<c:externalData r:id=")[^"]+(")', r'\g<1>' + xrel.rId + r'\g<2>', xml).encode('utf-8')
xrel.target_part._blob = W._xlsx(rows)
H = int(8.0 * 360000)
for ext in ch.iter('{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent'):
    ext.set('cy', str(H))
for ext in ch.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}ext'):
    if ext.getparent().tag.endswith('}xfrm'):
        ext.set('cy', str(H))
t = R.table(D.d, 'FF0000')
ch.addnext(t._tbl)
izsrc = D.para('Izoh: “Aloqabank” ATB empirik vaznlarni baholashda qatnashmagan')
iz = D.new_para_after(izsrc, R.IZOH, tag='sinf')
t._tbl.addnext(iz._p)
D.save(F)
print(part.partname, pardoz(F))
