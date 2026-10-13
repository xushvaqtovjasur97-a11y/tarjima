"""2.1.3-rasmni chiziqli diagramma bilan almashtirish (oddiy rangda)."""
import re, sys
sys.path.insert(0, '/home/user/tarjima/tools')
from docx.oxml.ns import qn
import kamlib as K
import wordart as W
import namuna_213l as N
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
cap = D.para('2.1.3-rasm. O‘zbekiston bank tizimida kapital monandligi ko‘rsatkichlari dinamikasi')
chp = cap._p.getprevious()
rid = [e.get(qn('r:id')) for e in chp.iter('{http://schemas.openxmlformats.org/drawingml/2006/chart}chart')][0]
part = D.d.part.related_parts[rid]
xrel = [r for r in part.rels.values() if r.reltype.endswith('/package')][0]
xml, rows = N.chart()
part._blob = re.sub(r'(<c:externalData r:id=")[^"]+(")', r'\g<1>' + xrel.rId + r'\g<2>', xml).encode('utf-8')
xrel.target_part._blob = W._xlsx(rows)
D.save(F)
print(part.partname, pardoz(F))
