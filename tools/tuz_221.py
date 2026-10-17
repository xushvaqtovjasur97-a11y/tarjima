"""2.2.1-rasmni chiziqli diagramma bilan almashtirish (oddiy rangda) va manbani 5.1.2-jadvalga aniqlashtirish (qizil)."""
import re, sys
sys.path.insert(0, '/home/user/tarjima/tools')
from docx.oxml.ns import qn
import kamlib as K
import wordart as W
import namuna_221l as N
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
OLD1 = 'O‘zbekiston Respublikasi Markaziy banki. Statistik byulleten: 2021-yil. – Toshkent, 2022. – 5.1.8b-jadval; '
OLD2 = '5.1.2, 5.1.3, 5.1.4, 5.1.6, 5.1.7, 5.1.8-jadvallar'
D = K.Doc(F)
cap = D.para('2.2.1-rasm. Davlat ulushi mavjud banklarning bank tizimidagi ulushi')
chp = cap._p.getprevious()
rid = [e.get(qn('r:id')) for e in chp.iter('{http://schemas.openxmlformats.org/drawingml/2006/chart}chart')][0]
part = D.d.part.related_parts[rid]
xrel = [r for r in part.rels.values() if r.reltype.endswith('/package')][0]
xml, rows = N.chart()
part._blob = re.sub(r'(<c:externalData r:id=")[^"]+(")', r'\g<1>' + xrel.rId + r'\g<2>', xml).encode('utf-8')
xrel.target_part._blob = W._xlsx(rows)
for ext in chp.iter('{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent'):
    print('extent', ext.get('cx'), ext.get('cy'))
ids = [i for i in D.fn_ids(OLD1) if i in D.fn_ids(OLD2) and i in ('285', '286')]
P = {i: D.fn_para(fid=i) for i in ids}
for i in ids:
    D.replace(P[i], OLD1, '', tag='221')
    D.replace(P[i], OLD2, '5.1.2-jadval', tag='221')
print('snoskalar:', ids)
D.save(F)
print(part.partname, pardoz(F))
