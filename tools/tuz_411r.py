"""4.1.1-rasm: to‘liq empirik ROC egri chiziqlari, MBII-8 chegaralari egri chiziq ustida, legenda bo‘sh burchakda
(ma’lumot va diagramma: tools/namuna_411r.py); keyingi xatboshida 10 % dagi aniqlik 73 foiz (ko‘k)."""
import re, sys
sys.path.insert(0, '/home/user/tarjima/tools')
from docx.oxml.ns import qn
import kamlib as K
import wordart as W
import namuna_411r as N
from pardoz import pardoz

K.RED = '0070C0'
F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
cap = D.para('4.1.1-rasm. Integral indeks, muammoli kreditlar ulushi va Z-score')
chp = cap._p.getprevious()
rid = [e.get(qn('r:id')) for e in chp.iter('{http://schemas.openxmlformats.org/drawingml/2006/chart}chart')][0]
part = D.d.part.related_parts[rid]
xrel = [r for r in part.rels.values() if r.reltype.endswith('/package')][0]
xml, rows = N.chart()
part._blob = re.sub(r'(<c:externalData r:id=")[^"]+(")', r'\g<1>' + xrel.rId + r'\g<2>', xml).encode('utf-8')
xrel.target_part._blob = W._xlsx(rows)
for ext in chp.iter('{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent'):
    ext.set('cx', str(int(N.W_CM * 360000))); ext.set('cy', str(int(N.H_CM * 360000)))

p = D.para('Rasmdan ko‘rinib turibdiki, integral indeks va muammoli kreditlar ulushining ROC egri chiziqlari')
old = 'integral indeks muammo hodisalarining qariyb 70 foizini'
assert p.text.count(old) == 1
D.replace(p, 'qariyb 70 foizini', '73 foizini')
D.save(F)
print(pardoz(F))
