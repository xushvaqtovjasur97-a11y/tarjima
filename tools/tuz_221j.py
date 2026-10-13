"""2.2.1-jadval: sarlavhadan “yil oxiriga” olib tashlandi; 2021-yil ROA 1,3; 2024-yil zaxiralar 45,4 va depozitlar/kreditlar 53,5 (qizil);
matn va 2.2.4-rasmda ROA 2021 = 1,3."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
from docx.oxml.ns import qn
import kamlib as K
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
cap = D.para('O‘zbekiston bank tizimining moliyaviy barqarorlik ko‘rsatkichlari, yil oxiriga, %')
txt = D.para('Soliq to‘lagunga qadar aktivlar rentabelligi 2021-yilda 1,4 foiz')
t = [tb for tb in D.d.tables if tb.rows[0].cells[0].text.strip() == 'Ko‘rsatkich' and any('Aktivlar rentabelligi (soliq' in r.cells[0].text for r in tb.rows)][0]
R = {r.cells[0].text.strip(): i for i, r in enumerate(t.rows)}
cap4 = D.para('2.2.4-rasm.')
rid = [e.get(qn('r:id')) for e in cap4._p.getprevious().iter('{http://schemas.openxmlformats.org/drawingml/2006/chart}chart')][0]
ch = D.d.part.related_parts[rid].partname
del cap4

D.replace(cap, ', yil oxiriga', '', tag='221J')
D.cell_replace(t, R['Aktivlar rentabelligi (soliq to‘lagunga qadar)'], 1, '1,4', '1,3', tag='221J')
D.cell_replace(t, R['Maxsus zaxiralarning muammoli kreditlarga nisbati'], 4, '45,5', '45,4', tag='221J')
D.cell_replace(t, R['Mijozlar depozitlarining jami kreditlarga nisbati'], 4, '53,6', '53,5', tag='221J')
D.replace(txt, '2021-yilda 1,4 foiz', '2021-yilda 1,3 foiz', tag='221J')
D.chart_set(ch, 0, [1.3], tag='221J')
K.chart_sync_xlsx(D, ch)
D.save(F)
print(ch, pardoz(F))
