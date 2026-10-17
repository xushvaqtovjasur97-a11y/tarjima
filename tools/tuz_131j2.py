"""1.3.1-jadval: CAMEL yillari (1978–1979) va SAABA turi (statistik-ekspert) aniqlashtirildi (qizil)."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
import kamlib as K
import namuna_131j as N
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
t = [tb for tb in D.d.tables if tb.rows[0].cells[0].text.strip() == 'Mamlakat' and any('CAMEL (1979)' in r.cells[1].text for r in tb.rows)][0]
R = {r.cells[1].text.strip(): i for i, r in enumerate(t.rows)}
D.cell_replace(t, R['CAMEL (1979), CAMELS (1997)'], 1, 'CAMEL (1979)', 'CAMEL (1978–1979)', tag='131J2')
D.cell_replace(t, R['ORAP (1997), SAABA (1997)'], 2, 'Reyting va statistik model', 'Reyting va statistik-ekspert model', tag='131J2')
D.save(F)
print(pardoz(F))
for r in t.rows[1:2] + t.rows[5:6]:
    pass
