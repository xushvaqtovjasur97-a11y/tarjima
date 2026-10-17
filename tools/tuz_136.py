"""1.3.6-rasm: stress-test gorizonti (2–3 yil) va teskari aloqa (makroprudensial stress-testlarda) aniqlashtirildi."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
from docx.oxml.ns import qn
import kamlib as K
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
cap = D.para('1.3.6-rasm. Nazorat stress-testining umumiy arxitekturasi')
fid = [r.get(qn('w:id')) for r in cap._p.iter(qn('w:footnoteReference'))][0]
drw = cap._p.getprevious()
REP = {'prognozi (9 chorak)': 'prognozi (2–3 yil)',
       'Teskari aloqa: kredit taklifining iqtisodiyotga ta’siri': 'Teskari aloqa (makroprudensial stress-testlarda): kredit taklifining iqtisodiyotga ta’siri'}
n = 0
for t in drw.iter():
    if t.tag.endswith('}t') and t.text in REP:
        t.text = REP[t.text]; n += 1
print('almashtirildi:', n)
assert n >= 2
D.insert(D.fn_para(fid), 'P. 195–196', '; European Banking Authority. 2025 EU-wide stress test: Results. – Paris: EBA, 1 August 2025', tag='136')
D.save(F)
print(pardoz(F))
