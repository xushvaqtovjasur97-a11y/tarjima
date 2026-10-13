"""3.1.1-jadval: 2- va 5-qatorlardagi dalillar (qizil)."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
import docx
import kamlib as K
import namuna_311j as N
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
cap = [p for p in D.d.paragraphs if p.text.startswith('Amaldagi baholash tizimi muammolari')][0]
tbl = docx.table.Table(cap._p.getnext(), D.d)
for i, (old, new) in N.FIX.items():
    c = tbl.cell(i, 1)
    assert c.text == old, (i, c.text)
    if i == 1:
        D.cell_replace(tbl, i, 1, '1,65–3,30', '1,65–4,90', tag='311j')
    else:
        D.cell_replace(tbl, i, 1, old[len('2.1-§: '):], new[len('2.1-§: '):], tag='311j')
D.save(F)
print(pardoz(F))
