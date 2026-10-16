"""4.2.6-jadval (davlat banklari va boshqa banklar): birinchi ikki ustun kengaytirildi – qatorlar bir qatorga sig‘adi,
jadval sahifalar orasida bo‘linmaydi (sonli ustunlar torayadi, katak chekinishi 0,1 sm)."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
from docx.oxml.ns import qn
from docx.shared import Cm
import kamlib as K
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
t = D.table_after('Davlat banklari va boshqa banklar moliyaviy barqarorlik ko‘rsatkichlarining ssenariylar')
W_CM = [5.5, 2.9, 1.9] + [1.55] * 4
# katak ichki chekinishi 0,1 sm (sarlavhalardagi so‘zlar bo‘linmasligi uchun)
tp = t._tbl.tblPr
mar = tp.find(qn('w:tblCellMar'))
if mar is None:
    mar = tp.makeelement(qn('w:tblCellMar'), {})
    look = tp.find(qn('w:tblLook'))
    (look.addprevious(mar) if look is not None else tp.append(mar))   # sxema tartibi: tblLayout, tblCellMar, tblLook
for side in ('left', 'right'):
    e = mar.find(qn('w:' + side))
    if e is None:
        e = mar.makeelement(qn('w:' + side), {}); mar.append(e)
    e.set(qn('w:w'), '57'); e.set(qn('w:type'), 'dxa')
for gc, w in zip(t._tbl.tblGrid.findall(qn('w:gridCol')), W_CM):
    gc.set(qn('w:w'), str(int(Cm(w).twips)))
for r in t.rows:
    for j, c in enumerate(r.cells):
        c.width = Cm(W_CM[j])
D.save(F)
print(pardoz(F))
