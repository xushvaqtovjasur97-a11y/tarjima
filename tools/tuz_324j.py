"""3.2.4-jadval nomi, izohi; keyingi xatboshida tayanch vaznli indeks (qizil)."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
import kamlib as K
import namuna_324j as N
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
cap = D.para(N.CAP_FULL + N.CAP_OLD)
D.replace(cap, N.CAP_OLD, N.CAP_NEW, tag='324j')
t = D.table_after(N.CAP_FULL)
izsrc = D.para('Izoh: “Aloqabank” ATB empirik vaznlarni baholashda qatnashmagan')
iz = D.new_para_after(izsrc, N.IZOH, tag='324j')
t._tbl.addnext(iz._p)
D.replace(D.para(N.OLD3), 'empirik asoslangan indeks', 'empirik tekshiruvdan o‘tgan, tayanch vaznli indeks', tag='324j')
D.save(F)
print(pardoz(F))
