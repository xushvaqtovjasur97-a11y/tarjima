"""MBII-13 / MBII-8 belgilari 3.2-§ da kiritildi, 3.3-§ da havola, sinf chegaralari va 3.2.4 izohida (qizil)."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
import kamlib as K
import namuna_mbii8 as N
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
p = D.para(N.ANC1)
D.replace(p, '), unda kapital bloki', ')' + N.INS1 + ' kapital bloki', tag='m8')
p2 = D.para(N.OLD2)
D.replace(p2, N.OLD2, N.NEW2, tag='m8')
D.insert(D.para(N.OLD3), N.OLD3, ' MBII-8 bo‘yicha', tag='m8')
D.insert(D.para(N.OLD4), '8 ko‘rsatkichli choraklik indeks', ' (MBII-8)', tag='m8')
D.save(F)
print(pardoz(F))
