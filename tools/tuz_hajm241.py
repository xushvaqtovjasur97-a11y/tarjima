"""Kirish: asosiy matn hajmi 238 → 241 bet (qizil)."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
import kamlib as K
from pardoz import pardoz
F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
D.replace(D.para('Dissertatsiyaning asosiy matni 238 betni'), '238', '241', tag='hajm')
D.save(F)
print(pardoz(F))
