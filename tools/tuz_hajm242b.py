"""Kirish: asosiy matn hajmi 243 → 242 bet (qizil), bog'lovchi tahrirlardan keyin."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
import kamlib as K
from pardoz import pardoz
F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
D.replace(D.para('Dissertatsiyaning asosiy matni 243 betni'), '243', '242', tag='hajm')
D.save(F)
print(pardoz(F))
