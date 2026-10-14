"""Kirish: asosiy matn hajmi 242 → 243 bet (qizil), bog'lovchi tahrirlardan keyin."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
import kamlib as K
from pardoz import pardoz
F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
D.replace(D.para('Dissertatsiyaning asosiy matni 242 betni'), '242', '243', tag='hajm')
D.save(F)
print(pardoz(F))
