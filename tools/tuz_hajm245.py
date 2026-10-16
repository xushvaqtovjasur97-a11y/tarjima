"""Kirish: asosiy matn hajmi 244 → 245 bet (NSFR/LCR tuzatishlaridan keyin, adabiyotlar 249-betdan)."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
import kamlib as K
from pardoz import pardoz
F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
D.replace(D.para('Dissertatsiyaning asosiy matni 244 betni'), '244', '245', tag='hajm')
D.save(F)
print(pardoz(F))
