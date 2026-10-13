"""Muassasa nomi: “Denov tadbirkorlik va pedagogika universiteti” → “instituti” (qizil)."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
import kamlib as K
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
D.replace(D.para('DENOV TADBIRKORLIK VA PEDAGOGIKA UNIVERSITETI'), 'UNIVERSITETI', 'INSTITUTI', tag='inst')
D.replace(D.para('Denov tadbirkorlik va pedagogika universiteti'), 'pedagogika universiteti', 'pedagogika instituti', tag='inst')
D.save(F)
print(pardoz(F))
