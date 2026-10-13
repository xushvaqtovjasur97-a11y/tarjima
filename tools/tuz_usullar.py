"""Kirish, tadqiqot usullari: analogiya (CAMEL vaznlariga tayanish) va statistik verifikatsiya (qizil)."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
import kamlib as K
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
p = [q for q in D.d.paragraphs if q.text.startswith('Tadqiqotning usullari.')][0]
D.insert(p, 'indeks usuli,', ' analogiya (CAMEL tizimi vaznlariga tayanish), statistik verifikatsiya,', tag='usullar')
D.save(F)
print(p.text, pardoz(F))
