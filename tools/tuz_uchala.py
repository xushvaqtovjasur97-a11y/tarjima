"""2.3-§: yashirin kredit riski koeffitsiyenti bo‘yicha “ikkala bank” → “uchala bank” (qizil)."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
import kamlib as K
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
R = [('koeffitsiyenti ikkala bankda ham tahlil davrining', 'ikkala bankda', 'uchala bankda'),
     ('Ikkala bank uchun ham koeffitsiyentning', 'Ikkala bank uchun', 'Uchala bank uchun'),
     ('koeffitsiyenti ikkala bankda ham 1,65–3,30', 'ikkala bankda ham 1,65–3,30', 'uchala bankda ham 1,65–4,90')]
for key, old, new in R:
    ps = [p for p in D.d.paragraphs if key in p.text]
    assert len(ps) == 1, key
    D.replace(ps[0], old, new, tag='uchala')
D.save(F)
print(pardoz(F))
