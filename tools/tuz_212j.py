"""2.1.2-jadval: aktivlar sifati holati (restrukturizatsiya) va sanatsiya muddati (qizil)."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
import kamlib as K
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
t = [tb for tb in D.d.tables if tb.rows[0].cells[0].text.strip() == 'Soha'][0]
R = {r.cells[0].text.strip(): i for i, r in enumerate(t.rows)}
c = t.rows[R['Aktivlar sifati']].cells[2].paragraphs[0]
D.replace(c, 'ayniqsa imtiyozli va yo‘naltirilgan kreditlarda', 'ayniqsa restrukturizatsiya qilingan kreditlarda (6 oygacha imtiyozli davr tasnifga ta’sir qilmaydi, sinov muddati yo‘q)', tag='212J')
c = t.rows[R['Banklarni sanatsiya qilish']].cells[4].paragraphs[0]
D.replace(c, '2026–2028-y.', '2025–2028-y.', tag='212J')
D.save(F)
print(pardoz(F))
for i in (R['Aktivlar sifati'], R['Banklarni sanatsiya qilish']):
    print(' | '.join(x.text for x in K.Doc(F).d.tables[[k for k, tb in enumerate(K.Doc(F).d.tables) if tb.rows[0].cells[0].text.strip() == 'Soha'][0]].rows[i].cells))
