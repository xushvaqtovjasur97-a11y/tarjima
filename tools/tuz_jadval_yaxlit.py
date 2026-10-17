"""Kichik jadvallarni bir sahifada yaxlit saqlash: oxirgi qatordan boshqa barcha qatorlar paragraflariga keepNext
(jadval sahifa oxirida bir-ikki qator bo‘lib bo‘linmasligi uchun). Foydalanish: python3 tuz_jadval_yaxlit.py "<jadval nomi boshi>" ..."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
import kamlib as K
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
for cap in sys.argv[1:]:
    t = D.table_after(cap)
    for r in list(t.rows)[:-1]:
        for c in r.cells:
            for p in c.paragraphs:
                p.paragraph_format.keep_with_next = True
    D.para(cap).paragraph_format.keep_with_next = True
    print('yaxlit:', cap[:60])
D.save(F)
print(pardoz(F))
