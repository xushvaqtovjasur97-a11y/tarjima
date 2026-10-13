"""“tayanch vaznlari…” → “tayanch vaznlar…” (sifatlovchi bilan egalik qo‘shimchasi kerak emas)."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn
import kamlib as K
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
n = 0
for pe in list(D.d.element.body.iter(qn('w:p'))):
    p = Paragraph(pe, D.d)
    for old, new in [('tayanch vaznlari', 'tayanch vaznlar'), ('Tayanch vaznlari', 'Tayanch vaznlar')]:
        while old in p.text:
            D.replace(p, old, new, tag='tayanch2'); n += 1
D.save(F)
print(n, pardoz(F))
