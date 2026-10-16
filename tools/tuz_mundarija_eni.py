"""Mundarija bandlari eni bo‘yicha (justify) tekislanadi: “MUNDARIJA” sarlavhasidan keyingi birinchi Heading 1 gacha."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
d = docx.Document(F)
P = d.paragraphs
i0 = next(i for i, p in enumerate(P) if p.text.strip() == 'MUNDARIJA')
n = 0
for p in P[i0 + 1:]:
    if p.style.name.startswith('Heading'):
        break
    if p.text.strip():
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY; n += 1
d.save(F)
print('mundarija bandlari:', n, pardoz(F))
