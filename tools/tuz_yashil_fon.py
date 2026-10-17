"""Barcha yashil fon (w:highlight green) belgilarini olib tashlash: matn oddiy qora bo‘ladi."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
import kamlib as K
from docx.oxml.ns import qn
from lxml import etree
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
n = 0
roots = [D.d.element]
if D.fx is not None:
    roots.append(D.fx)
for root in roots:
    for h in list(root.iter(qn('w:highlight'))):
        if h.get(qn('w:val')) == 'green':
            h.getparent().remove(h); n += 1
D.save(F)
print('olib tashlandi:', n, pardoz(F))
