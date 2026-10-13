"""1.3.1-jadval: BCBS WP4 bo‘yicha aniqlashtirishlar va snoskaga birlamchi manba (qizil)."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
from docx.oxml.ns import qn
import kamlib as K
import namuna_131j as N
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
title = D.para('Rivojlangan mamlakatlarda banklarni baholash tizimlarining qiyosiy tavsifi')
fid = [r.get(qn('w:id')) for r in title._p.iter(qn('w:footnoteReference'))][0]
fn = D.fn_para(fid=fid)
assert ''.join(x.text or '' for x in fn._p.iter(qn('w:t'))).strip().startswith('Tadqiqot natijasida muallif tomonidan quyidagi manbalar asosida ishlab chiqildi: Karaliyev')
del title
N.apply(D)
D.insert(fn, 'ishlab chiqildi: ', N.BCBS + '; ', tag='131J')
D.save(F)
print(pardoz(F))
