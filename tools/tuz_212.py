"""2.1.2-rasm (NSFR – monitoring; oddiy rangda), snoska va 2.1.1-jadval NSFR qatori (qizil)."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
from docx.oxml.ns import qn
import kamlib as K
import namuna_212 as N
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
cap = D.para('2.1.2-rasm. O‘zbekistonda banklar moliyaviy barqarorligini baholashning amaldagi tizimi')
t = [tb for tb in D.d.tables if tb.rows[0].cells[0].text.strip() == 'Ko‘rsatkich' and any(r.cells[0].text.strip() == 'NSFR' for r in tb.rows)][0]
ri = [i for i, r in enumerate(t.rows) if r.cells[0].text.strip() == 'NSFR'][0]
fid = [r.get(qn('w:id')) for r in cap._p.iter(qn('w:footnoteReference'))][0]
print('chizma:', N.fix(cap._p.getprevious()))
D.replace(D.fn_para(fid), 'Muallif ishlanmasi.', N.NEW_FN, tag='212')
D.set_cell(t, ri, 1, N.T_NEW[1], tag='212')
D.set_cell(t, ri, 3, N.T_NEW[3], tag='212')
D.save(F)
print(pardoz(F))
