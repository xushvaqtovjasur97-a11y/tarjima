"""Yashirin kredit riski: (2.3) tuzatish formulasi, 2.3.2-jadvalga koeffitsiyent qatori, 4.3-§ da havola (qizil)."""
import copy, sys
sys.path.insert(0, '/home/user/tarjima/tools')
from docx.oxml.ns import qn
import kamlib as K
import namuna_ykr as N
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)


def after(anchor, text, pstyle):
    np_ = D.new_para_after(anchor, text, tag='ykr')
    old = np_._p.find(qn('w:pPr'))
    new = copy.deepcopy(pstyle._p.find(qn('w:pPr')))
    if old is not None:
        np_._p.remove(old)
    np_._p.insert(0, new)
    return np_


izoh = D.para(N.AFTER_P)                       # (2.2) dan keyingi "bu yerda ..." xatboshisi
matn = D.para('Ikkinchisi – dissertatsiyada taklif etilayotgan yashirin kredit riski')
formula = D.para('Kykr = (S3 / LMHXS)')
p1 = after(izoh, N.P1, matn)
p2 = after(p1, N.F3, formula)
after(p2, N.P2, izoh)
# 2.3.2-jadval: koeffitsiyent qatori
t = D.table_after('“O‘zsanoatqurilishbank” ATB ning asosiy ko‘rsatkichlari: milliy hisobot va MHXS')
D.add_row(t, len(t.rows) - 1, N.ROW, tag='ykr')
# 4.3-§
p43 = D.para(N.OLD43)
D.insert(p43, 'yashirin kredit riski koeffitsiyenti bilan ', '(2.3) formula bo‘yicha ',
         occ=p43.text[:p43.text.find(N.OLD43)].count('yashirin kredit riski koeffitsiyenti bilan '), tag='ykr')
D.save(F)
print(pardoz(F))
