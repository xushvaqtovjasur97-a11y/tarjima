"""3.2.1-jadval: entropiya va logit qatorlari aniqlashtirildi; 3.3/4.1 ga havola; 0,31 → 0,30 (qizil)."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
import kamlib as K
import namuna_321j as N
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)


def minimal(p, old, new):
    """faqat o'zgargan qismni qizil qiladi"""
    a = 0
    while a < min(len(old), len(new)) and old[a] == new[a]:
        a += 1
    b = 0
    while b < min(len(old), len(new)) - a and old[-1 - b] == new[-1 - b]:
        b += 1
    mo, mn = old[a:len(old) - b], new[a:len(new) - b]
    pos = p.text.find(old) + a
    if mo == '':
        anchor = p.text[max(0, pos - 25):pos]
        D.insert(p, anchor, mn, occ=p.text[:pos].count(anchor) - 1, tag='321j')
    else:
        occ = p.text[:pos].count(mo)
        D.replace(p, mo, mn, occ=occ, tag='321j')


t = D.table_after('Integral indeks vaznlarini aniqlashning empirik usullari')
for r, c, old, new in N.CELLS:
    p = [pp for pp in t.rows[r].cells[c].paragraphs if old in pp.text][0]
    minimal(p, old, new)
D.insert(D.para(N.ANC_33), N.ANC_33, N.INS_33, tag='321j')
D.replace(D.para(N.OLD_T), '(0,31)', '(0,30)', tag='321j')
D.save(F)
print(pardoz(F))
