"""1.3.1-rasm: BAITE olib tashlandi, FIMS = SEER, tasnif manbasi; matn va snoska aniqlashtirildi (qizil)."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
import kamlib as K
import namuna_131 as N
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
txt = D.para(N.OLD_T)
fn = D.fn_para(contains='Tadqiqot natijasida muallif tomonidan quyidagi manbalar asosida ishlab chiqildi: Устойчивость банковской системы и развитие банковской политики: монография / под ред. О.И. Лаврушина. – М.: КНОРУС, 2014. – С. 56–57; Karaliyev')
cap = [p for p in D.d.paragraphs if p.text.startswith('1.3.1-rasm')][0]
print('rasm:', N.fix(cap._p.getprevious()))
del cap
D.replace(txt, N.OLD_T, N.NEW_T, tag='131')
D.insert(fn, 'ishlab chiqildi: ', N.BCBS + '; ', tag='131')
D.save(F)
print(pardoz(F))

# 2-bosqich: jadvallardagi “BAITE” (mavjud bo‘lmagan tizim) olib tashlanadi
if __name__ == '__main__' and 'baite' in sys.argv:
    D = K.Doc(F)
    for t in D.d.tables:
        for ri, r in enumerate(t.rows):
            for ci, c in enumerate(r.cells):
                if 'BAITE (1997), RATE (1998)' in c.text:
                    D.cell_replace(t, ri, ci, 'BAITE (1997), RATE (1998)', 'RATE (1998)', tag='131')
                elif 'BAITE, RATE' in c.text:
                    D.cell_replace(t, ri, ci, 'BAITE, RATE', 'RATE', tag='131')
    D.save(F)
    print('baite:', pardoz(F))
