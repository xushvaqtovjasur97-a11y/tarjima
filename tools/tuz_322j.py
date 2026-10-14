"""3.2.2-jadval izohi va 428-izoh; 3.2.1-rasm olib tashlandi; tahlil jadvalga bog‘landi; rasmlar 3.2.2→3.2.1, 3.2.3→3.2.2 (qizil)."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
from docx.oxml.ns import qn
import kamlib as K
import namuna_322j as N
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
TT = 'Integral indeks bloklarining tayanch va empirik vaznlari'
# 1) jadval izohi
D.insert(D.para('Izoh: * – logit usulida baholanmagan, tayanch vazni saqlangan.'), N.IZ_ANC, N.IZ_INS, tag='322j')
# 2) jadval nomidagi izoh (birinchi uchrashuv – jadval nomi)
tcap = [p for p in D.d.paragraphs if p.text.strip() == TT][0]
fid = [r.get(qn('w:id')) for r in tcap._p.iter(qn('w:footnoteReference'))][0]
D.insert(D.fn_para(fid=fid), '', N.FN_INS, tag='322j')
# 3) jadval tahlili
D.replace(D.para(N.TXT[0][0]), 'Eng katta farq likvidlik va daromadlilik bloklarida kuzatiladi:',
          'Likvidlik va daromadlilik bloklarida ham sezilarli farq kuzatiladi:', tag='322j')
D.replace(D.para(N.TXT[1][0].strip()), N.TXT[1][0], '', tag='322j')
# 4) 3.2.1-rasmni olib tashlash
fcap = [p for p in D.d.paragraphs if p.text.startswith('3.2.1-rasm. ' + TT)][0]
chp = fcap._p.getprevious()
rid = [e.get(qn('r:id')) for e in chp.iter('{http://schemas.openxmlformats.org/drawingml/2006/chart}chart')][0]
ffid = [r.get(qn('w:id')) for r in fcap._p.iter(qn('w:footnoteReference'))]
for f in D.fx.findall(qn('w:footnote')):
    if f.get(qn('w:id')) in ffid:
        D.fx.remove(f)
chp.getparent().remove(chp)
D.delete_para(fcap, tag='322j')
D.d.part.drop_rel(rid)
# 5) rasmdan keyingi xatboshi jadvalga bog‘lanadi
p5 = D.para('Rasmdan ko‘rinib turibdiki, usullar o‘rtasidagi farq asosan likvidlik')
D.replace(p5, 'Rasmdan', 'Jadval ma’lumotlaridan', tag='322j')
D.insert(p5, 'usullar o‘rtasidagi farq asosan ', 'kapital, ', tag='322j')
# 6) rasmlar raqami (avval 3.2.2, keyin 3.2.3)
D.replace(D.para('tanlangan holatda (3.2.2-rasm)'), '(3.2.2-rasm)', '(3.2.1-rasm)', tag='322j')
D.replace(D.para('3.2.2-rasm. Monte-Karlo'), '3.2.2-rasm', '3.2.1-rasm', tag='322j')
D.replace(D.para('dinamikasini baholash uchun qo‘llandi (3.2.3-rasm)'), '(3.2.3-rasm)', '(3.2.2-rasm)', tag='322j')
D.replace(D.para('3.2.3-rasm. Tanlangan banklarda'), '3.2.3-rasm', '3.2.2-rasm', tag='322j')
D.replace(D.para('(3.2.3-rasmda uzuq chiziq'), '(3.2.3-rasmda', '(3.2.2-rasmda', tag='322j')
D.save(F)
print(pardoz(F))
