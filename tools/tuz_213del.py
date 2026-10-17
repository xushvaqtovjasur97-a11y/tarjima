"""2.1.3-rasm (kapital monandligi) 2.2.2-rasm bilan takror bo‘lgani uchun olib tashlanadi; 2.1.4, 2.1.5-rasmlar qayta raqamlanadi (qizil)."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
from docx.oxml.ns import qn
import kamlib as K
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
cap = D.para('2.1.3-rasm. O‘zbekiston bank tizimida kapital monandligi')
chp = cap._p.getprevious()
assert list(chp.iter('{http://schemas.openxmlformats.org/drawingml/2006/chart}chart'))
txt = D.para('Amaldagi baholash tizimi natijalariga ko‘ra')
c4 = D.para('2.1.4-rasm. Salbiy stress-ssenariyda'); t4 = D.para('Shu bilan birga, stress-test natijalari')
c5 = D.para('2.1.5-rasm. Banklar moliyaviy barqarorligini'); t5 = D.para('Jadvalda keltirilgan tafovutlarni tahlil qilish')
# o‘chiriladigan sarlavha snoskalari
fids = [r.get(qn('w:id')) for r in cap._p.iter(qn('w:footnoteReference'))]
for f in list(D.fx.findall(qn('w:footnote'))):
    if f.get(qn('w:id')) in fids:
        D.fx.remove(f)
chp.getparent().remove(chp)
D.delete_para(cap, tag='213DEL')
D.replace(txt, '(2.1.3-rasm)', '(batafsil tahlil 2.2-paragrafda keltirilgan)', tag='213DEL')
D.replace(c4, '2.1.4-rasm', '2.1.3-rasm', tag='213DEL'); D.replace(t4, '2.1.4-rasm', '2.1.3-rasm', tag='213DEL')
D.replace(c5, '2.1.5-rasm', '2.1.4-rasm', tag='213DEL'); D.replace(t5, '2.1.5-rasm', '2.1.4-rasm', tag='213DEL')
D.save(F)
print('snoskalar o‘chirildi:', fids, pardoz(F))
