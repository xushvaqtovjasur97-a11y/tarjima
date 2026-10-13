"""2.2.2-rasm chiziqli diagramma (13% minimum bilan) ga almashtiriladi; 2.1-§ va 2.2-§ dagi tegishli matnlar qayta ishlanadi (qizil)."""
import re, sys
sys.path.insert(0, '/home/user/tarjima/tools')
from docx.oxml.ns import qn
import kamlib as K
import wordart as W
import namuna_222l as N
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
p21 = D.para('Amaldagi baholash tizimi natijalariga ko‘ra')
p22a = D.para('Bu risklarni kapital yetarliligidan boshlab ko‘rib chiqamiz.')
cap = D.para('2.2.2-rasm. O‘zbekiston bank tizimida kapital yetarliligi')
p22b = D.para('Xalqaro mezonlar nuqtai nazaridan bu ko‘rsatkichlar yuqori')
fn = D.fn_para(contains='Tadqiqot natijasida muallif tomonidan 2.2.1-jadval ma’lumotlari asosida ishlab chiqildi.')

# 1) diagramma
chp = cap._p.getprevious()
rid = [e.get(qn('r:id')) for e in chp.iter('{http://schemas.openxmlformats.org/drawingml/2006/chart}chart')][0]
part = D.d.part.related_parts[rid]
xrel = [r for r in part.rels.values() if r.reltype.endswith('/package')][0]
xml, rows = N.chart()
part._blob = re.sub(r'(<c:externalData r:id=")[^"]+(")', r'\g<1>' + xrel.rId + r'\g<2>', xml).encode('utf-8')
xrel.target_part._blob = W._xlsx(rows)
for ext in chp.iter('{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent'):
    ext.set('cy', str(int(8.5 * 360000)))
del chp

# 2) 2.1-§: raqamlar 2.2-§ ni takrorlamasligi uchun qisqa umumiy xulosa
D.replace(p21, 'O‘zbekiston bank tizimi kapital yetarliligi bo‘yicha barqaror holatda. Kapital monandligi 2021–2025-yillarda 17,4–18,3 foiz oralig‘ida bo‘lib, 2023-yil oxirida 17,5 foizni',
          'O‘zbekiston bank tizimi kapital yetarliligi bo‘yicha barqaror holatda: 2021–2025-yillarda kapital monandligi 13 foizlik regulyativ minimumdan barqaror yuqori bo‘lgan', tag='222')
D.replace(p21, ', 2025-yil oxirida 18,3 foizni tashkil etgan; asosiy kapital (CET1) monandligi 2025-yil oxirida 14,7 foizga teng bo‘lgan',
          ', 2025-yil oxirida esa 18,3 foizga yetgan', tag='222')
D.replace(p21, 'Muammoli kreditlar ulushi esa', 'Rasmiy muammoli kreditlar ulushi ham', tag='222')
D.replace(p21, ' (batafsil tahlil 2.2-paragrafda keltirilgan).',
          '. Ushbu ko‘rsatkichlar dinamikasi va ularning xalqaro mezonlar bilan qiyosi 2.2-paragrafda batafsil tahlil qilinadi.', tag='222')

# 3) 2.2-§: rasm tahlili va stress-test takrorini havola bilan almashtirish
D.replace(p22b, 'Xalqaro mezonlar nuqtai nazaridan',
          'Rasmdan ko‘rinadiki, regulyativ kapital monandligi 13 foizlik minimumdan har yili 4,4–5,3 foiz punkt yuqori bo‘lgan va bu zaxira 2025-yilda eng katta qiymatga yetgan, I darajali asosiy kapital monandligi esa 2023-yilgacha 14,1 foizgacha pasayib, keyin 14,7 foizgacha tiklangan. Xalqaro mezonlar nuqtai nazaridan', tag='222')
D.replace(p22b, 'Uchinchidan, Markaziy bankning stress-testlari', 'Uchinchidan, 2.1-paragrafda keltirilganidek, Markaziy bankning stress-testlari', tag='222')
D.replace(p22b, 'pasayishi mumkinligini ko‘rsatgan', 'pasayishi mumkinligini ko‘rsatgan (2.1.3-rasm)', tag='222')

# 4) rasm manbasi: 13% minimum manbasi
D.replace(fn, '2.2.1-jadval ma’lumotlari asosida',
          '2.2.1-jadval ma’lumotlari va Markaziy bankning “Tijorat banklari kapitalining monandligiga qo‘yiladigan talablar to‘g‘risida”gi nizomi (ro‘yxat raqami 2693) asosida', tag='222')
D.save(F)
print(part.partname, pardoz(F))
