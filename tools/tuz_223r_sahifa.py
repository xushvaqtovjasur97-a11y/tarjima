"""2.2.3-rasm kiritilgandan keyingi sahifalash (tuz_223r.py dan keyin):
- “2.2.4-rasm ma’lumotlarini…” xatboshisi yana rasmdan keyinga qaytarildi (2.2.3-rasm ixchamlashgani sababli joy yetadi);
- 2.2.2–2.2.4-jadvallar izohi 2 qatorga qisqartirildi (izoh jadval bilan bir sahifada turishi uchun);
- jadvallardan oldingi xatboshi ixchamlashtirildi (mazmun saqlangan). O‘zgarishlar ko‘k rangda.
"""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
import kamlib as K
from pardoz import pardoz

K.RED = '0070C0'
F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)

p = D.para('2.2.4-rasm ma’lumotlarini xalqaro mezonlar bilan solishtiradigan')
cap = D.para('2.2.4-rasm. O‘zbekiston bank tizimida aktivlar va kapital rentabelligi')
if p._p.getnext() is not None and p._p.getnext().find('.//' + qn('w:drawing')) is not None:
    p._p.getparent().remove(p._p)
    cap._p.addnext(p._p)

for bank in ('“O‘zsanoatqurilishbank”', '“Mikrokreditbank”', '“Aloqabank”'):
    t = D.table_after(bank + ' ATB ning moliyaviy barqarorlik ko‘rsatkichlari')
    iz = Paragraph(t._tbl.getnext(), t._parent)
    assert iz.text.startswith('Izoh: muammoli kreditlar'), iz.text[:40]
    D.replace(iz, 'umidsiz deb tasniflangan kreditlar', 'umidsiz kreditlar')
    D.replace(iz, 'ROA va ROE sof foydaning aktivlar va kapitalning o‘rtacha yillik qiymatiga nisbati sifatida hisoblandi.',
              'ROA va ROE – sof foydaning o‘rtacha yillik aktivlar va kapitalga nisbati.')

q = D.para('Tizim darajasidagi agregat ko‘rsatkichlar alohida banklar o‘rtasidagi farqlarni yashiradi')
for old, new in [('Shu sababli tahlilni uchta', 'Shu sababli tahlil uchta'),
                 ('misolida chuqurlashtirish maqsadga muvofiq.', 'misolida chuqurlashtiriladi.'),
                 ('bankning 2021–2025-yillardagi asosiy', 'bankning asosiy'),
                 ('so‘ngra ular o‘zaro qiyosiy', 'so‘ngra ular qiyosiy')]:
    assert q.text.count(old) == 1, old
    D.replace(q, old, new)
D.save(F)
print(pardoz(F))
