"""Nuqtali vergul bilan ulanib qolgan mustaqil gaplarni mantiqan bog‘lash (ko‘k rangda)."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
import kamlib as K
from pardoz import pardoz

K.RED = '0070C0'          # o‘zgarishlar ko‘k rangda
F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
E = [
    ('umumiy bahosidir; 3, 4', 'umumiy bahosidir. Bunda 3, 4'),
    ('e’lon qilindi; ular faqat', 'e’lon qilindi. Bu takliflar faqat'),
    ('amalga kirdi; unda ichki', 'amalga kirdi. Unda ichki'),
    ('zarur bo‘lgan; yuqori sifatli', 'zarur bo‘lgan. Bundan tashqari, yuqori sifatli'),
    ('qilib belgilandi; 2024-yil dekabrida', 'qilib belgilandi. Biroq 2024-yil dekabrida'),
    ('hisobga oladi; uning minimal', 'hisobga oladi. Uning minimal'),
    ('monitoring qilingan; inqirozdan keyin esa', 'monitoring qilingan. Inqirozdan keyin esa'),
    ('saqlanib kelmoqda; ikkalasi qo‘shilganda', 'saqlanib kelmoqda. Ikkala toifa qo‘shilganda esa'),
    ('har yili aniqlanadi; 2026-yil uchun', 'har yili aniqlanadi. Xususan, 2026-yil uchun'),
    ('tavsiya qilgan; bunda kamsitib', 'tavsiya qilgan. Bunda kamsitib'),
    ('ega emas; ularni joriy etish', 'ega emas. Shu sababli ularni joriy etish'),
    ('sof foyda olgan; 2025-yilda ham kapital', 'sof foyda olgan. Xuddi shu holat 2025-yilda ham kuzatilgan: kapital'),
    ('past bo‘lgan; asosiy kapital monandligi', 'past bo‘lgan, asosiy kapital monandligi esa'),
    ('yuqori bo‘lgan; 2022–2024-yillarda', 'yuqori bo‘lgan. Ayniqsa, 2022–2024-yillarda'),
    ('ikkinchi o‘rinda turadi; uning yetakchiligi', 'ikkinchi o‘rinda turadi. Uning yetakchiligi esa'),
    ('panel indeksi keltirilgan; 13 ko‘rsatkichli', 'panel indeksi keltirilgan. To‘liq, 13 ko‘rsatkichli'),
]


def minimal(D, p, old, new):
    a = 0
    while a < min(len(old), len(new)) and old[a] == new[a]:
        a += 1
    b = 0
    while b < min(len(old), len(new)) - a and old[-1 - b] == new[-1 - b]:
        b += 1
    mo, mn = old[a:len(old) - b], new[a:len(new) - b]
    pos = p.text.find(old) + a
    D.replace(p, mo, mn, occ=p.text[:pos].count(mo), tag='nv')


D = K.Doc(F)
for old, new in E:
    ps = [p for p in D.d.paragraphs if old in p.text]
    assert len(ps) == 1, (old, len(ps))
    minimal(D, ps[0], old, new)
D.save(F)
print(len(E), pardoz(F))
