"""Ikki nuqta bilan ulanib qolgan mustaqil gaplarni mantiqan bog‘lash (ko‘k rangda)."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
import kamlib as K
from pardoz import pardoz

K.RED = '0070C0'          # o‘zgarishlar ko‘k rangda
F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
E = [
    ('ham davom etdi: Yevropa Ittifoqida', 'ham davom etdi. Jumladan, Yevropa Ittifoqida'),
    ('ishonch egallaydi: muallif fikricha,', 'ishonch egallaydi. Muallif fikricha,'),
    ('birinchi o‘ringa chiqadi: muallif moliyaviy', 'birinchi o‘ringa chiqadi. Muallif moliyaviy'),
    ('bankning o‘zi bilan cheklanmaydi: O.N. Afanasyeva', 'bankning o‘zi bilan cheklanmaydi, u butun tizimga tarqalishi mumkin. Shu sababli O.N. Afanasyeva'),
    ('yagona bahoga keltirmaydi: masalan, B.T.', 'yagona bahoga keltirmaydi. Masalan, B.T.'),
    ('ishonchliroq qiladi: Sokolinskaya prognoziga', 'ishonchliroq qiladi. Masalan, N.E. Sokolinskaya prognoziga'),
    ('axborot manbai bo‘ladi: muallif keltirgan misolda', 'axborot manbai bo‘ladi. Jumladan, N.E. Sokolinskaya keltirgan misolda'),
    ('sababi kapital manbalarida: bankning', 'sababi kapital manbalari bilan bog‘liq. Xususan, bankning'),
    ('AQSh tajribasini eslatadi: J.M. Hendrickson', 'AQSh tajribasini eslatadi. Unda J.M. Hendrickson'),
    ('Bu holat tasodifiy emas: 2021-yildan', 'Bu holat tasodifiy emas, chunki 2021-yildan'),
    ('prognoz davri uchun ahamiyatli: bunda bankning', 'prognoz davri uchun ahamiyatli. Bunday davrlarda bankning'),
    ('asosida baholanadi: P. Athanasoglou', 'asosida baholanadi. Jumladan, P. Athanasoglou'),
    ('bilan izohlanadi: G. Iannotta', 'bilan izohlanadi. Xususan, G. Iannotta'),
    ('tadqiqotlarga asoslanadi: D. Louzis', 'tadqiqotlarga asoslanadi. Jumladan, D. Louzis'),
    ('sabab bo‘lgan: II bobda ko‘rsatilganidek,', 'sabab bo‘lgan. II bobda ko‘rsatilganidek,'),
    ('zararlaridir: bank faoliyatining', 'zararlaridir. O‘sha paytda bank faoliyatining'),
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
    D.replace(p, mo, mn, occ=p.text[:pos].count(mo), tag='in')


D = K.Doc(F)
for old, new in E:
    ps = [p for p in D.d.paragraphs if old in p.text]
    assert len(ps) == 1, (old, len(ps))
    minimal(D, ps[0], old, new)
D.save(F)
print(len(E), pardoz(F))
