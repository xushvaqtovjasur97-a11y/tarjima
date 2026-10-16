"""NSFR (regulyativ kapital – ikki marta hisoblash xatosi) va LCR (aniqliklar, “Mikrokreditbank” ATB) bo‘yicha tuzatishlar.
Asos: manbalar/hisob_kitob/nsfr_tekshiruv.py, NSFR_LCR_tekshiruv_va_tuzatishlar.docx. Kiritilgandan keyin tools/qora.py bilan qora qilinadi."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
import kamlib as K
from pardoz import pardoz

K.RED = '0070C0'
F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)


def tx(el):
    return ''.join(x.text or '' for x in el.iter(qn('w:t')))


def rep(p, old, new):
    assert tx(p._p).count(old) == 1, (old[:60], tx(p._p).count(old))
    D.replace(p, old, new)


# 1) Kirish – amaliy natijalar, 3-band
p = D.para('sof barqaror moliyalashtirish me’yori hisob-kitobida Bazel III standartida va Markaziy bankning 2709-son nizomining')
rep(p, 'belgilangan bo‘lsa-da, banklarning prudensial hisobot shaklida xususiy kapital olinishi me’yorni taxminan 7,1 foiz punktga pasaytirishi '
       '(122,2 foiz; regulyativ kapital olinganda – 129,4 foiz) aniqlanib, Markaziy bank tomonidan o‘rnatiladigan me’yor elementlarida barqaror '
       'moliyalashtirish manbai sifatida regulyativ kapitalni belgilash va bu elementlar hamda koeffitsiyentlarni oshkor qilish taklif etilgan;',
    'belgilangani, banklarning prudensial hisobot shaklida esa xususiy kapital olinib, kapitalga kiritiladigan muddatsiz obligatsiyalar va '
    'subordinar qarz bir yildan ortiq muddatli majburiyatlar tarkibida hisobga olinishi aniqlanib (kapital instrumentlari ikki marta '
    'hisoblanmaganda ikki yondashuv o‘rtasidagi farq tahlil qilingan uchta bankda 1,3 foiz punktdan oshmaydi), Markaziy bank tomonidan '
    'o‘rnatiladigan me’yor elementlarida kapital instrumentlarini bir marta – kapital sifatida – hisobga olish tartibini belgilash hamda bu '
    'elementlar va koeffitsiyentlarni oshkor qilish taklif etilgan;')

# 2) 2.3-§ – SQB, NSFR
p = D.para('Ikki bank misolida olingan natijalarni umumlashtirish uchun “O‘zsanoatqurilishbank” ATB ning 2025-yil yakunidagi')
rep(p, 'xususiy kapitaldan (11 875,2 mlrd so‘m) 4 323,9 mlrd so‘mga ko‘p bo‘lgani uchun regulyativ kapital olinganda me’yor 122,2 foiz emas, '
       '129,4 foizni tashkil etadi, ya’ni amaldagi hisobot shakli me’yorni 7,1 foiz punktga past, ehtiyotkorroq ko‘rsatadi.',
    'xususiy kapitaldan (11 875,2 mlrd so‘m) 4 323,9 mlrd so‘mga ko‘p, biroq bu farqning asosiy qismini qo‘shimcha birinchi darajali kapitalga '
    'kiritilgan muddatsiz obligatsiyalar (3 614,6 mlrd so‘m) va ikkinchi darajali kapitalga kiritilgan subordinar qarz (1 524,8 mlrd so‘m) tashkil '
    'etadi. Ular hisobot shaklida bir yildan ortiq muddatli majburiyatlar tarkibida 100 foizlik koeffitsiyent bilan allaqachon hisobga olingan. '
    'Shu sababli kapital instrumentlari ikki marta hisoblanmagan holda regulyativ kapital olinganda me’yor 120,9 foizni (Bazel III dagidek '
    'chegirmalargacha olinganda – 123,4 foizni) tashkil etadi, ya’ni hisobot shaklidagi qiymatdan (122,2 foiz) 1,3 foiz punktdan ko‘p farq qilmaydi.')
rep(p, 'Shu sababli ularda barqaror moliyalashtirish manbai sifatida regulyativ kapitalni belgilash va elementlar hamda koeffitsiyentlarni '
       'oshkor qilish maqsadga muvofiq.',
    'Shu sababli ularda kapital instrumentlarini barqaror moliyalashtirish tarkibida bir marta – kapital sifatida – hisobga olish tartibini aniq '
    'belgilash hamda elementlar va koeffitsiyentlarni oshkor qilish maqsadga muvofiq.')

# 3) shu xatboshidagi izoh (2709-son nizom, 53–55-bandlar …)
fn_done = 0
for f in D.fx.findall(qn('w:footnote')):
    t = tx(f)
    if 'Regulyativ kapital regulyativ kapital monandligi (17,02 foiz)' in t:
        pe = [x for x in f.findall(qn('w:p')) if '17,02' in tx(x)][0]
        P = Paragraph(pe, None)
        old = ('(“Likvidlilik tahlili”, 16-jadval; “Prudensial me’yorlar hisob-kitobi”, 22-jadval). Regulyativ kapital regulyativ kapital '
               'monandligi (17,02 foiz) va risk bo‘yicha tortilgan aktivlar (95 179,7 mlrd so‘m) asosida muallif tomonidan hisoblandi.')
        assert tx(pe).count(old) == 1
        D.replace(P, old, '(“Regulyativ kapital hisob-kitobi”, 5-jadval; “Likvidlilik tahlili”, 16-jadval; “Bank majburiyatlari tahlili”, '
                          '18-jadval; “Prudensial me’yorlar hisob-kitobi”, 22-jadval). Hisob-kitob muallif tomonidan bajarildi.')
        fn_done += 1
    if t.rstrip().endswith('16-jadval, 1.4, 1.12, 2.3–2.6-qatorlar; hisob-kitob muallif tomonidan bajarildi.') and '“Aloqabank” ATB ning 2020–2025' in t:
        pe = [x for x in f.findall(qn('w:p')) if '2.3–2.6-qatorlar' in tx(x)][0]
        D.replace(Paragraph(pe, None), '16-jadval, 1.4, 1.12, 2.3–2.6-qatorlar; hisob-kitob muallif tomonidan bajarildi.',
                  '5-jadval; 16-jadval, 1.4, 1.12, 2.3–2.6-qatorlar; 17- va 18-jadvallar; “Mikrokreditbank” ATB bo‘yicha – crs005L shaklining '
                  '16- va 17-jadvallari; hisob-kitob muallif tomonidan bajarildi.')
        fn_done += 1
assert fn_done == 2, fn_done

# 4) 2.3-§ – SQB, LCR
p = D.para('Jadval ma’lumotlari ikki yondashuv o‘rtasidagi tafovut moliyaviy barqarorlik bloklari bo‘yicha bir tekis')
rep(p, 'Koeffitsiyent minimal talabdan yuqori qolsa-da, likvidlik buferining qariyb uchdan bir qismi xalqaro standart bo‘yicha bufer '
       'hisoblanmaydi, demak, bu blokda ham yashirin risk mavjud.',
    'Chiqarilgan mablag‘larning asosiy qismi (4 882,6 mlrd so‘m) boshqa banklarga overnayt joylashtirilgan mablag‘lar bo‘lib, Bazel III da ular '
    '30 kunlik kirim sifatida hisobga olinadi, shu sababli 154,1 foiz koeffitsiyentning quyi chegarasi hisoblanadi. Koeffitsiyent minimal '
    'talabdan ancha yuqori qolsa-da, likvidlik buferining qariyb uchdan bir qismi xalqaro standart bo‘yicha yuqori likvid aktiv hisoblanmaydi, '
    'ya’ni buferning tarkibi milliy va xalqaro yondashuvda sezilarli farq qiladi.')

# 5) 2.3-§ – Aloqabank (+ Mikrokreditbank)
p = D.para('Likvidlik ko‘rsatkichlari ham “O‘zsanoatqurilishbank” ATB dagi qonuniyatni tasdiqlaydi.')
rep(p, 'xorijiy banklardagi vakillik hisobvaraqlaridagi mablag‘lardan iborat.',
    'boshqa banklardagi xorijiy valyutadagi vakillik hisobvaraqlaridan iborat, shundan 1 209,7 mlrd so‘mi mahalliy banklarda joylashgan.')
rep(p, 'koeffitsiyent 103,2 foizgacha, ya’ni minimal talab darajasigacha pasayadi.',
    'koeffitsiyent 103,2 foizgacha, ya’ni minimal talab darajasigacha pasayadi. “Mikrokreditbank” ATB da bu holat yanada yaqqol namoyon bo‘ladi, '
    'ya’ni yuqori likvid aktivlarning 52,4 foizi (1 986,9 mlrd so‘m) boshqa banklardagi vakillik hisobvaraqlaridan iborat bo‘lib, ular '
    'chiqarilganda koeffitsiyent 231,9 foizdan 110,3 foizgacha pasayadi.')
rep(p, 'Sof barqaror moliyalashtirish me’yori esa milliy hisobot shakli bo‘yicha 103,0 foizni tashkil etib, Bazel III dagidek xususiy kapital '
       '(3 846,9 mlrd so‘m) o‘rniga regulyativ kapital (5 026,8 mlrd so‘m) olinganda 108,7 foizga oshadi.',
    'Sof barqaror moliyalashtirish me’yori esa “Aloqabank” ATB da milliy hisobot shakli bo‘yicha 103,0 foizni tashkil etadi, regulyativ kapital '
    'olinganda ham (subordinar qarz ikki marta hisoblanmagan holda) 103,5 foizdan oshmaydi, ya’ni bank barqaror moliyalashtirish bo‘yicha '
    'minimal talab chegarasida turadi.')

# 6) 3.1-§ – jadval katagi
n = 0
for t in D.d.tables:
    for row in t.rows:
        for cell in row.cells:
            for cp in cell.paragraphs:
                if '27,3–30,9 foizi chiqariladi, NSFR metodika farqi 5,4–5,7 foiz punkt' in tx(cp._p):
                    D.replace(cp, '27,3–30,9 foizi chiqariladi, NSFR metodika farqi 5,4–5,7 foiz punkt',
                              '27,3–52,4 foizi chiqariladi, NSFR metodika farqi 5,4 foiz punkt'); n += 1
assert n == 1, n
D.save(F)
print(pardoz(F))
