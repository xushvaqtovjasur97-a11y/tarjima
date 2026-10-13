"""2.2-§ qayta tuzilishi (namuna, 2-variant): uch bank jadvali, takroriy rasmlar olib tashlanadi,
tizim rasmlari mezon chiziqlari bilan, umumlashtiruvchi jadvalga “Aloqabank” ustuni."""
import re, sys
sys.path.insert(0, '/home/user/tarjima/tools')
import docx
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from lxml import etree
import wordart as W
import namuna_222j as J
from pardoz import pardoz

DST = '/home/user/tarjima/Namuna_2.2-paragraf_qayta_tuzilishi.docx'
RED, BLUE, GREY = 'FF0000', '0070C0', '595959'
Y = ['2021', '2022', '2023', '2024', '2025']

A_SQB = ('“O‘zsanoatqurilishbank” ATB uch bank ichida eng yirigi: uning aktivlari 2021–2025-yillarda 56,5 trln so‘mdan 102,4 trln so‘mga, ya’ni 1,8 barobar o‘sgan. '
         'Jadval bu bankning asosiy xususiyatini ko‘rsatadi: rasmiy ko‘rsatkichlar – kapital, muammoli kreditlar, daromadlilik va likvidlik – barqaror va me’yoriy '
         'talablardan yuqori, ammo asosiy kapital monandligi izchil pasayib, substandart kreditlar ulushi tez o‘sgan, ya’ni barqarorlik zaxirasi asta-sekin qisqarib bormoqda.')
A_MKB = ('“Mikrokreditbank” ATB ning aktivlari 13,5 trln so‘mdan 30,7 trln so‘mga, ya’ni 2,3 barobar o‘sgan. Bu bankning asosiy xususiyati – ijtimoiy yo‘naltirilgan '
         'kreditlash bilan bog‘liq past daromadlilik va yuqori kredit riski: ROA butun davrda 1 foizdan past bo‘lib, 2024-yilda bank yirik zarar ko‘rgan, '
         'muammoli kreditlar ulushi esa uch bank ichida eng yuqori.')
A_ALQ = ('“Aloqabank” ATB tizimli ahamiyatga molik banklar qatoriga kirmaydi va tahlilga natijalarning umumiyligini tekshirish uchun kiritildi. Bankning aktivlari '
         '12,6 trln so‘mdan 31,5 trln so‘mga, ya’ni 2,5 barobar o‘sgan. Yuqoridagi ikki bankdan farqli ravishda u kreditlarni asosan mijozlar depozitlari hisobiga '
         'moliyalashtiradi (kreditlarning depozitlarga nisbati 93,7–121,1 foiz), biroq kapital va likvidlik bo‘yicha zaxirasi uch bank ichida eng kichik.')
KAP = [('Uch jadvalni solishtirish banklarning moliyaviy barqarorlik profili turlicha bo‘lsa-da, ularda umumiy qonuniyatlar ham borligini ko‘rsatadi. Kapital bo‘yicha '
        'uchala bankda ham regulyativ kapital monandligi minimumga yaqin yoki beqaror: “O‘zsanoatqurilishbank” ATB da zaxira 2,3–4,0 foiz punktni tashkil etgan, '
        'asosiy kapital monandligi esa 2022-yildagi 12,1 foizdan 2025-yilda 10,2 foizgacha pasayib, 8 foizlik minimumga yaqinlashgan; 2025-yilda birinchi darajali '
        'kapital monandligining 14 foizgacha oshishi asosan xalqaro bozorda joylashtirilgan qo‘shimcha birinchi darajali kapital obligatsiyalari hisobiga ta’minlangan. ', RED),
       ('“Aloqabank” ATB da zaxira 2023-yildan tashqari 1,1 foiz punktdan oshmagan, 2022-yilda esa hisoblangan qiymat minimumdan 0,5 foiz punkt past bo‘lgan. ', BLUE),
       ('“Mikrokreditbank” ATB da ko‘rsatkich keskin tebranadi: 2021-yilda 13,8 foiz bilan minimum chegarasida bo‘lgan ko‘rsatkich 2022-yilda 20,7 foizga ko‘tarilgan, '
        '2024-yilda 15,4 foizgacha tushib, 2025-yilda yana 18,5 foizga yetgan (2.2.2–2.2.4-jadvallar).', RED)]
KAP2 = [('Bu tebranishlarning sababi kapital manbalarini tahlil qilganda ayon bo‘ladi. ', RED),
        ('“Mikrokreditbank” ATB ning xususiy kapitali 2022-yilda 2,3 trln so‘mdan 3,6 trln so‘mga oshgan … (xatboshining davomi o‘zgarmaydi) … CAMELS reytingini yomonlashtirishini ko‘rsatgan edi. ', None),
        ('“Aloqabank” ATB da esa kapitalning 2023-yildagi o‘sishi (847,8 mlrd so‘m) asosan o‘z foydasi (607,7 mlrd so‘m) hisobiga ta’minlangan, '
         '2024–2025-yillarda esa kapital o‘sishining yarmidan ko‘pi tashqi manbalarga to‘g‘ri kelgan.', BLUE)]
AKT = [('Aktivlar sifati bo‘yicha ham uch bank turlicha manzarani namoyon etadi (2.2.2–2.2.4-jadvallar). ', RED),
       ('“O‘zsanoatqurilishbank” ATB da rasmiy muammoli kreditlar ulushi 2,2–3,8 foiz oralig‘ida past darajada saqlangan, biroq substandart kreditlar ulushi … '
        '24,5 foizga yetgan (hozirgi matn o‘zgarmaydi). ', None),
       ('“Aloqabank” ATB da rasmiy muammoli kreditlar ulushi uch bank ichida eng past (1,5–4,1 foiz), ammo substandart kreditlar ulushi 2023-yilda 11,8 foizga ko‘tarilgan. ', BLUE),
       ('Rasmiy ko‘rsatkichlarning bu manzarasini uchala bankning auditdan o‘tgan MHXS bo‘yicha moliyaviy hisobotlari bilan solishtirish kredit riskining haqiqiy '
        'ko‘lamini aniqroq ko‘rsatadi (2.2.5-rasm).', RED)]
MHXS = [('Rasmdan ko‘rinib turibdiki, uchala bankda ham … (hozirgi matn; “ikkala bankda ham” → “uchala bankda ham”, “Rasmda taqqoslash uchun keltirilgan uchinchi davlat banki – “Aloqabank” ATB da ham” → ““Aloqabank” ATB da ham”, “har ikki bankda ham” → “uchala bankda ham”).', GREY)]
ROE = [('Aktivlar sifatidagi bu muammolar banklarning daromadliligida ham aks etgan (2.2.2–2.2.4-jadvallar). ', RED),
       ('“O‘zsanoatqurilishbank” ATB ning kapital rentabelligi 2022-yildagi pasayishdan keyin tiklanib … keyin esa yana davlat kapitali hisobiga to‘ldirilgan (hozirgi matn o‘zgarmaydi). ', None),
       ('“Aloqabank” ATB da ROE 2023-yilda 27,4 foizga yetgan bo‘lsa-da, keyingi yillarda 6,6–7,3 foizgacha pasaygan; 2025-yilda esa bank milliy hisobot bo‘yicha '
        '251,5 mlrd so‘m foyda ko‘rsatgan bir paytda MHXS bo‘yicha 193,9 mlrd so‘m zarar qayd etgan (batafsil – 2.3-paragrafda).', BLUE)]
LIK = [('Likvidlik bo‘yicha “O‘zsanoatqurilishbank” ATB va “Mikrokreditbank” ATB Bazel III talablarini bajargan: … depozitlardan qariyb 1,1 barobar ko‘p bo‘lgan. '
        'Bunday tuzilma tashqi moliyalashtirish sharoiti yomonlashganda likvidlik va valyuta riskini oshiradi (“ikkala bank” iboralari shu tarzda aniqlashtiriladi). ', GREY),
       ('“Aloqabank” ATB da esa likvidlik zaxirasi torroq: LCR 2023-yilda 103,1 foizgacha, NSFR 2022-yilda 100,9 foizgacha tushgan, biroq kreditlarning depozitlarga '
        'nisbati 93,7–121,1 foiz bo‘lib, bank tashqi qarzlarga emas, depozit bazasiga tayanadi.', BLUE)]
SUM_T = 'Tahlil natijalarini umumlashtirish uchun bank tizimi va uch bankning so‘nggi ko‘rsatkichlari xalqaro mezonlar bilan solishtirildi (2.2.5-jadval).'
HDR = ['Ko‘rsatkich', 'Bank tizimi', 'SQB', 'MKB', 'Aloqabank', 'Xalqaro mezon', 'Baho']
SUMROWS = [
    ('Regulyativ kapital monandligi, %', '18,3', '17,0', '18,5', '14,1', '≥ 10,5 (Bazel III, bufer bilan)', 'Yetarli, ammo zaxira kichik va davlat kapitaliga bog‘liq'),
    ('Asosiy kapital monandligi, %', '14,7', '10,2', '16,4', '10,1', '≥ 7 (Bazel III, bufer bilan)', 'Yetarli, SQB va Aloqabank da zaxira kichik'),
    ('Muammoli kreditlar ulushi, %', '3,0', '2,5', '4,4', '2,5', '< 5 (erta ogohlantirish)', 'Rasmiy jihatdan qoniqarli'),
    ('MHXS 9 bo‘yicha 3-bosqich, %', '7,8 (2023)', '7,5', '12,3', '7,2', '< 5', 'Zaif'),
    ('Substandart va muammoli kreditlar, %', '15', '20,9', '13,4', '9,0', '< 10', 'SQB va MK da zaif'),
    ('ROA, %', '2,2', '2,0', '0,6', '0,9', '> 1 (CAMELS 1-reyting)', 'Tizim va SQB yuqori, MK va Aloqabank past'),
    ('LCR, %', '207,5', '222,9', '231,9', '142,0', '≥ 100 (Bazel III)', 'Yuqori'),
    ('NSFR, %', '119,9', '122,2', '113,0', '103,0', '≥ 100 (Bazel III)', 'Yetarli, Aloqabank da zaxira kichik'),
    ('Kreditlar / depozitlar, %', '158', '206,5', '208,5', '121,1', '≤ 100', 'SQB va MK da yuqori risk'),
]


def line_chart(series, ytitle, ymin, ymax, major, pos):
    xml, rows = W.line_table_chart(Y, series, ytitle=ytitle, ymin=ymin, ymax=ymax, major=major, fmt='0.0', w=31750, dtable=False)
    xml = xml.replace('<c:tickLblPos val="none"/>', '<c:tickLblPos val="low"/>', 1)
    sers = re.findall(r'<c:ser>.*?</c:ser>', xml, re.S)
    for s, sr, p in zip(sers, series, pos):
        if p:
            xml = xml.replace(s, s.replace('<c:cat>', W._dlbls(1000, p, sr[2], fmt='0.0') + '<c:cat>', 1), 1)
        elif p == '':
            pass
        else:
            xml = xml.replace(s, re.sub(r'<c:marker>.*?</c:marker>', '<c:marker><c:symbol val="none"/></c:marker>', s, flags=re.S), 1)
    xml = xml.replace('<c:size val="5"/>', '<c:size val="6"/>')
    xml = xml.replace('</c:plotArea>', '</c:plotArea>' + W._legend('b'), 1)
    return xml, rows


C223 = [('Substandart va muammoli kreditlar', [12.5, 15.9, 19.5, 21.0, 15.0], '7030A0'),
        ('Muammoli kreditlar, davlat ulushi mavjud banklar', [5.4, 3.9, 4.0, 3.9, 3.2], 'EB6834'),
        ('Muammoli kreditlar, jami', [5.2, 3.6, 3.5, 4.0, 3.0], '2A78D6'),
        ('Muammoli kreditlar, boshqa banklar', [4.1, 2.1, 2.4, 4.1, 2.6], '70AD47'),
        ('Erta ogohlantirish chegarasi (5%)', [5.0] * 5, 'C00000', 'dash')]
C224 = [('ROE', [6.1, 13.3, 14.2, 6.6, 12.4], 'EB6834'),
        ('ROA', [1.3, 2.5, 2.6, 1.4, 2.2], '2A78D6'),
        ('ROA mezoni – CAMELS 1-reyting (1%)', [1.0] * 5, 'C00000', 'dash')]

if __name__ == '__main__':
    dd = docx.Document()
    sec = dd.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin, sec.right_margin, sec.top_margin, sec.bottom_margin = Cm(3), Cm(1.5), Cm(2), Cm(2)
    st = dd.styles['Normal']; st.font.name = 'Times New Roman'; st.font.size = Pt(14)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    st.paragraph_format.line_spacing = 1.5; st.paragraph_format.space_after = Pt(0)

    def par(parts, bold=False, italic=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY, indent=True, size=14, ls=None):
        if isinstance(parts, str):
            parts = [(parts, None)]
        q = dd.add_paragraph(); q.alignment = align
        q.paragraph_format.first_line_indent = Cm(1.25) if indent else Cm(0)
        if ls:
            q.paragraph_format.line_spacing = ls
        for text, col in parts:
            r = q.add_run(text); r.bold = bold; r.italic = italic or col == GREY; r.font.size = Pt(size)
            if col:
                r.font.color.rgb = RGBColor.from_string(col)
        return q

    def h(t):
        par([(t, None)], bold=True, size=13, indent=False, align=WD_ALIGN_PARAGRAPH.LEFT)

    def chart(series, ytitle, ymin, ymax, major, pos, name):
        q = dd.add_paragraph(); q.alignment = WD_ALIGN_PARAGRAPH.CENTER
        rid = W.add_chart(dd, line_chart(series, ytitle, ymin, ymax, major, pos))
        q._p.append(etree.fromstring(W.chart_run_xml(rid, 16.0, 8.5, name)))

    par([('2.2-§ ni qayta tuzish (namuna, 2-variant)', None)], bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False)
    par([('Qizil – yangi/o‘zgartirilgan matn; ko‘k – “Aloqabank” ATB bo‘yicha; kulrang kursiv – o‘zgarmaydigan yoki faqat so‘z almashadigan qismlar ko‘rsatmasi.', GREY)], size=12, indent=False)

    h('1. Bank tizimi rasmlari: takror emas, mezon chiziqlari bilan')
    par([('2.2.2-rasmda 13 foizlik minimum chizig‘i allaqachon bor. 2.2.3 va 2.2.4-rasmlar ham chiziqli shaklga o‘tkazilib, mezon chiziqlari qo‘shiladi – shunda rasm 2.2.1-jadvalni takrorlamaydi, ko‘rsatkichlarni mezon bilan solishtiradi.', GREY)], size=12, indent=False)
    chart(C223, 'Kreditlardagi ulushi, %', 0, 25, 5, ('t', '', 'b', '', None), '2.2.3-rasm')
    par('2.2.3-rasm. O‘zbekiston bank tizimida aktivlar sifati ko‘rsatkichlari, %', bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False)
    par([('Matnga qo‘shiladi: ', GREY), ('Rasmda ko‘rinadiki, rasmiy muammoli kreditlar ulushi 5 foizlik erta ogohlantirish chegarasidan past bo‘lsa-da, substandart kreditlar bilan birgalikda bu ko‘rsatkich chegaradan 2,5–4,2 barobar yuqori.', RED)], size=13)
    chart(C224, 'Rentabellik, %', 0, 16, 4, ('t', 't', None), '2.2.4-rasm')
    par('2.2.4-rasm. O‘zbekiston bank tizimida aktivlar va kapital rentabelligi dinamikasi, %', bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False)
    par([('Matnga qo‘shiladi: ', GREY), ('Aktivlar rentabelligi butun davrda CAMELS tizimidagi eng yuqori reyting mezoni – 1 foizdan yuqori bo‘lgan, biroq 2021 va 2024-yillarda (1,3 va 1,4 foiz) bu mezonga yaqinlashgan.', RED)], size=13)

    h('2. Banklar: uchta alohida jadval')
    par([(J.INTRO, RED)])
    par([('2.2.2-jadval (“O‘zsanoatqurilishbank” ATB), 2.2.3-jadval (“Mikrokreditbank” ATB), 2.2.4-jadval (“Aloqabank” ATB) – avvalgi namunadagi kabi.', GREY)], size=12, indent=False)
    par([('2.2.2-jadvaldan keyin: ', GREY), (A_SQB, RED)])
    par([('2.2.3-jadvaldan keyin: ', GREY), (A_MKB, RED)])
    par([('2.2.4-jadvaldan keyin: ', GREY), (A_ALQ, BLUE)])

    h('3. Uch bankni bir-biriga bog‘lab tahlil (2.2.5, 2.2.6, 2.2.8-rasmlar olib tashlanadi)')
    par(KAP); par(KAP2); par(AKT)
    par([('2.2.5-rasm (hozirgi 2.2.7-rasm – MHXS 9 bo‘yicha 3-bosqich va rasmiy muammoli kreditlar, o‘zgarmaydi)', GREY)], size=12, indent=False)
    par(MHXS); par(ROE); par(LIK)

    h('4. Umumlashtiruvchi jadval (hozirgi 2.2.3-jadval → 2.2.5-jadval)')
    par([(SUM_T, RED)])
    par([('2.2.5-jadval', RED)], italic=True, align=WD_ALIGN_PARAGRAPH.RIGHT, indent=False)
    par([('Bank tizimi va tanlangan banklarning moliyaviy barqarorlik ko‘rsatkichlari: xalqaro mezonlar bo‘yicha baho, 2025-yil', None)], bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, ls=1.0)
    t = dd.add_table(rows=len(SUMROWS) + 1, cols=7); t.style = 'Table Grid'; t.autofit = False
    Wd = [Cm(3.4), Cm(1.7), Cm(1.6), Cm(1.6), Cm(1.9), Cm(2.7), Cm(3.6)]
    for gc, w in zip(t._tbl.tblGrid.findall(qn('w:gridCol')), Wd):
        gc.set(qn('w:w'), str(int(w.twips)))
    for ri, row in enumerate([HDR] + SUMROWS):
        for ci, v in enumerate(row):
            c = t.cell(ri, ci); c.width = Wd[ci]
            q = c.paragraphs[0]; q.paragraph_format.line_spacing = 1.0
            q.alignment = WD_ALIGN_PARAGRAPH.LEFT if ci in (0, 6) and ri else WD_ALIGN_PARAGRAPH.CENTER
            r = q.add_run(v); r.font.size = Pt(10); r.bold = ri == 0
            if ci == 4 and ri:
                r.font.color.rgb = RGBColor.from_string(BLUE)
            elif ci == 6 and ri and 'Aloqabank' in v or (ci == 4 and ri == 0):
                r.font.color.rgb = RGBColor.from_string(BLUE)
    par([('Izoh: SQB – “O‘zsanoatqurilishbank” ATB, MKB – “Mikrokreditbank” ATB.', None)], italic=True, size=11, indent=False, ls=1.0)
    par([('Jadvaldan keyingi xatboshi va paragraf xulosasida “ikki bank” → “uch bank”; xulosaga qo‘shiladi: ', GREY),
         ('tizimli ahamiyatga molik bo‘lmagan “Aloqabank” ATB misolida esa bu qonuniyatlar faqat yirik davlat banklariga xos emasligi tasdiqlandi.', BLUE)], size=13)
    dd.save(DST)
    print(pardoz(DST), DST)
