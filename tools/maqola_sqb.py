"""Maqola: milliy tasnif va MHXS 9 bosqichlari o‘rtasidagi tafovut (SQB, 2021–2025) – barcha hisob-kitoblar bitta faylda."""
import re, sys
sys.path.insert(0, '/home/user/tarjima/tools')
import docx
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree
import wordart as W
import namuna_234 as C234
import namuna_233j as ST
import namuna_mezon as MZ
from pardoz import pardoz

DST = '/home/user/tarjima/Maqola_SQB_milliy_va_MHXS9.docx'
Y = [2021, 2022, 2023, 2024, 2025]
# ---- 1) milliy (brutto) va MHXS kreditlari, hisoblangan foizlar (mlrd so‘m)
NAT = [43147.8, 48075.9, 57124.4, 65498.4, 64935.6]
ACR = [1320.2, 1933.0, 2908.2, 3894.5, 6711.8]
IFR = [44529.794, 51298.388, 60407.775, 69772.021, 71008.798]
# ---- 2) bosqichlararo o‘tishlar (mln so‘m): (1→2, 1→3, 2→1, 2→3, 3→1, 3→2), segmentlar yig‘indisi
TR = {2021: [(2686846, 1176909, 934919, 764472, 112400, 1230420), (0, 25941, 5327666, 0, 1674, 0), (103543, 111459, 217446, 68552, 124708, 45260)],
      2022: [(995508, 1136042, 3460426, 1368111, 275721, 610152), (1843922, 0, 0, 0, 0, 0), (71340, 43508, 114386, 22971, 92739, 36548)],
      2023: [(3485974, 743080, 873121, 581673, 198302, 305326), (214886, 0, 1655737, 0, 0, 12464), (158379, 33490, 67031, 16067, 33611, 19838)],
      2024: [(6651620, 1354668, 1362905, 1203974, 275700, 424791), (2252854, 13769, 0, 0, 0, 0), (444442, 147513, 88009, 65526, 15381, 17219)],
      2025: [(6982044, 842765, 2120335, 1286264, 263861, 543115), (1668444, 0, 1668990, 0, 0, 0), (441567, 134550, 311604, 95512, 55371, 64733)]}
TRS = {y: [sum(seg[i] for seg in v) / 1000 for i in range(6)] for y, v in TR.items()}
# ---- 3) 3-bosqich va milliy muammoli (kesishma jadvaldan)
S3 = [2777.940, 3784.041, 2904.365, 4085.617, 5310.186]
NPL = [1787.830, 2685.671, 1536.951, 2494.525, 2272.875]
S3SS = [990.110, 1098.370, 1367.414, 1591.092, 3037.311]


def n1(x):
    return f'{x:,.1f}'.replace(',', ' ').replace('.', ',')


def pc(x):
    return f'{x:.1f}'.replace('.', ',')


if __name__ == '__main__':
    for y in Y:
        print(y, [round(x, 1) for x in TRS[y]])
    dd = docx.Document()
    sec = dd.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin, sec.right_margin, sec.top_margin, sec.bottom_margin = Cm(2.5), Cm(2), Cm(2), Cm(2)
    st = dd.styles['Normal']; st.font.name = 'Times New Roman'; st.font.size = Pt(14)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    st.paragraph_format.line_spacing = 1.15; st.paragraph_format.space_after = Pt(4)

    def par(text, bold=False, italic=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY, size=13, indent=True, ls=None):
        q = dd.add_paragraph(); q.alignment = align
        q.paragraph_format.first_line_indent = Cm(1.0) if indent else Cm(0)
        if ls:
            q.paragraph_format.line_spacing = ls
        r = q.add_run(text); r.bold = bold; r.italic = italic; r.font.size = Pt(size)
        return q

    def head(text):
        par(text, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, size=13, indent=False)

    def cap(num, text, kind='jadval'):
        if kind == 'jadval':
            par(f'{num}-jadval', italic=True, align=WD_ALIGN_PARAGRAPH.RIGHT, size=12, indent=False)
            par(text, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, size=12, indent=False, ls=1.0)
        else:
            par(f'{num}-rasm. {text}', bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, size=12, indent=False, ls=1.0)

    def src(text):
        par('Manba: ' + text, italic=True, size=10.5, indent=False, ls=1.0)

    def table(rows, widths, hdr_rows=1, size=10.5, left_cols=(0,)):
        t = dd.add_table(rows=len(rows), cols=len(widths)); t.style = 'Table Grid'; t.alignment = WD_TABLE_ALIGNMENT.CENTER; t.autofit = False
        for gc, w in zip(t._tbl.tblGrid.findall(qn('w:gridCol')), widths):
            gc.set(qn('w:w'), str(int(Cm(w).twips)))
        for ri, row in enumerate(rows):
            for ci, v in enumerate(row):
                c = t.cell(ri, ci); c.width = Cm(widths[ci])
                q = c.paragraphs[0]; q.paragraph_format.line_spacing = 1.0; q.paragraph_format.space_after = Pt(0)
                q.alignment = WD_ALIGN_PARAGRAPH.LEFT if (ci in left_cols and ri >= hdr_rows) else WD_ALIGN_PARAGRAPH.CENTER
                r = q.add_run(v); r.font.size = Pt(size); r.bold = ri < hdr_rows or (ci == 0 and v in ('Jami',))
                c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                if ri < hdr_rows:
                    sh = OxmlElement('w:shd'); sh.set(qn('w:val'), 'clear'); sh.set(qn('w:fill'), 'DEEBF7'); c._tc.get_or_add_tcPr().append(sh)
        return t

    # ---------------- sarlavha
    par('KREDIT RISKINI TAN OLISHDA MILLIY TASNIF VA MHXS 9 BOSQICHLARI O‘RTASIDAGI TAFOVUT: “O‘ZSANOATQURILISHBANK” ATB MISOLIDA',
        bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, size=14, indent=False)
    par('Maxmudov Omon Tuxtayevich', bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, size=13, indent=False)
    par('Annotatsiya. Maqolada “O‘zsanoatqurilishbank” ATB ning 2021–2025-yillardagi auditdan o‘tgan konsolidatsiyalangan moliyaviy hisobotlarida oshkor qilingan '
        'kesishma jadvallar asosida kredit portfelining Markaziy bank tasnifi toifalari va MHXS 9 bosqichlari bo‘yicha taqsimoti, bosqichlararo o‘tishlar va ularni '
        'belgilash mezonlari tahlil qilingan. Aniqlanishicha, milliy tasnif bo‘yicha muammoli kreditlar MHXS 9 bo‘yicha qadrsizlangan kreditlarning faqat 43–71 foizini '
        'qamraydi, 2025-yilda esa 3-bosqich kreditlarining 57,2 foizi milliy hisobotda “standart” yoki “substandart” toifasida bo‘lgan. Kredit riski sezilarli oshgan '
        '2-bosqich kreditlari ulushi 2023–2025-yillarda 15,8 foizdan 35,7 foizga ko‘tarilgan, biroq ularning birortasi ham milliy hisobotda muammoli kredit sifatida '
        'aks etmagan. Tafovutning asosiy manbalari – restrukturizatsiya qilingan kreditlarni tasniflash qoidalari va yaxshilanish uchun sinov muddatining yo‘qligi.',
        italic=True, size=12)
    par('Kalit so‘zlar: kredit riski, MHXS 9, kutilgan kredit zararlari, aktivlarni tasniflash, muammoli kreditlar, restrukturizatsiya, yashirin kredit riski.',
        italic=True, size=12)

    head('Kirish')
    par('Tijorat banklarining moliyaviy barqarorligini baholashda aktivlar sifati ko‘rsatkichlari markaziy o‘rinni egallaydi, biroq ularning qiymati kredit riskini qaysi '
        'mezonlar bo‘yicha tan olishga bog‘liq. O‘zbekistonda prudensial nazorat Markaziy bankning aktivlarni tasniflash qoidalariga asoslangan milliy hisobot bo‘yicha '
        'amalga oshiriladi, banklarning moliyaviy hisobotlari esa 2021-yildan Moliyaviy hisobotning xalqaro standartlari (MHXS) bo‘yicha tuziladi. Natijada bir bankning '
        'kredit portfeli bir vaqtning o‘zida ikki xil tasnifga ega bo‘ladi. Xalqaro valyuta jamg‘armasining moliya sektorini baholash dasturi doirasida qayd etilishicha, '
        'O‘zbekistonda restrukturizatsiya qilingan kreditlarning asosiy qismi “standart” toifada qoladi, 6 oygacha imtiyozli davr tasnifga ta’sir qilmaydi va muammoli '
        'toifadan chiqarish uchun sinov muddati belgilanmagan [5]. Maqolaning maqsadi – bu tafovutning ko‘lami, dinamikasi va sabablarini bitta yirik bank misolida '
        'miqdoriy baholash.')

    head('Adabiyotlar sharhi')
    par('Bazel qo‘mitasi kutilgan kredit zararlari bo‘yicha yo‘riqnomada buxgalteriya talablari va kredit riskini boshqarish amaliyotini uyg‘unlashtirishni tavsiya etgan [1], '
        'keyinchalik esa kutilgan zararlar modeliga o‘tishning regulyativ kapitalga ta’sirini bosqichma-bosqich hisobga olish qoidalarini qabul qilgan [2]. D. Bholat va '
        'hammualliflar muammoli kreditlarning buxgalteriya va nazorat ta’riflari mamlakatlar bo‘yicha turlicha ekanini va MHXS 9 bu ta’riflarni yaqinlashtirish imkonini '
        'berishini ko‘rsatgan [3]. J. Abad va J. Suarez kutilgan kredit zararlari modeli zaxiralarni iqtisodiy pasayish boshlanishida ertaroq oshirishini isbotlagan [4]. '
        'H. Huizinga va L. Laeven esa buxgalteriya hisobidagi erkinlik banklarga kapital ko‘rsatkichlarini yaxshiroq ko‘rsatish imkonini berishini aniqlagan [6]. '
        'Biroq O‘zbekiston banklari bo‘yicha ikki tasnif o‘rtasidagi tafovut bank darajasidagi kesishma ma’lumotlar asosida hali tizimli o‘rganilmagan.')

    head('Ma’lumotlar va metodologiya')
    par('Tadqiqotda “O‘zsanoatqurilishbank” ATB ning 2023-yil (2021–2023-yillar ma’lumotlari bilan) va 2025-yil (2024–2025-yillar) uchun auditdan o‘tgan konsolidatsiyalangan '
        'moliyaviy hisobotlari [7; 8] hamda bankning milliy hisobot ma’lumotlari [9] qo‘llandi. Hisobotlarda amortizatsiyalangan qiymatda baholanadigan mijozlarga kreditlar '
        'uch segment (korporativ, davlat va munitsipal tashkilotlar, jismoniy shaxslar) bo‘yicha MHXS 9 bosqichlari va Markaziy bank tasnifi toifalari kesimida '
        '(kesishma jadval), shuningdek bosqichlararo o‘tishlar bilan oshkor qilingan. Segmentlar ma’lumotlari yig‘ilib, bank bo‘yicha jami ko‘rsatkichlar hisoblandi. '
        'Bosqichlarni belgilash mezonlari bankning hisob siyosatidan, milliy tasnif mezonlari Markaziy bank nizomi [10] va toifalar ta’riflaridan olindi.')

    head('Tahlil va natijalar')
    par('1. Ikki hisobotda kredit portfeli hajmi. Milliy hisobot va MHXS bo‘yicha kreditlar hajmi farq qiladi (1-jadval). Farqning asosiy qismi hisoblangan foizlar bilan '
        'izohlanadi: MHXS bo‘yicha ular kreditning yalpi qiymatiga kiradi, milliy hisobotda esa alohida moddada yuritiladi. Ular qo‘shilganda farq 2021, 2023 va 2024-yillarda '
        '0,1–0,6 foizgacha qisqaradi, qolgan qismi konsolidatsiya doirasi va boshqa aktivlar bo‘yicha hisoblangan foizlar bilan bog‘liq.')
    cap(1, 'Kreditlar hajmi: milliy hisobot va MHXS, mlrd so‘m')
    rows = [['Yil', 'Milliy hisobot (brutto)', 'Hisoblangan foizlar', 'Milliy + foizlar', 'MHXS (yalpi qiymat)', 'Farq (MHXS − milliy), %', 'Farq (foizlar bilan), %']]
    for i, y in enumerate(Y):
        a = NAT[i] + ACR[i]
        rows.append([str(y), n1(NAT[i]), n1(ACR[i]), n1(a), n1(IFR[i]), pc(100 * (IFR[i] / NAT[i] - 1)), pc(100 * (IFR[i] / a - 1))])
    table(rows, [1.3, 2.6, 2.4, 2.5, 2.6, 2.6, 2.6])
    src('[7; 8; 9] asosida muallif hisobi.')

    par('2. Bosqichlarni belgilash mezonlari. Bankning hisob siyosatiga ko‘ra, MHXS 9 bosqichlari milliy toifalarni quyi chegara sifatida qo‘llaydi: “substandart” kredit '
        'kamida 2-bosqichga, “qoniqarsiz”, “shubhali” va “umidsiz” kreditlar esa 3-bosqichga kiritiladi. Bunga qo‘shimcha ravishda restrukturizatsiya, boshqa banklardagi '
        'qarzdorlik, tashqi reyting, garov qiymati, moliyaviy ko‘rsatkichlar va huquqiy holat hisobga olinadi, yaxshilanish uchun esa sinov muddati talab qilinadi (2-jadval).')
    cap(2, 'Kreditlarni milliy tasnif va MHXS 9 bosqichlari bo‘yicha tasniflash mezonlari')
    table([['Mezon', 'Milliy tasnif (Markaziy bank)', 'MHXS 9 bosqichlari (bank siyosati)']] + [list(r) for r in MZ.ROWS], [3.4, 6.3, 6.9], size=10, left_cols=(0, 1, 2))
    src('[5; 8; 10] asosida muallif tomonidan ishlab chiqildi.')

    par('3. Kredit portfelining bosqichlar va toifalar bo‘yicha tarkibi. 3-jadvaldan ko‘rinadiki, 1-bosqich kreditlarining ulushi 2023-yildagi 79,4 foizdan 2025-yilda '
        '56,9 foizgacha qisqargan, 2-bosqich kreditlarining ulushi esa 15,8 foizdan 35,7 foizgacha, ya’ni 2,3 barobar oshgan. 2-bosqich kreditlarining barchasi milliy '
        'hisobotda “standart” yoki “substandart” toifasida bo‘lib, 2025-yilda ularning 51,2 foizi “standart” deb tasniflangan. Milliy tasnif bo‘yicha muammoli deb tan '
        'olingan kreditlarning barchasi 3-bosqichda bo‘lsa-da, ular 3-bosqich kreditlarining faqat 42,8–71,0 foizini tashkil etgan.')
    cap(3, 'Kredit portfelining MHXS 9 bosqichlari va milliy tasnif toifalari bo‘yicha tarkibi, mlrd so‘m (qavsda – %)')
    hdr = ['Yil', 'Jami kreditlar', '1-bosqich (standart)', '2-bosqich, jami', '2-b.: standart', '2-b.: substandart', '3-bosqich, jami', '3-b.: standart va substandart', '3-b.: muammoli']
    table([hdr] + [[c.replace('\n', ' ') for c in r] for r in ST.ROWS], [1.1, 1.9, 2.0, 2.0, 1.9, 1.9, 1.9, 2.0, 1.9], size=9)
    src('[7; 8] asosida muallif hisobi. Izoh: bosqichlar uchun – jami kreditlardagi, toifalar uchun – tegishli bosqichdagi ulush.')

    par('4. Qadrsizlangan kreditlarning milliy tasnifdagi aksi. 1-rasm 3-bosqich kreditlarining milliy toifalar bo‘yicha tarkibini ko‘rsatadi. Milliy hisobotda muammoli '
        'hisoblanmagan qism 2021-yilda 35,6 foiz, 2022-yilda 29,0 foiz, 2023-yilda 47,1 foiz, 2024-yilda 38,9 foiz va 2025-yilda 57,2 foizni tashkil etgan. '
        '3-bosqich kreditlarining milliy muammoli kreditlarga nisbati (yashirin kredit riski koeffitsiyenti) mos ravishda 1,55; 1,41; 1,89; 1,64 va 2,34 ga teng bo‘lgan, '
        'ya’ni 2025-yilda milliy ko‘rsatkich qadrsizlangan kreditlarning yarmidan kamini aks ettirgan.')
    q = dd.add_paragraph(); q.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rid = W.add_chart(dd, C234.chart())
    q._p.append(etree.fromstring(W.chart_run_xml(rid, 15.0, 7.5, '1-rasm')))
    cap(1, '3-bosqich kreditlarining milliy tasnif toifalari bo‘yicha tarkibi, 2021–2025-yillar, % (ustunlar ustida – jami, mlrd so‘m)', 'rasm')
    src('[7; 8] asosida muallif hisobi.')

    par('5. 2025-yil holati. 4-jadval 2025-yil yakunidagi to‘liq kesishma jadvalni beradi. Milliy tasnif bo‘yicha muammoli kreditlar 2 272,9 mlrd so‘mni (3,2 foiz) tashkil '
        'etgan bo‘lsa, 3-bosqich kreditlari 5 310,2 mlrd so‘m (7,5 foiz) bo‘lgan. Milliy tasnifda “standart” deb ko‘rsatilgan 54 080,8 mlrd so‘mlik kreditlarning 24,0 foizi '
        '2-bosqichda, 1,4 foizi esa 3-bosqichda bo‘lgan. Kutilgan kredit zararlari bo‘yicha zaxira 4 250,9 mlrd so‘mni tashkil etib, uning 47,1 foizi 3-bosqich, 40,3 foizi '
        '2-bosqich kreditlariga to‘g‘ri kelgan; 3-bosqich kreditlarining zaxira bilan qoplanishi 37,7 foiz bo‘lgan.')
    cap(4, 'Kreditlarning milliy tasnif toifalari va MHXS 9 bosqichlari bo‘yicha taqsimoti, 2025-yil 31-dekabr, mlrd so‘m')
    tot = [sum(r[i] for r in MZ.M25) for i in (1, 2, 3)]; T = sum(tot)
    rows = [['Milliy toifa', '1-bosqich', '2-bosqich', '3-bosqich', 'Jami', 'Ulushi, %']]
    for r in MZ.M25:
        rows.append([r[0], n1(r[1] / 1000) if r[1] else '–', n1(r[2] / 1000) if r[2] else '–', n1(r[3] / 1000), n1(sum(r[1:]) / 1000), pc(100 * sum(r[1:]) / T)])
    rows.append(['Jami', n1(tot[0] / 1000), n1(tot[1] / 1000), n1(tot[2] / 1000), n1(T / 1000), '100,0'])
    rows.append(['Ulushi, %'] + [pc(100 * x / T) for x in tot] + ['100,0', ''])
    rows.append(['Kutilgan kredit zararlari bo‘yicha zaxira'] + [n1(x / 1000) for x in MZ.ECL25] + [n1(sum(MZ.ECL25) / 1000), ''])
    rows.append(['Zaxira bilan qoplanish, %'] + [pc(100 * e / x) for e, x in zip(MZ.ECL25, tot)] + [pc(100 * sum(MZ.ECL25) / T), ''])
    table(rows, [5.0, 2.4, 2.4, 2.4, 2.5, 2.0])
    src('[8] asosida muallif hisobi.')

    par('6. Bosqichlararo o‘tishlar. Hisobotlardagi harakat jadvallari kreditlarning yil davomida bir bosqichdan boshqasiga o‘tishini ko‘rsatadi (5-jadval). '
        '2024 va 2025-yillarda har yili 9 trln so‘mdan ortiq kreditlar 1-bosqichdan 2-bosqichga o‘tgan. 3-bosqichga yangi o‘tgan kreditlar (1→3 va 2→3) 2021–2025-yillarda '
        'yiliga 1,4–2,8 trln so‘mni tashkil etgan. 2025-yilda bank hukumat ko‘rsatmalari asosida 14 129,2 mlrd so‘mlik kreditlarni restrukturizatsiya qilgan va bundan '
        '184,7 mlrd so‘m zarar tan olgan [8]; bank siyosatiga ko‘ra takroriy restrukturizatsiya 2-bosqichga o‘tkazish asosi bo‘lgani 2-bosqichning keskin o‘sishini izohlaydi. '
        'Teskari o‘tishlar ham sezilarli: 2025-yilda 4,1 trln so‘m kredit 2-bosqichdan 1-bosqichga qaytgan, biroq hisobotda o‘tishlarning sabablari kreditlar kesimida oshkor qilinmaydi.')
    cap(5, 'Kreditlarning MHXS 9 bosqichlari o‘rtasida o‘tishi, mlrd so‘m')
    rows = [['Yil', '1→2', '1→3', '2→3', '2→1', '3→2', '3→1', '3-bosqichga sof oqim*']]
    for y in Y:
        a = TRS[y]
        net = a[1] + a[3] - a[4] - a[5]
        rows.append([str(y), n1(a[0]), n1(a[1]), n1(a[3]), n1(a[2]), n1(a[5]), n1(a[4]), n1(net)])
    table(rows, [1.3, 2.1, 2.0, 2.0, 2.1, 2.0, 2.0, 3.0])
    src('[7; 8] asosida muallif hisobi. * (1→3) + (2→3) − (3→1) − (3→2).')

    head('Muhokama')
    par('Natijalar ikki tasnif o‘rtasidagi tafovut tasodifiy emas, balki tizimli ekanini ko‘rsatadi. Birinchidan, milliy tasnif kreditning joriy to‘lov intizomiga, MHXS 9 '
        'esa riskning dastlabki tan olingan sanaga nisbatan o‘zgarishiga tayanadi. Ikkinchidan, milliy qoidalarda restrukturizatsiya qilingan kreditlar uchun yashovchanlik '
        'bahosi va sinov muddati yo‘q, shuning uchun to‘lov jadvali qayta tuzilgan kredit muammoli hisoblanmaydi. Uchinchidan, MHXS 9 bo‘yicha bosqichlarning milliy toifalarni '
        'quyi chegara sifatida qo‘llashi ikki tasnif o‘rtasidagi farq faqat bir tomonga – MHXS bo‘yicha riskning yuqoriroq bahosi tomonga yo‘nalishini ta’minlaydi. '
        'Bu xulosalar Bholat va hammualliflarning muammoli kreditlar ta’riflari standartlar bo‘yicha farq qilishi haqidagi xulosasiga [3] mos keladi.')

    head('Xulosa va takliflar')
    par('Tahlil quyidagi xulosalarga olib keldi: (1) milliy tasnif bo‘yicha muammoli kreditlar qadrsizlangan kreditlarning faqat bir qismini qamraydi va bu farq 2025-yilda '
        'eng yuqori darajaga yetgan; (2) kredit riski sezilarli oshgan kreditlar milliy hisobotda umuman muammoli sifatida aks etmaydi; (3) tafovutning asosiy manbai – '
        'restrukturizatsiya qilingan kreditlarni tasniflash qoidalari. Shu asosda quyidagilar taklif etiladi: Markaziy bankning aktivlarni tasniflash qoidalariga '
        'restrukturizatsiya qilingan kreditlar uchun yashovchanlik bahosi va kamida 6–12 oylik sinov muddatini kiritish; banklardan MHXS 9 bosqichlari va milliy toifalar '
        'kesishma jadvalini choraklik nazorat hisobotida taqdim etishni talab qilish; moliyaviy barqarorlikni baholashda milliy muammoli kreditlar ko‘rsatkichini '
        '3-bosqich va yashirin kredit riski koeffitsiyenti bilan to‘ldirish.')

    head('Foydalanilgan adabiyotlar')
    REFS = [
        'Basel Committee on Banking Supervision. Guidance on credit risk and accounting for expected credit losses. – Basel: BIS, December 2015.',
        'Basel Committee on Banking Supervision. Regulatory treatment of accounting provisions – interim approach and transitional arrangements. – Basel: BIS, March 2017.',
        'Bholat D., Lastra R.M., Markose S.M., Miglionico A., Sen K. Non-performing loans at the dawn of IFRS 9: regulatory and accounting treatment of asset quality // Journal of Banking Regulation. – 2018. – Vol. 19, No. 1. – P. 33–54.',
        'Abad J., Suarez J. Assessing the cyclical implications of IFRS 9 – a recursive model // European Systemic Risk Board Occasional Paper Series. – 2017. – No. 12.',
        'International Monetary Fund. Republic of Uzbekistan: Financial Sector Assessment Program – Financial System Stability Assessment. IMF Country Report No. 25/145. – Washington, DC: IMF, 2025.',
        'Huizinga H., Laeven L. Bank valuation and accounting discretion during a financial crisis // Journal of Financial Economics. – 2012. – Vol. 106, No. 3. – P. 614–634.',
        'Акционерный коммерческий банк “Узбекский промышленно-строительный банк” и его дочерние организации. Консолидированная финансовая отчетность и аудиторское заключение за годы, закончившиеся 31 декабря 2023, 2022 и 2021 гг. – Ташкент, 2024.',
        'Joint Stock Commercial Bank “Uzbek Industrial and Construction Bank” and its subsidiaries. Consolidated Financial Statements for the year ended 31 December 2025 and Independent Auditor’s Report. – Tashkent, 2026.',
        '“O‘zsanoatqurilishbank” ATB ning Markaziy bankka taqdim etilgan prudensial hisobot shakllari (crs003C, crs003L, crs003O), 2021–2025-yillar.',
        'Tijorat banklarida aktivlar sifatini tasniflash va aktivlar bo‘yicha ehtimoliy yo‘qotishlarni qoplash uchun zaxiralarni shakllantirish hamda ulardan foydalanish tartibi to‘g‘risidagi nizom (O‘zbekiston Respublikasi Markaziy banki Boshqaruvining qarori, ro‘yxat raqami 2696, 14.07.2015) // https://lex.uz.',
    ]
    for i, r in enumerate(REFS, 1):
        par(f'{i}. {r}', size=12, indent=False, ls=1.0)
    dd.save(DST)
    print(pardoz(DST), DST)
