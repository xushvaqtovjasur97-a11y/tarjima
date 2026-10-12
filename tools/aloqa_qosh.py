# «Aloqabank» ATB ma'lumotlarini dissertatsiyaga KO'K rangda qo'shish (qizil – kamchiliklar bo'yicha tuzatishlar, yashil – taqrizchi)
import sys, json, copy
sys.path.insert(0, '/home/user/tarjima/tools'); sys.path.insert(0, '/home/user/tarjima/manbalar/hisob_kitob')
import kamlib as K
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
import mbii as M

K.RED = '0070C0'   # ko'k rang
SRC = sys.argv[1] if len(sys.argv) > 1 else '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_tuzatilgan_v2.docx'
DST = sys.argv[2] if len(sys.argv) > 2 else '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_tuzatilgan_v3.docx'
HK = '/home/user/tarjima/manbalar/hisob_kitob/'
A = json.load(open(HK + 'aloqa23.json'))
AM = json.load(open(HK + 'aloqa_mbii.json'))
D = K.Doc(SRC)
Y = [str(y) for y in range(2021, 2026)]


def f(x, n=1, sign=False):
    s = f'{abs(x):,.{n}f}'.replace(',', ' ').replace('.', ',')
    if x < 0:
        return '−' + s
    return ('+' + s) if sign else s


FN = dict(
    ALFORM='«Aloqabank» ATB ning 2020–2025-yillar 31-dekabr holatiga Markaziy bankka taqdim etilgan prudensial hisobot shakllari (crs031C, crs031L, crs031O): regulyativ kapital hisob-kitobi (5-jadval), bank balansi (1-jadval), aktivlar tasnifi (2-jadval), likvidlilik tahlili (16-jadval) va prudensial me’yorlar hisob-kitobi (22-jadval). Shakllardagi jami aktivlar, xususiy kapital, kreditlar, sof foyda va muammoli kreditlar ulushi 2021–2025-yillarda O‘zbekiston Respublikasi Markaziy banki statistikasi va «SNS» reyting agentligi bazasidagi qiymatlarga to‘liq mos keladi.',
    ALIFRS='Joint-Stock Commercial “Aloqabank”: Consolidated Financial Statements and Independent Auditor’s Report, 31 December 2022; 31 December 2023. – Tashkent: Audit Organization “PricewaterhouseCoopers” LLC, 2023–2024; Joint-Stock Commercial “Aloqabank”: Consolidated Financial Statements and Independent Auditor’s Report, 31 December 2024. – Tashkent: Audit Organization “Deloitte & Touche” LLC, 2025; Joint Stock Commercial “Aloqabank” and its subsidiaries: Consolidated Financial Statements and Independent Auditor’s Report, 31 December 2025. – Tashkent: Audit Organization “PricewaterhouseCoopers” LLC, 2026. – P. 1–2, 35, 55, 64–65. – https://aloqabank.uz. Har bir yil ko‘rsatkichlari keyingi yil hisobotidagi qiyosiy ma’lumotlardan olindi.',
    ALCALC='Tadqiqot natijasida muallif tomonidan «Aloqabank» ATB ning prudensial hisobot shakllari va MHXS bo‘yicha konsolidatsiyalangan moliyaviy hisobotlari asosida hisoblandi (manbalar/hisob_kitob/aloqa23.py, aloqa_mbii.json).',
)


def para_after(anchor_el, template_p, text, tag='AL'):
    """anchor_el elementidan keyin template_p uslubidagi yangi paragraf qo'shadi."""
    np_ = copy.deepcopy(template_p._p)
    for e in list(np_):
        if e.tag != qn('w:pPr'):
            np_.remove(e)
    ppr = np_.find(qn('w:pPr'))
    if ppr is not None:
        for e in ppr.findall(qn('w:sectPr')):
            ppr.remove(e)
    tmpl = [r for r in template_p._p.iter(qn('w:r')) if r.find(qn('w:t')) is not None]
    for e in D._build(tmpl[0] if tmpl else None, text):
        np_.append(e)
    anchor_el.addnext(np_)
    K.LOG.append((tag, 'new_para', '', text[:80]))
    return Paragraph(np_, template_p._parent)


# =====================================================================
# 1) 2.3-§: uchinchi davlat banki – «Aloqabank» ATB (2.3.5-jadval)
# =====================================================================
a = {y: A[y] for y in Y}
koef = [a[y]['KYKR'] for y in Y]
np_n = sum(a[y]['NP_n'] for y in Y); np_i = sum(a[y]['NP_i'] for y in Y)
eqd = [(a[y]['EQ_i'] / a[y]['EQ_n'] - 1) * 100 for y in Y]
anchor = D.para('Rasmdan ko‘rinib turibdiki, nisbiy tafovut koeffitsiyenti muammoli kreditlar ulushi bo‘yicha')
body = anchor._parent
p1 = para_after(anchor._p, anchor,
    'Ikki bank misolida olingan natijalar umumiy xususiyatga egami yoki yo‘qligini tekshirish uchun xuddi shu metodika uchinchi davlat banki – «Aloqabank» ATB ga ham qo‘llandi. Bu bank tizimli ahamiyatga molik banklar qatoriga kirmaydi, uning aktivlari 2025-yil yakunida '
    + f(a['2025']['TA_n']) + ' mlrd so‘mni tashkil etgan. Tahlil uchun bankning Markaziy bankka taqdim etilgan prudensial hisobot shakllari[[FN:' + FN['ALFORM'] + ']] va auditdan o‘tgan MHXS bo‘yicha konsolidatsiyalangan moliyaviy hisobotlari[[FN:' + FN['ALIFRS'] + ']] qo‘llandi. Birinchi darajali asosiy kapital monandligi «O‘zsanoatqurilishbank» ATB dagi kabi MHXS bo‘yicha xususiy kapitaldan ozchilik ulushi, joriy yil foydasi va imtiyozli aksiyalarni ayirib, rasmiy hisobot shaklidagi chegirmalarni qo‘llash orqali qayta hisoblandi. Natijalar 2.3.5-jadvalda keltirildi.')
cap_t = [p for p in D.paras() if p.text.strip() == '2.3.3-jadval'][0]
p2 = para_after(p1._p, cap_t, '2.3.5-jadval')
ttl_t = D.para('«Mikrokreditbank» ATB ko‘rsatkichlarining', after='2.3.3-jadval') if False else Paragraph(cap_t._p.getnext(), body)
p3 = para_after(p2._p, ttl_t, '«Aloqabank» ATB ko‘rsatkichlarining milliy va MHXS hisobotlari bo‘yicha qiyosi, 2021–2025-yillar[[FN:' + FN['ALCALC'] + ']]')
mkb_tbl = Table(cap_t._p.getnext().getnext(), body)
nt = copy.deepcopy(mkb_tbl._tbl)
p3._p.addnext(nt)
T = Table(nt, body)
rows = [
    ('Ko‘rsatkich', Y),
    ('Jami aktivlar, mlrd so‘m: milliy', [f(a[y]['TA_n']) for y in Y]),
    ('MHXS', [f(a[y]['TA_i']) for y in Y]),
    ('Xususiy kapital, mlrd so‘m: milliy', [f(a[y]['EQ_n']) for y in Y]),
    ('MHXS', [f(a[y]['EQ_i']) for y in Y]),
    ('Tafovut, %', [f(v, 1, True) for v in eqd]),
    ('Sof foyda (zarar), mlrd so‘m: milliy', [f(a[y]['NP_n']) for y in Y]),
    ('MHXS', [f(a[y]['NP_i']) for y in Y]),
    ('ROE, %: milliy', [f(a[y]['ROE_n']) for y in Y]),
    ('MHXS', [f(a[y]['ROE_i']) for y in Y]),
    ('Muammoli kreditlar (milliy) / 3-bosqich (MHXS), %', [f"{f(a[y]['NPL_n'])} / {f(a[y]['S3_i'])}" for y in Y]),
    ('Substandart kreditlar (milliy) / 2-bosqich (MHXS), %', [f"{f(a[y]['S2_n'])} / {f(a[y]['S2_i'])}" for y in Y]),
    ('Birinchi darajali asosiy kapital monandligi, %: milliy', [f(a[y]['CET1_n'], 2) for y in Y]),
    ('MHXS asosida qayta hisoblangan', [f(a[y]['CET1_i'], 2) for y in Y]),
    ('Regulyativ kapital monandligi, %: milliy', [f(a[y]['CAR_n'], 2) for y in Y]),
    ('MHXS asosida qayta hisoblangan', [f(a[y]['CAR_i'], 2) for y in Y]),
    ('Yashirin kredit riski koeffitsiyenti', [f(v, 2) for v in koef]),
]
trs = nt.findall(qn('w:tr'))
while len(trs) < len(rows):            # yetishmagan qatorlarni oxirgi ma'lumot qatoridan nusxalash
    c = copy.deepcopy(trs[-2]); trs[-1].addprevious(c); trs = nt.findall(qn('w:tr'))
for ri, (lab, vals) in enumerate(rows):
    D.set_cell(T, ri, 0, lab, tag='AL')
    for j, v in enumerate(vals):
        D.set_cell(T, ri, j + 1, v, tag='AL')
# talqin
y25 = a['2025']
p4 = para_after(nt, anchor,
    'Jadval ma’lumotlari «Aloqabank» ATB da ham ikki hisobot o‘rtasidagi tafovut muntazam xususiyatga ega ekanini ko‘rsatadi. Birinchidan, MHXS 9 bo‘yicha 3-bosqich kreditlari ulushi tahlil davrining barcha yillarida rasmiy muammoli kreditlar ulushidan '
    + f(min(koef), 1) + '–' + f(max(koef), 1) + ' barobar yuqori bo‘lgan (' + f(min(a[y]['S3_i'] for y in Y)) + '–' + f(max(a[y]['S3_i'] for y in Y)) + ' foizga nisbatan ' + f(min(a[y]['NPL_n'] for y in Y)) + '–' + f(max(a[y]['NPL_n'] for y in Y)) + ' foiz), ya’ni yashirin kredit riski koeffitsiyenti bu bankda ikki tadqiqot obyektiga nisbatan ham yuqori. 2022-yilda esa milliy tasnif bo‘yicha substandart kreditlar ulushi atigi '
    + f(a['2022']['S2_n']) + ' foizni, MHXS bo‘yicha 2-bosqich kreditlari ulushi esa ' + f(a['2022']['S2_i']) + ' foizni tashkil etgan. Ikkinchidan, daromadlilik bo‘yicha tafovut ayniqsa katta: besh yil davomida bank milliy hisobot bo‘yicha jami ' + f(np_n) + ' mlrd so‘m, MHXS bo‘yicha esa ' + f(np_i) + ' mlrd so‘m sof foyda olgan, 2025-yilda esa milliy hisobotda ' + f(y25['NP_n']) + ' mlrd so‘m foyda ko‘rsatilgan bo‘lsa, MHXS bo‘yicha bank ' + f(-y25['NP_i']) + ' mlrd so‘m zarar ko‘rgan. Bunga MHXS bo‘yicha kutilgan kredit zararlari zaxirasining 2024–2025-yillarda ' + f(a['2023']['ECL_i']) + ' mlrd so‘mdan ' + f(y25['ECL_i']) + ' mlrd so‘mgacha oshishi sabab bo‘lgan. Uchinchidan, MHXS bo‘yicha xususiy kapital 2023 va 2025-yillarda milliy hisobotdagidan 9,3 foiz kam bo‘lgan, qayta hisoblangan birinchi darajali asosiy kapital monandligi esa 2025-yilda '
    + f(y25['CET1_n'], 2) + ' foiz o‘rniga ' + f(y25['CET1_i'], 2) + ' foizni, regulyativ kapital monandligi ' + f(y25['CAR_n'], 2) + ' foiz o‘rniga ' + f(y25['CAR_i'], 2) + ' foizni tashkil etadi. Bank tizimli ahamiyatga molik bo‘lmagani uchun unga 2026-yildan buferlar bilan birga CET1 bo‘yicha 8,5 foiz va jami kapital bo‘yicha 12 foiz talab qo‘llanadi, shu sababli MHXS bo‘yicha ham talablar bajariladi, biroq kapital zaxirasi CET1 bo‘yicha '
    + f(y25['CET1_n'] - 8.5, 2) + ' foiz punktdan ' + f(y25['CET1_i'] - 8.5, 2) + ' foiz punktgacha qisqaradi.')
# likvidlik
p5 = para_after(p4._p, anchor,
    'Likvidlik ko‘rsatkichlari ham «O‘zsanoatqurilishbank» ATB dagi qonuniyatni tasdiqlaydi. 2025-yil yakunida «Aloqabank» ATB ning likvidlikni qoplash koeffitsiyenti milliy metodika bo‘yicha 142,0 foizni tashkil etgan, biroq yuqori likvid aktivlarning 27,3 foizi (1 246,9 mlrd so‘m) xorijiy banklardagi vakillik hisobvaraqlaridagi mablag‘lardan iborat. Ular Bazel III talabiga muvofiq yuqori likvid aktivlar tarkibidan chiqarilganda koeffitsiyent 103,2 foizgacha, ya’ni minimal talab darajasigacha pasayadi. Sof barqaror moliyalashtirish me’yori esa milliy hisobot shakli bo‘yicha 103,0 foizni tashkil etib, nizom matnidagidek xususiy kapital (3 846,9 mlrd so‘m) o‘rniga regulyativ kapital (5 026,8 mlrd so‘m) olinganda 108,7 foizga oshadi[[FN:' + FN['ALFORM'] + ' 16-jadval, 1.4, 1.12, 2.3–2.6-qatorlar; hisob-kitob muallif tomonidan bajarildi.]]. Demak, uchinchi bank misolida ham milliy hisobot asosidagi baholash aktivlar sifati va daromadlilikni oshirib, likvidlik buferining sifatini esa haqiqiydan yuqori ko‘rsatadi, bu esa 2.3-paragrafda aniqlangan tafovutlarning alohida bankka xos emas, balki tizimli xususiyatga ega ekanini tasdiqlaydi.')
# 2.3-§ xulosasi
D.insert('Tijorat banklari moliyaviy ko‘rsatkichlarining milliy va MHXS hisobotlari bo‘yicha qiyosiy tahlili asosida xulosa qilish mumkinki', 'muntazam xususiyatga ega bo‘lib',
         ' (bu uchta davlat banki – «O‘zsanoatqurilishbank», «Mikrokreditbank» va «Aloqabank» ATBlari misolida tasdiqlandi)', tag='AL')
# 2.3-§ ma'lumotlar paragrafi
D.insert('Har ikki bankning 2021–2025-yillar yakunidagi Markaziy bankka taqdim etilgan prudensial hisobot shakllari', 'moliyaviy hisobotlari bilan solishtirildi.',
         ' Natijalarning umumiyligini tekshirish uchun uchinchi davlat banki – «Aloqabank» ATB ning prudensial shakllari va MHXS hisobotlari ham qo‘llandi.', tag='AL')

# =====================================================================
# 2) 3.2-§ va 1-ilova: tanlanmadan tashqari tekshiruv (6-bank)
# =====================================================================
am = {y: AM[y] for y in Y}
b25 = am['2025']
D.insert('Rasmdan ko‘rinib turibdiki, besh yil davomida banklarning moliyaviy barqarorligi turlicha', 'asosan davlat ko‘magiga tayanishini ko‘rsatadi.',
         ' Indeks vaznlar va sinflar chegaralarini baholashda qatnashmagan oltinchi bank – «Aloqabank» ATB ma’lumotlarida ham tekshirildi (1-ilova, 1.6). Bankning MHXS asosidagi indeksi 2021–2024-yillarda '
         + f(min(am[y]['mbii'] for y in Y[:4]), 2) + '–' + f(max(am[y]['mbii'] for y in Y[:4]), 2) + ' ball bilan II sinfda bo‘lgan, 2025-yilda esa MHXS bo‘yicha zarar natijasida daromadlilik bloki '
         + f(b25['bloklar']['D'], 2) + ' ballgacha pasayib, indeks ' + f(b25['mbii'], 2) + ' ballga tushgan va 4.1-paragrafdagi qizil zona chegarasiga (2,75) yaqinlashgan. Milliy hisobot bo‘yicha esa xuddi shu yili indeks '
         + f(b25['milliy'], 3) + ' ballni, ya’ni I sinf chegarasini tashkil etgan. Demak, tanlanmadan tashqari bankda ham ΔMBII ('
         + f(b25['mbii'] - b25['milliy'], 2) + ' ball) va uning asosiy manbalari – aktivlar sifati va daromadlilik bloklari – beshta bank bo‘yicha olingan natijalarga mos keladi.', tag='AL')
# 1-ilova: 1.6 «Aloqabank» ATB jadvali (1.5 jadvalidan nusxa)
cap15 = D.para('1.5. «Ipoteka-bank» ATB')
t15 = D.table_after('1.5. «Ipoteka-bank» ATB')
after_el = t15._tbl.getnext()   # bo'sh paragraf
emp = Paragraph(after_el, body)
c16 = para_after(after_el, cap15, '1.6. «Aloqabank» ATB (vaznlar baholanishida qatnashmagan, tashqi tekshiruv)')
nt16 = copy.deepcopy(t15._tbl); c16._p.addnext(nt16)
T16 = Table(nt16, body)
IND = [('K1. CET1 monandligi, %', 'CET1'), ('K2. Regulyativ kapital monandligi, %', 'CAR'), ('K3. Kapital / aktivlar, %', 'LEV'),
       ('A1. 3-bosqich kreditlari ulushi, %', 'S3'), ('A2. 2-bosqich kreditlari ulushi, %', 'S2'), ('A3. Qoplanmagan NPL / kapital, %', 'NPLnet'),
       ('L1. Yuqori likvid aktivlar / aktivlar, %', 'HLA'), ('L2. Kreditlar / depozitlar, %', 'LTD'), ('D1. ROA, %', 'ROA'), ('D2. ROE, %', 'ROE'),
       ('D3. Xarajatlar / daromadlar, %', 'CI'), ('V1. Valyutadagi kreditlar ulushi, %', 'FXL'), ('V2. Valyutadagi majburiyatlar ulushi, %', 'FXLI')]
AN = {n: (dr, an) for _, _, n, dr, an in M.IND}
D.set_cell(T16, 0, 0, 'Ko‘rsatkich', tag='AL')
for j, y in enumerate(Y):
    D.set_cell(T16, 0, j + 1, y, tag='AL')
for ri, (lab, k) in enumerate(IND, start=1):
    D.set_cell(T16, ri, 0, lab, tag='AL')
    for j, y in enumerate(Y):
        v = am[y]['ind'].get(k)
        D.set_cell(T16, ri, j + 1, '–' if v is None else f"{f(v)} ({f(M.score(v, *AN[k]), 2)})", tag='AL')
for ri, (lab, b) in enumerate([('Kapital yetarliligi bloki', 'K'), ('Aktivlar sifati bloki', 'A'), ('Likvidlik bloki', 'L'), ('Daromadlilik bloki', 'D'), ('Valyuta riski bloki', 'V')], start=14):
    D.set_cell(T16, ri, 0, lab, tag='AL')
    for j, y in enumerate(Y):
        D.set_cell(T16, ri, j + 1, f(am[y]['bloklar'][b], 2), tag='AL')
D.set_cell(T16, 19, 0, 'MBII (sinf)', tag='AL')
for j, y in enumerate(Y):
    D.set_cell(T16, 19, j + 1, f"{f(am[y]['mbii'], 2)} ({M.cls(am[y]['mbii'])})", tag='AL')
e2 = copy.deepcopy(after_el); nt16.addnext(e2)

# =====================================================================
# 3) 4.3.2-jadval va 4.3-§, IV bob xulosalari
# =====================================================================
t = D.table('Yashirin kredit riski koeffitsiyenti | MBII-8: milliy hisobot')
D.add_row(t, len(t.rows) - 1, ['«Aloqabank» ATB', f(a['2025']['NPL_n']), f(a['2025']['S3_i']), f(a['2025']['KYKR'], 2), '3,26', f(b25['mbii'], 2), 'yashil', 'sariq'], tag='AL')
D.replace('Beshta bankda milliy hisobot va MHXS hisoboti asosidagi integral indeks', 'Beshta bankda', 'Oltita bankda', tag='AL')
D.insert('Jadvaldan ko‘rinib turibdiki, MHXS ma’lumotlariga o‘tish beshta bankdan uchtasining', 'sariq zonadan qizil zonaga o‘tadi.',
         ' Vaznlar baholanishida qatnashmagan «Aloqabank» ATB da ham xuddi shunday holat kuzatiladi: u milliy hisobot bo‘yicha yashil zonada (3,26 ball), MHXS bo‘yicha esa sariq zonada (' + f(b25['mbii'], 2) + ' ball) joylashadi, ya’ni oltita bankdan to‘rttasining zonasi yomonlashadi.', tag='AL')
D.insert('MHXS hisobotlari mavjud bo‘lgan beshta bank misolida milliy ko‘rsatkichlardan MHXS ko‘rsatkichlariga o‘tish', 'uchta bankning erta ogohlantirish zonasini yomon tomonga o‘zgartirishi ko‘rsatildi',
         ' (tashqi tekshiruv sifatida qo‘shilgan «Aloqabank» ATB bilan – oltita bankdan to‘rttasining)', tag='AL')

# =====================================================================
# 4) Kirish, II bob xulosalari, Xulosa
# =====================================================================
D.insert('Tadqiqotning obyekti sifatida O‘zbekiston Respublikasi tijorat banklari', '«Mikrokreditbank» ATB', ', natijalarni tekshirish uchun esa «Aloqabank» ATB', tag='AL')
D.insert('ATB misolida MHXS 9 bo‘yicha qadrsizlangan kreditlar ulushi 2021–2025-yillarda rasmiy muammoli kreditlar ulushidan mos ravishda', '1,6–3,0 va 2,8–3,3 barobar yuqori bo‘lgani aniqlangan',
         ' (uchinchi davlat banki – «Aloqabank» ATB da ' + f(min(koef), 1) + '–' + f(max(koef), 1) + ' barobar, 2025-yilda esa milliy hisobotdagi ' + f(y25['NP_n']) + ' mlrd so‘m foyda o‘rniga MHXS bo‘yicha ' + f(-y25['NP_i']) + ' mlrd so‘m zarar)', tag='AL')
D.insert('o‘yicha 3-bosqich kreditlari ulushining milliy tasnifdagi muammoli kreditlar ulushiga nisbati taklif etildi. Koeffitsiyent', '«Mikrokreditbank» ATB da 2,77–3,30',
         ', tekshiruv uchun qo‘shilgan «Aloqabank» ATB da ' + f(min(koef), 2) + '–' + f(max(koef), 2), tag='AL')
D.insert('sh darajasini o‘lchash uchun yashirin kredit riski koeffitsiyenti taklif etildi. U tahlil qilingan banklarda', '«Mikrokreditbank» ATB da 2,77–3,30 ni',
         ', «Aloqabank» ATB da ' + f(min(koef), 2) + '–' + f(max(koef), 2) + ' ni', tag='AL')
D.insert('Milliy prudensial hisobot va MHXS bo‘yicha hisobot o‘rtasidagi tafovut muntazam xususiyatga ega bo‘lib', '',
         ' Uchinchi davlat banki – «Aloqabank» ATB misolida bu xulosa tasdiqlandi: besh yillik sof foyda milliy hisobot bo‘yicha ' + f(np_n) + ' mlrd so‘m, MHXS bo‘yicha ' + f(np_i) + ' mlrd so‘m bo‘lib, 2025-yilda MHXS bo‘yicha bank zarar ko‘rgan, Bazel III bo‘yicha qayta hisoblangan likvidlikni qoplash koeffitsiyenti esa 142,0 foizdan 103,2 foizgacha pasayadi.', tag='AL')

# =====================================================================
# 5) Adabiyotlar ro'yxati (VI bo'lim) va qayta raqamlash
# =====================================================================
anc = [p for p in D.paras() if p.text[:4].rstrip('. ').isdigit() and 'Микрокредитбанк' in p.text][-1]
for s in reversed([
        '«Aloqabank» ATB ning Markaziy bankka taqdim etilgan prudensial hisobot shakllari (crs031C, crs031L, crs031O): bank balansi, daromadlar, regulyativ kapital hisob-kitobi, aktivlarni tasniflash, likvidlilik tahlili va prudensial me’yorlar jadvallari, 2020–2025-yillar 31-dekabr holatiga.',
        'Joint Stock Commercial “Aloqabank” and its subsidiaries. Consolidated Financial Statements and Independent Auditor’s Report for the year ended 31 December 2025. – Tashkent: PricewaterhouseCoopers, 2026. – https://aloqabank.uz.',
        'Joint-Stock Commercial “Aloqabank”. Consolidated Financial Statements and Independent Auditor’s Report for the years ended 31 December 2022, 2023 and 2024. – Tashkent: PricewaterhouseCoopers, Deloitte & Touche, 2023–2025.']):
    D.new_para_after(anc, '0. ' + s, tag='AL')
import re
start = False; n = 0
for p in D.paras():
    if p.text.strip() == 'FOYDALANILGAN ADABIYOTLAR RO‘YXATI':
        start = True; continue
    if start and p.text.strip() == 'ILOVALAR':
        break
    if start:
        m = re.match(r'^(\d+)\. ', p.text)
        if m:
            n += 1
            r0 = D._runs(p)[0].find(qn('w:t'))
            if r0.text.startswith(m.group(1) + '.'):
                r0.text = str(n) + r0.text[len(m.group(1)):]
            elif r0.text == m.group(1):
                r0.text = str(n)
print('adabiyotlar:', n)
D.replace('Dissertatsiyaning tuzilishi va hajmi.', 'ro‘yxati 183 ta manbani', f'ro‘yxati {n} ta manbani', tag='AL')
D.save(DST)

# mundarija sahifalari
sys.argv = ['x']
_src = open('/home/user/tarjima/tools/kam_tuzat.py').read().split('# ---------------- mundarija sahifa raqamlarini yangilash')[1].split("print('mundarija:'")[0]
exec(_src.split('\n', 1)[1])
print('mundarija:', mundarija(DST), mundarija(DST))
json.dump(K.LOG, open(HK + 'aloqa_qosh_log.json', 'w'), ensure_ascii=False, indent=0)
print('saqlandi', DST, len(K.LOG))
