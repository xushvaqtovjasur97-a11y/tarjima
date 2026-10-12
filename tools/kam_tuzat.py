# Kamchiliklar jadvali (49 band) bo'yicha tuzatishlarni taqrizchi tuzatgan nusxaga (yashil) qizil rangda kiritish
# Kirish: /tmp/claude-0/up_kam/tuz.docx  Chiqish: Dissertatsiya_Maxmudov_O.T_tuzatilgan_v2.docx
import sys, json
sys.path.insert(0, '/home/user/tarjima/tools'); sys.path.insert(0, '/home/user/tarjima/manbalar/hisob_kitob')
import kamlib as K

SRC = sys.argv[1] if len(sys.argv) > 1 else '/home/user/tarjima/taqriz/Dissertatsiya_taqrizchi_tuzatgan.docx'
DST = sys.argv[2] if len(sys.argv) > 2 else '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_tuzatilgan_v2.docx'
D = K.Doc(SRC)
HK = '/home/user/tarjima/manbalar/hisob_kitob/'
R32 = json.load(open(HK + 'mbii32.json'))
FC = json.load(open(HK + 'fc42.json'))
VZ = json.load(open(HK + 'vazn32.json'))
RB = json.load(open(HK + 'robust_kam.json'))


def f2(x, n=2):
    return f'{x:.{n}f}'.replace('.', ',').replace('-', '−')


# ---------------- izoh matnlari (to'liq iqtibos)
FN = dict(
    N2709='O‘zbekiston Respublikasi Markaziy banki Boshqaruvining 2015-yil 22-iyuldagi 19/14-son qarori bilan tasdiqlangan «Tijorat banklarining likvidliligini boshqarishga qo‘yiladigan talablar to‘g‘risida»gi nizom (ro‘yxat raqami 2709) // O‘zbekiston Respublikasi qonun hujjatlari to‘plami. – 2015. – 32 (688)-son. – 437-modda. – 53–55-bandlar va 2-ilova; «O‘zsanoatqurilishbank» ATB ning 2025-yil 30-dekabr holatiga Markaziy bankka taqdim etilgan prudensial hisobot shakllari («Likvidlilik tahlili», 16-jadval; «Prudensial me’yorlar hisob-kitobi», 22-jadval). Regulyativ kapital regulyativ kapital monandligi (17,02 foiz) va risk bo‘yicha tortilgan aktivlar (95 179,7 mlrd so‘m) asosida muallif tomonidan hisoblandi.',
    SQB22='«O‘zsanoatqurilishbank» ATB ning 2025-yil 30-dekabr holatiga Markaziy bankka taqdim etilgan prudensial hisobot shakli «Prudensial me’yorlar hisob-kitobi» (22-jadval): yuqori likvid aktivlarning jami aktivlardagi ulushi – minimal talab 10 foiz; I darajali asosiy kapital monandligi – 8 foiz; I darajali kapital – 10 foiz; regulyativ kapital – 13 foiz; leverej – 6 foiz.',
    SQB16='«O‘zsanoatqurilishbank» ATB ning 2025-yil 30-dekabr holatiga Markaziy bankka taqdim etilgan prudensial hisobot shakli «Likvidlilik tahlili» (16-jadval), 1.4, 1.5, 1.11 va 2.3-qatorlar; Basel Committee on Banking Supervision. Basel III: The Liquidity Coverage Ratio and liquidity risk monitoring tools. – Basel: BIS, January 2013. – Para. 31, 49–50. Hisob-kitob muallif tomonidan bajarildi.',
    MBS25='O‘zbekiston Respublikasi Markaziy banki. Moliyaviy barqarorlik sharhi: 2025-yil. – Toshkent, 2026. – B. 129 (1-jadval «Banklar kapitali uchun o‘rnatilgan minimal talablar va buferlar»), 171.',
    MBS25A='O‘zbekiston Respublikasi Markaziy banki. Moliyaviy barqarorlik sharhi: 2025-yil. – Toshkent, 2026. – B. 30 (aktivlar); kreditlar (brutto) – «SNS» reyting agentligining O‘zbekiston tijorat banklari asosiy ko‘rsatkichlari bazasi (Banks_2025_IV, 33 bank) asosida muallif hisobi (603,8 trln so‘m).',
    EBARD25='European Banking Authority. Risk Dashboard: Data as of Q4 2024. – Paris: EBA, March 2025. – P. 33 (Statistical Annex: IFRS 9 specific data).',
    CGM='Cameron A.C., Gelbach J.B., Miller D.L. Bootstrap-Based Improvements for Inference with Clustered Errors // The Review of Economics and Statistics. – 2008. – Vol. 90, No. 3. – P. 414–427.',
    WILSON='Wilson E.B. Probable Inference, the Law of Succession, and Statistical Inference // Journal of the American Statistical Association. – 1927. – Vol. 22, No. 158. – P. 209–212.',
    NICKELL='Nickell S. Biases in Dynamic Models with Fixed Effects // Econometrica. – 1981. – Vol. 49, No. 6. – P. 1417–1426.',
    SC26='Sullivan & Cromwell LLP. FFIEC Proposes Revisions to CAMELS Rating System: Proposed Revisions Would Limit – but Would Not Eliminate – Consideration of Management Component and Would Emphasize Material Financial Risks. – New York, 20 May 2026. – https://www.sullcrom.com/insights/memo/2026/May/FFIEC-Proposes-Revisions-CAMELS-Rating-System.',
    EBARD='European Banking Authority. EBA updates its risk dashboard for EU banking sector: press release. – London, 11 December 2014. – https://www.eba.europa.eu/eba-updates-its-risk-dashboard-for-eu-banking-sector.',
    ROBUST='Muallif hisob-kitobi: manbalar/hisob_kitob/robust_kam.py va robust_kam.json (25 ta bank paneli, 2021-yil IV chorak – 2025-yil IV chorak).',
)
FSSA_URL = ' – https://www.imf.org/en/publications/cr/issues/2025/06/23/republic-of-uzbekistan-financial-sector-assessment-program-financial-system-stability-567987.'

# =====================================================================
# A. HISOB-KITOB XATOLARI
# =====================================================================
# 1-band: NSFR – regulyativ va xususiy kapital (2.3-§ ga hisob-kitob qo'shildi)
D.insert('to‘rt variantda hisoblandi. Likvidlikni', 'to‘rt variantda hisoblandi.',
         ' Hisob-kitob jarayonida nizom matni va bankning hisobot shakli o‘rtasida farq aniqlandi: nizomda mavjud barqaror moliyalashtirish tarkibiga bankning regulyativ kapitali kiritilishi belgilangan, bankning prudensial hisobot shaklida esa xususiy kapital ko‘rsatilgan[[FN:' + FN['N2709'] + ']]. «O‘zsanoatqurilishbank» ATB da regulyativ kapital (16 199,1 mlrd so‘m) xususiy kapitaldan (11 875,2 mlrd so‘m) 4 323,9 mlrd so‘mga ko‘p bo‘lgani uchun nizom matni bo‘yicha hisoblanganda me’yor 122,2 foiz emas, 129,4 foizni tashkil etadi, ya’ni amaldagi hisobot shakli me’yorni 7,1 foiz punktga past, ehtiyotkorroq ko‘rsatadi. Bu farq barqaror moliyalashtirish manbai sifatida qaysi kapital ko‘rsatkichi olinishini nizomda aniq belgilash zarurligini ko‘rsatadi.',
         tag='1')
# 2-band: 2,7–2,8 – dekompozitsiya tartibi izohi (2.3-§)
D.replace('Bazel III koeffitsiyentlari me’yorni 5,4 foiz punktga oshiradi', 'MHXS tuzatishlari esa 2,8 foiz punktga pasaytiradi',
          'MHXS tuzatishlari esa 2,7–2,8 foiz punktga pasaytiradi (farq dekompozitsiya tartibiga bog‘liq: Markaziy bank metodikasi bo‘yicha 2,7, Bazel III metodikasi bo‘yicha 2,8 foiz punkt)', tag='2')

# 3-4-bandlar: CET1 va regulyativ kapital – yagona metodika (2.3-§), 3.1.5-jadval, 1-ilova, 3.2, 4.1-4.3
t = D.table('SQB: milliy | SQB: MHXS')
D.set_cell(t, 1, 2, '8,87 (3,35)', tag='3'); D.set_cell(t, 2, 2, '15,32 (3,93)', tag='3')
D.set_cell(t, 1, 4, '18,24 (4,00)', tag='4'); D.set_cell(t, 2, 4, '19,96 (4,00)', tag='4')
D.set_cell(t, 14, 2, '2,81', tag='3')
D.replace('Izoh: qavs ichida – ball. MHXS variantida CET1', 'MHXS variantida CET1 monandligi (MHXS bo‘yicha xususiy kapital − ozchilik ulushi − joriy yil sof foydasi) / risk bo‘yicha tortilgan aktivlar formulasi bo‘yicha, regulyativ kapital monandligi – hisobotda oshkor qilingan qiymat bo‘yicha olindi',
          'MHXS variantida CET1 va regulyativ kapital monandligi 2.3-paragrafda Markaziy bank metodikasi bo‘yicha MHXS ma’lumotlari asosida qayta hisoblangan qiymatlarda olindi (2.3.4-jadval va «Mikrokreditbank» ATB bo‘yicha jadval), chunki banklar MHXS hisobotida oshkor qiladigan kapital koeffitsiyentlari milliy buxgalteriya hisobi ma’lumotlari asosida hisoblanadi', tag='3')
D.insert('Izoh: qavs ichida – ball. MHXS variantida CET1', 'valyuta bloki bitta ko‘rsatkich bo‘yicha baholandi.',
         ' Yuqori likvid aktivlar MHXS hisobotlarida alohida oshkor qilinmagani sababli L1 ko‘rsatkichi ikkala variantda milliy balans asosida olindi, shu sababli likvidlik blokidagi ΔMBII faqat L2 ko‘rsatkichi hisobiga shakllanadi va hisobot standartining likvidlikka ta’siri to‘liq aks etmaydi.', tag='30')
P = 'Jadval ma’lumotlari metodikaning asosiy afzalligini yaqqol ko‘rsatadi.'
D.replace(P, 'indeksi 2,83 ni', 'indeksi 2,81 ni', tag='3')
D.replace(P, 'uchun −0,44 ni', 'uchun −0,46 ni', tag='3')
P = 'Rasmdan ko‘rinib turibdiki, ikkala bankda ham indeks pasayishining asosiy sababi aktivlar sifati blokidir'
D.replace(P, 'ΔMBII ning 86–87 foizi', 'ΔMBII ning 83–86 foizi', tag='3')
D.replace('Metodikani ikki bankka qo‘llash MHXS ma’lumotlari asosida', 'ΔMBII ning 85 foizdan ortig‘i', 'ΔMBII ning 83–86 foizi', tag='3')
D.chart_set('/word/charts/chart25.xml', 1, [3.76, 2.28, 2.33, 3.83, 1.45], tag='3')
D.chart_set('/word/charts/chart26.xml', 0, [3.27, 2.81, 3.13, 2.47], tag='3')
D.chart_set('/word/charts/chart26.xml', 1, [3.06, 2.73, 3.07, 2.58], tag='3')
D.chart_set('/word/charts/chart26.xml', 2, [3.57, 3.02, 3.02, 2.13], tag='3')

# 1-ilova: SQB va MKB qatorlari (CET1, CAR, K bloki, MBII)
import mbii as M
def ilova1(caption, code, red_rows=True):
    t = D.table_after(caption)
    for j, y in enumerate(range(2021, 2026)):
        x = R32[code][str(y)]; i = x['ind']
        D.set_cell(t, 1, j + 1, f"{f2(i['CET1'],1)} ({f2(M.score(i['CET1'],1,[4.5,7,8,10.5]))})", tag='3')
        D.set_cell(t, 2, j + 1, f"{f2(i['CAR'],1)} ({f2(M.score(i['CAR'],1,[8,10.5,13,15.5]))})", tag='3')
        D.set_cell(t, 14, j + 1, f2(x['bloklar']['K']), tag='3')
        D.set_cell(t, 19, j + 1, f"{f2(x['mbii'])} ({M.cls(x['mbii'])})", tag='3')
ilova1('1.1. «O‘zsanoatqurilishbank» ATB', 'SQB')
ilova1('1.2. «Mikrokreditbank» ATB', 'MKB')
D.insert('Izoh: har bir katakda ko‘rsatkich qiymati va qavs ichida uning bali (1–4) keltirilgan', '',
         ' «O‘zsanoatqurilishbank» ATB va «Mikrokreditbank» ATB uchun CET1 va regulyativ kapital monandligi 2.3-paragrafdagi metodika bo‘yicha MHXS ma’lumotlari asosida qayta hisoblangan; qolgan banklar uchun chegirmalar oshkor qilinmagani sababli CET1 soddalashtirilgan formula (MHXS bo‘yicha xususiy kapital − ozchilik ulushi − joriy yil sof foydasi) bo‘yicha hisoblangan, regulyativ kapital monandligi esa hisobotda oshkor qilingan qiymatda olingan.', tag='3')

# 3.2-§ (dinamika matni) va 3.2.3-rasm
P = 'Rasmdan ko‘rinib turibdiki, besh yil davomida banklarning moliyaviy barqarorligi turlicha'
D.replace(P, '2,75–2,93 ball oralig‘ida', '2,72–2,92 ball oralig‘ida', tag='3')
D.replace(P, '2,20–2,47 ball oralig‘ida', '2,18–2,47 ball oralig‘ida', tag='3')
D.chart_set('/word/charts/chart29.xml', 3, [round(R32['SQB'][str(y)]['mbii'], 2) for y in range(2021, 2026)], tag='3')
D.chart_set('/word/charts/chart29.xml', 4, [round(R32['MKB'][str(y)]['mbii'], 2) for y in range(2021, 2026)], tag='3')
# 4.1-§
P = 'Rasmdan ko‘rinib turibdiki, bankning integral indeksi 2024-yil IV choragidagi yirik zarardan ancha oldin qizil zonada'
D.replace(P, 'butun davr mobaynida 2,20–2,47 ball', 'butun davr mobaynida 2,18–2,47 ball', tag='3')
P = 'Rasmdan ko‘rinib turibdiki, milliy hisobot ma’lumotlari asosidagi indeks 2021–2025-yillarda 3,17–3,61'
D.replace(P, 'butun davr mobaynida 2,75–2,93 ball bilan sariq zonada bo‘lgan, 2024-yilda esa qizil zona chegarasiga yetgan',
          '2021–2023 va 2025-yillarda 2,81–2,92 ball bilan sariq zonada bo‘lgan, 2024-yilda esa 2,72 ballgacha pasayib, qizil zonaga tushgan', tag='3')
# 4.1.3 va 4.1.4-rasmlar: MHXS indeksi (yillik nuqtalar)
D.chart_set('/word/charts/chart35.xml', 1, [round(R32['MKB'][str(y)]['mbii'], 2) for y in range(2021, 2026)], tag='3')
D.chart_set('/word/charts/chart36.xml', 1, [round(R32['SQB'][str(y)]['mbii'], 2) for y in range(2021, 2026)], tag='3')

# 4.2.6-jadval va 4.2.3-rasm (SQB MHXS)
t = D.table('Bank va hisobot | 2025 (haqiqiy)')
S = FC['banklar']['SQB']
vals = [S['2025']['MHXS'], S['Asosiy|2027']['MHXS'], S['Ijobiy|2027']['MHXS'], S['Salbiy|2027']['MHXS'], S['Keskin salbiy|2027']['MHXS']]
z = lambda m: 'yashil' if m >= 3.25 else 'sariq' if m >= 2.75 else 'qizil'
s3 = [S['2025']['S3_MHXS'], S['Asosiy|2027']['S3_MHXS'], S['Ijobiy|2027']['S3_MHXS'], S['Salbiy|2027']['S3_MHXS'], S['Keskin salbiy|2027']['S3_MHXS']]
for j, (m, s) in enumerate(zip(vals, s3)):
    D.set_cell(t, 2, j + 1, f'{f2(m)} ({z(m)}; {f2(s,1)} %)', tag='3')
D.chart_set('/word/charts/chart39.xml', 1, vals, tag='3')
P = 'Jadvaldan ko‘rinib turibdiki, ikki hisobot asosidagi prognozlar bir-biridan tubdan farq qiladi.'
D.replace(P, 'salbiy ssenariyda 2,71 ballgacha', 'salbiy ssenariyda 2,70 ballgacha', tag='3')
D.replace(P, 'keskin salbiy ssenariyda esa 2,61 ballgacha', 'keskin salbiy ssenariyda esa 2,59 ballgacha', tag='3')
# 4.3.2-jadval
t = D.table('Yashirin kredit riski koeffitsiyenti | MBII: milliy hisobot')
D.set_cell(t, 1, 5, '2,81', tag='3')
# III bob xulosalari
P = 'qadamni ifodalaydi' if False else 'imkoniyatidir. 2025-yil ma’lumotlariga ko‘ra'
D.replace(P, 'MHXS bo‘yicha 2,83 ballni', 'MHXS bo‘yicha 2,81 ballni', tag='3')
D.replace(P, 'ΔMBII −0,44 va −0,66', 'ΔMBII −0,46 va −0,66', tag='3')
D.replace(P, 'uning 86–87 foizi', 'uning 83–86 foizi', tag='3')
P = 'ATB ning indeksi 2,66 balldan 3,38 ballgacha izchil oshdi'
D.replace(P, '2,75–2,93 ball oralig‘ida', '2,72–2,92 ball oralig‘ida', tag='3')
D.replace(P, '(2,20–2,47 ball)', '(2,18–2,47 ball)', tag='3')
# Xulosa
P = 'o‘rtasidagi farq (ΔMBII) «O‘zsanoatqurilishbank» ATB da'
D.replace(P, 'ATB da −0,44', 'ATB da −0,46', tag='3')
D.replace(P, 'uning 86–87 foizi', 'uning 83–86 foizi', tag='3')

# 5-band: MBII-8 va MBII-13 belgilari
D.insert('Bog‘liq o‘zgaruvchi sifatida 3.1-paragrafda taklif etilgan integral indeksning choraklik',
         'mazmunan 3.1-paragrafdagi «MBII-milliy» variantiga mos keladi.',
         ' Ikki indeksni farqlash uchun keyingi o‘rinlarda 13 ko‘rsatkichli to‘liq indeks MBII-13, 8 ko‘rsatkichli panel indeksi esa MBII-8 deb belgilanadi.', tag='5')
t = D.table('Bank va hisobot | 2025 (haqiqiy)')
D.cell_replace(t, 1, 0, 'milliy hisobot', 'milliy hisobot (MBII-8)', tag='5')
D.cell_replace(t, 2, 0, 'MHXS hisoboti', 'MHXS hisoboti (MBII-13)', tag='5')
D.cell_replace(t, 3, 0, 'milliy hisobot', 'milliy hisobot (MBII-8)', tag='5')
D.cell_replace(t, 4, 0, 'MHXS hisoboti', 'MHXS hisoboti (MBII-13)', tag='5')
t = D.table('Yashirin kredit riski koeffitsiyenti | MBII: milliy hisobot')
D.cell_replace(t, 0, 4, 'MBII: milliy hisobot', 'MBII-8: milliy hisobot', tag='5')
D.cell_replace(t, 0, 5, 'MBII: MHXS hisoboti', 'MBII-13: MHXS hisoboti', tag='5')
P = 'Rasmdan ko‘rinib turibdiki, MHXS asosidagi indeks barcha ssenariylarda milliy indeksdan past'
D.replace(P, 'farq 0,78–0,88 ballni', 'farq 0,79–0,90 ballni', tag='3')
D.insert(P, 'barcha ssenariylarda milliy indeksdan past',
         ' (bu farqning bir qismi indekslar tarkibidagi farq bilan bog‘liq: milliy hisobot bo‘yicha MBII-8, MHXS bo‘yicha esa MBII-13 qo‘llangan; bir xil tarkibdagi MBII-13 bo‘yicha 2025-yilda farq «O‘zsanoatqurilishbank» ATB da 0,46 ballni (3,27 va 2,81), «Mikrokreditbank» ATB da 0,66 ballni (3,13 va 2,47) tashkil etadi va zona o‘zgarishi to‘g‘risidagi xulosa – yashildan sariqqa va sariqdan qizilga o‘tish – o‘z kuchida qoladi)', tag='5')

# 14-band: davlat banki belgisi vaqt bo'yicha o'zgarishi (3.3.3-jadval izohi)
D.replace('Qayd etilgan effektli modelda vaqt bo‘yicha o‘zgarmaydigan davlat banki', 'vaqt bo‘yicha o‘zgarmaydigan davlat banki o‘zgaruvchisi bank effektlari bilan to‘liq qamrab olinadi',
          'vaqt bo‘yicha deyarli o‘zgarmaydigan davlat banki o‘zgaruvchisi (u faqat «Ipoteka-bank» ATB da 2023-yil III choragida xususiylashtirish munosabati bilan o‘zgargan) bank effektlari bilan deyarli to‘liq qamrab olinadi va modeldan chiqarildi', tag='14')

# 15-band: 2.1.1-jadval – kapital talablari va buferlar (Markaziy bank manbasi bo'yicha)
t = D.table('Ko‘rsatkich | Talab | Bazel III minimumi | Izoh')
D.set_cell(t, 1, 1, '≥ 8% (2025-yilgacha); ≥ 4,5% + buferlar (2026-yildan)', tag='15')
D.set_cell(t, 1, 3, '2026-yildan buferlar bilan 8,5%, tizimli ahamiyatga molik banklar uchun 9,5%', tag='15')
D.set_cell(t, 2, 1, '≥ 13% (2025-yilgacha); ≥ 8% + buferlar (2026-yildan)', tag='15')
D.set_cell(t, 2, 3, '2025-yilgacha 3% lik saqlash buferini o‘z ichiga olgan; 2026-yildan buferlar bilan 12%, tizimli ahamiyatga molik banklar uchun 13%', tag='15')
D.set_cell(t, 3, 1, '≥ 6% (2025-yilgacha); ≥ 3% (2026-yildan)', tag='15')
D.set_cell(t, 3, 3, '2026-yildan Bazel III minimumiga tenglashtirildi', tag='15')
D.set_cell(t, 4, 1, '2,5% (CET1 hisobidan)', tag='15')
D.set_cell(t, 4, 3, '2026-yil 1-yanvardan (2025-yilgacha – 3%, I darajali kapital tarkibida)', tag='15')
D.add_row(t, 8, ['Yuqori likvid aktivlar / jami aktivlar', '≥ 10%', '–', 'Milliy prudensial me’yor'], tag='15')
D.insert('Jadvaldan ko‘rinadiki, O‘zbekistonda miqdoriy kapital talablari Bazel III minimumlaridan yuqori belgilangan', 'milliy normativlar ham saqlanib qolgan.',
         ' Jadvaldagi 2025-yilgacha amal qilgan talablar bankning prudensial hisobot shaklidagi belgilangan me’yorlar[[FN:' + FN['SQB22'] + ']], 2026-yildan amal qiladigan talablar va buferlar esa Markaziy bankning moliyaviy barqarorlik sharhi bo‘yicha keltirildi[[FN:' + FN['MBS25'] + ']]. 2026-yildan minimal talablar Bazel III darajasiga tushirilib, ularning ustiga buferlar qo‘shilganligi sababli buferlar bilan birga umumiy talab CET1 bo‘yicha 8,5 foizni, tizimli ahamiyatga molik banklar uchun 9,5 foizni tashkil etadi.', tag='15')
D.replace('Rasmdan ko‘rinadiki, salbiy ssenariyda kapital ko‘rsatkichlari amaldagi regulyativ minimumlardan', '(CET1 bo‘yicha 8 foiz, jami kapital bo‘yicha 13 foiz)',
          '(2025-yilda CET1 bo‘yicha 8 foiz va jami kapital bo‘yicha 13 foiz, 2026-yildan esa buferlar bilan birga mos ravishda 8,5–9,5 va 12–13 foiz)', tag='15')
D.insert('Jadval ma’lumotlari shuni ko‘rsatadiki, tahlil davrining barcha yillarida MHXS bo‘yicha xususiy kapital', 'ya’ni 2,5 barobar kichik.',
         ' 2026-yildan tizimli ahamiyatga molik banklar uchun buferlar bilan birga belgilangan 9,5 foizlik talabga nisbatan esa vaziyat tubdan farq qiladi: rasmiy hisobot bo‘yicha bank bu talabdan 0,71 foiz punkt yuqori, MHXS bo‘yicha esa 0,63 foiz punkt past turadi, ya’ni hisobot standarti bankning yangi bufer talablarini bajarish-bajarmasligi to‘g‘risidagi xulosani o‘zgartiradi.', tag='15')

# 16-band: 7-ilova – chegaraga yaqin qiymatlar uch xona aniqlikda
t = D.table_after('Tanlangan 25 ta bankning integral indeksi va erta ogohlantirish zonasi')
for ri, row in enumerate(t.rows):
    for ci, c in enumerate(row.cells):
        tx = c.text
        if tx.startswith('Garant bank') and False:
            pass
def fix7(bank, col, txt):
    for ri, row in enumerate(t.rows):
        if row.cells[0].text.startswith(bank):
            D.set_cell(t, ri, col, txt, tag='16')
            return
    raise KeyError(bank)
hdr = [c.text for c in t.rows[0].cells]
ci = {h: i for i, h in enumerate(hdr)}
print('7-ilova sarlavha:', hdr)
fix7('Garant bank', [i for h, i in ci.items() if 'Ijobiy' in h][0], '3,248 (sariq)')
fix7('Davr bank', [i for h, i in ci.items() if 'Ijobiy' in h][0], '3,251 (yashil)')
fix7('Tenge bank', [i for h, i in ci.items() if h.startswith('Salbiy') or h == 'Salbiy'][0], '3,252 (yashil)')
fix7('Xalq banki', [i for h, i in ci.items() if 'Keskin' in h][0], '3,254 (yashil)')
D.insert('Izoh: (D) – davlat ishtirokidagi bank; qavs ichida erta ogohlantirish zonasi.', '',
         ' Zonalar yaxlitlanmagan qiymatlar bo‘yicha aniqlangan; 3,25 ballik chegaraga yaqin qiymatlar uch xona aniqlikda berilgan.', tag='16')

# 17-band: 2.2.1-jadval izohi
D.replace('Izoh: «–» – byulletenda ma’lumot berilmagan.', '«–» – byulletenda ma’lumot berilmagan.',
          'substandart va muammoli kreditlarning umumiy ulushi 2024–2025-yillar uchun manbada butun songacha yaxlitlangan holda e’lon qilingan.', tag='17')

# 19 va 28-bandlar: satellit model koeffitsiyentlari – indikativ, Nickell siljishi
D.insert('Sakkiz chorak lag sababli model 2024-yil I choragidan', 'kuzatuvlar asosida baholandi.',
         ' Makroiqtisodiy omillar barcha banklar uchun bir xil bo‘lgani sababli ularning koeffitsiyentlari amalda sakkizta vaqt nuqtasi asosida aniqlanadi va indikativ xususiyatga ega; prognozda 4.2-paragrafda uzoqroq qator bo‘yicha baholangan koeffitsiyent qo‘llanadi. Qayd etilgan effektli modelda lagli bog‘liq o‘zgaruvchi qisqa panelda (T = 8) S. Nickell siljishiga olib keladi[[FN:' + FN['NICKELL'] + ']]: bunday modelda inersiya koeffitsiyenti pastga, Pooled modelda esa yuqoriga siljigan bo‘ladi, shu sababli uning haqiqiy qiymati 0,62–0,81 oralig‘ida joylashadi deb baholanadi.', tag='19,28')

# =====================================================================
# B. ILMIY-METODOLOGIK KAMCHILIKLAR
# =====================================================================
# 23-band: klasterlangan bootstrap, muqobil hodisa ta'riflari, qoplanmaydigan oynalar (4.1-§)
H5, H3, H10, Q4, MZ = RB['hodisa_5'], RB['hodisa_3'], RB['hodisa_10'], RB['IV_chorak'], RB['farq_MBII_Z']
c = lambda v: f'{f2(v[0])}–{f2(v[1])}'
txt23 = ('Natijalarning barqarorligi qo‘shimcha ravishda uch yo‘nalishda tekshirildi. Birinchidan, bir bankka tegishli kuzatuvlar o‘zaro bog‘liq va to‘rt choraklik bashorat oynalari bir-birini qoplagani sababli ishonch oraliqlari bank bo‘yicha klasterlangan bootstrap usulida, ya’ni tanlanmaga banklar to‘liq holda qaytariladigan tanlov bilan olinib, 2000 marta qayta hisoblandi[[FN:' + FN['CGM'] + ']]. Bunda integral indeks AUC qiymatining 95 foizlik ishonch oralig‘i ' + c(H5['CI_MBII']) + ' ni tashkil etib kengaydi, biroq quyi chegarasi tasodifiy bashorat darajasidan (0,5) yuqori qoldi. Z-score bilan umumiy kuzatuvlarda integral indeksning ustunligi klasterlangan bootstrapda ham tasdiqlandi: AUC farqining ishonch oralig‘i ' + c(MZ['CI']) + ' ni tashkil etib, noldan katta. Muammoli kreditlar ulushi bilan farq esa ' + f2(H5['farq_MBII_S3_CI'][0]) + ' dan +' + f2(H5['farq_MBII_S3_CI'][1]) + ' gacha oraliqda bo‘lib, statistik ahamiyatsiz. Ikkinchidan, muammo hodisasining muqobil ta’riflari qo‘llandi: chorak zararining kapitalga nisbati chegarasi 3 foiz olinganda (' + str(H3['banklar']) + ' ta bankda ' + str(H3['hodisalar']) + ' ta hodisa) indeksning AUC qiymati ' + f2(H3['AUC_MBII']) + ' ni (' + c(H3['CI_MBII']) + '), 10 foiz olinganda (' + str(H10['banklar']) + ' ta bankda ' + str(H10['hodisalar']) + ' ta hodisa) ' + f2(H10['AUC_MBII']) + ' ni tashkil etdi, Z-score niki esa barcha ta’riflarda ' + f2(min(H3['AUC_Z'], H10['AUC_Z'], H5['AUC_Z'])) + '–' + f2(max(H3['AUC_Z'], H10['AUC_Z'], H5['AUC_Z'])) + ' oralig‘ida qoldi. Uchinchidan, bir-birini qoplaydigan oynalar ta’sirini bartaraf etish uchun har bir bank bo‘yicha faqat IV chorak kuzatuvlari olinganda (' + str(Q4['N']) + ' ta kuzatuv, ' + str(Q4['pos']) + ' tasi hodisadan oldin) indeksning AUC qiymati ' + f2(Q4['AUC_MBII']) + ' ni (' + c(Q4['CI_MBII']) + ') tashkil etdi[[FN:' + FN['ROBUST'] + ']]. Demak, indeksning bashorat kuchi hodisa ta’rifi va baholash usuliga nisbatan barqaror, Z-score dan ustunligi statistik ahamiyatli, biroq muammoli kreditlar ulushiga nisbatan bashorat ustunligi tasdiqlanmaydi. Hodisalar soni kamligi sababli ishonch oraliqlari keng bo‘lib, natijalar yangi ma’lumotlar to‘planishi bilan qayta tekshirilishi lozim.')
D.new_para_after('Natijalarning ishonchliligini tekshirish uchun vaqtdan tashqari sinov o‘tkazildi', txt23, tag='23')
D.insert('Izoh: ishonch oralig‘i 2000 ta bootstrap tanlanma asosida hisoblangan.', 'bootstrap tanlanma asosida hisoblangan',
         ' (kuzatuvlar bo‘yicha); bank bo‘yicha klasterlangan bootstrap natijalari matnda keltirilgan', tag='23')
D.replace('integral indeksga asoslangan uch zonali erta ogohlantirish tizimi ishlab chiqilgan va indeksning (ochiq', '25 ta bank paneli ma’lumotlari (10 ta muammo hodisasi) asosida empirik jihatdan ko‘rsatilgan',
          '25 ta bank paneli ma’lumotlari (10 ta muammo hodisasi) asosida empirik jihatdan ko‘rsatilgan (bank bo‘yicha klasterlangan bootstrap 95 foizlik ishonch oralig‘i ' + c(H5['CI_MBII']) + ', Z-score dan ustunligi statistik ahamiyatli)', tag='23')

# 24-band: I bob xulosalari va 1.2-§ dagi ishchi ta'rif – bloklar tarkibini muvofiqlashtirish
D.insert('Muallif tomonidan moliyaviy barqarorlikni baholashning konseptual modeli taklif etildi.', 'erta ogohlantirish signallari bilan bog‘lanadi.',
         ' Boshqaruv sifatini ochiq ma’lumotlar asosida miqdoriy baholab bo‘lmagani sababli III bobda ishlab chiqilgan integral indeksda u alohida blok sifatida emas, balki kuratorlarning asoslangan mulohazasi orqali hisobga olinadi, uning o‘rniga dollarlashuv sharoitida muhim bo‘lgan valyuta riskiga sezgirlik bloki kiritiladi.', tag='24')
D.insert('Yuqoridagilarga asoslanib, muallif tomonidan tijorat banki moliyaviy barqarorligini baholashning quyidagi ishchi ta’rifi', 'bozor risklariga sezgirlik',
         ' (O‘zbekiston sharoitida uning asosiy shakli – valyuta riskiga sezgirlik)', tag='24')

# 25-band: 3.2-§ – empirik o'rtacha faqat empirik usullardan, variatsiya koeffitsiyenti ustuni, xulosa
W = VZ['vaznlar']
t = D.table('Blok | Ekspert vaznlari (3.1-§)')
for r, b in enumerate(['K', 'A', 'L', 'D', 'V'], start=1):
    D.set_cell(t, r, 6, f2(W['O‘rtacha (empirik)'][b]), tag='25')
D.replace('Izoh: * – logit usulida baholanmagan, ekspert vazni saqlangan.', 'ular bloklar dispersiyasining 66,0 foizini izohlaydi',
          'ular bloklar dispersiyasining 66,7 foizini izohlaydi', tag='3,25')
D.insert('Izoh: * – logit usulida baholanmagan, ekspert vazni saqlangan.', '',
         ' Empirik usullar o‘rtachasi faqat barcha bloklarni empirik baholagan uch usul – entropiya, variatsiya koeffitsiyenti va bosh komponentlar usullari bo‘yicha hisoblangan; kapital va valyuta bloklarini baholamagan logit usuli o‘rtachaga kiritilmadi va bashorat jihatidan tekshiruv sifatida qo‘llanadi.', tag='25')
P = 'Jadvaldan ko‘rinib turibdiki, barcha empirik usullar aktivlar sifati blokiga yuqori vazn bergan'
D.replace(P, 'Demak, aktivlar sifati blokiga eng yuqori vazn berish to‘g‘risidagi ekspert qarori empirik jihatdan tasdiqlanadi.',
          'Demak, aktivlar sifati blokining ekspert vazni (0,30) empirik baholar oralig‘ida (0,27–0,57) joylashadi va uchta tarqoqlik usulining o‘rtachasiga (0,31) deyarli teng. Biroq tarqoqlikka asoslangan usullarda eng yuqori vazn daromadlilik blokiga (0,30–0,46) berilgan, ya’ni ularda aktivlar sifati bloki ikkinchi o‘rinda turadi; uning yetakchiligi faqat bashorat kuchiga asoslangan logit usulida (0,57) namoyon bo‘ladi.', tag='25')
D.replace(P, 'barcha empirik usullar aktivlar sifati blokiga yuqori vazn bergan', 'barcha empirik usullar aktivlar sifati blokiga ekspert vazniga yaqin yoki undan yuqori vazn bergan', tag='25')
D.chart_set('/word/charts/chart27.xml', 4, [W['O‘rtacha (empirik)'][b] for b in 'KALDV'], tag='25')
# 3.2.3-jadval: variatsiya koeffitsiyenti ustuni va yangi qiymatlar
t = D.table('Ko‘rsatkich | Ekspert vaznlari | Entropiya usuli')
BK = ['SQB', 'MKB', 'BRB', 'XALQ', 'IPOTEKA']
BM = VZ['bank_mbii']
cv = ['Variatsiya koeffitsiyenti usuli'] + [f2(BM['Variatsiya'][f'{b}|2025']) for b in BK] + [str(VZ['sinf_mos']['Variatsiya']), f2(VZ['spearman']['Variatsiya']), f2(RB['AUC_vaznlar']['Variatsiya'], 3)]
D.add_col(t, 2, cv, tag='25')
cols = {'Ekspert': 1, 'Bosh komponentlar': 4, 'Logit (K, V – ekspert)': 5, 'O‘rtacha (empirik)': 6, 'Entropiya': 2}
for nm, ci in cols.items():
    for r, b in enumerate(BK, start=1):
        old = t.rows[r].cells[ci].text; new = f2(BM[nm][f'{b}|2025'])
        if old != new:
            D.set_cell(t, r, ci, new, tag='3,25')
    for r, val in ((6, str(VZ['sinf_mos'][nm])), (7, f2(VZ['spearman'][nm])), (8, f2(RB['AUC_vaznlar'][nm], 3))):
        if t.rows[r].cells[ci].text != val:
            D.set_cell(t, r, ci, val, tag='25')
P = 'Jadvaldan ko‘rinib turibdiki, vaznlar tanlovi indeksning mutlaq qiymatiga ta’sir qiladi'
D.replace(P, 'korrelyatsiya koeffitsiyenti 0,95–0,99 ni', 'korrelyatsiya koeffitsiyenti 0,95–0,98 ni', tag='25')
D.insert(P, 'ekspert vaznlarining bashorat kuchi esa bosh komponentlar usuli bilan bir xil.',
         ' Faqat empirik usullar asosidagi o‘rtacha vaznlar bilan indeksning AUC qiymati 0,871 ni tashkil etib, ekspert vaznlaridagidan (0,864) atigi 0,007 ga yuqori, ya’ni vaznlarni empirik usulda almashtirish bashorat kuchini sezilarli oshirmaydi.', tag='25')
P = 'Rasmdan ko‘rinib turibdiki, vaznlar ekspert vaznlari atrofida o‘zgarganda'
D.replace(P, '91,9 foizga', '91,7 foizga', tag='3')
D.replace(P, 'esa 0,85 ni', 'esa 0,84 ni', tag='3')
D.replace(P, 'o‘rtacha 55,7 foiz', 'o‘rtacha 55,8 foiz', tag='3')
MC1, MC2 = VZ['mc_ekspert_atrofida']['bank_yil'], VZ['mc_bir_tekis']['bank_yil']
import statistics as st
mb = lambda d, b: round(st.mean(v for k, v in d.items() if k.startswith(b + '|')), 1)
D.chart_set('/word/charts/chart28.xml', 0, [mb(MC1, b) for b in BK] + [VZ['mc_ekspert_atrofida']['sinf_saqlanishi_ortacha']], tag='3')
D.chart_set('/word/charts/chart28.xml', 1, [mb(MC2, b) for b in BK] + [VZ['mc_bir_tekis']['sinf_saqlanishi_ortacha']], tag='3')
P = 'Shunday qilib, empirik tahlil 3.1-paragrafda taklif etilgan integral indeksning asosiy metodik elementlarini tasdiqladi.'
D.replace(P, 'Birinchidan, barcha empirik usullar aktivlar sifati blokiga eng yuqori yoki ekspert vazniga teng vazn berdi, ya’ni metodikaning markaziy elementi empirik jihatdan asoslandi.',
          'Birinchidan, aktivlar sifati blokining ekspert vazni barcha empirik baholar oralig‘ida (0,27–0,57) joylashadi va tarqoqlik usullarining o‘rtachasiga (0,31) deyarli teng, bashorat kuchiga asoslangan logit usulida esa bu blok yagona statistik ahamiyatli blok bo‘ldi.', tag='25')
D.replace(P, 'tartib korrelyatsiyasi 0,95–0,99 ni', 'tartib korrelyatsiyasi 0,95–0,98 ni', tag='25')
D.replace(P, 'Uchinchidan, barqarorlik sinflari chegaralari muammo ehtimolidagi sifat o‘zgarishiga mos keladi',
          'Uchinchidan, barqarorlik sinflari chegaralari tahlil davri ma’lumotlarida muammo ehtimolidagi sifat o‘zgarishiga mos keladi', tag='27')
D.insert(P, 'mos ravishda 2,0; 13,3 va 90,0 foizni tashkil etdi', ' (95 foizlik ishonch oraliqlari: 0,8–5,0; 7,9–21,4 va 59,6–98,2 foiz)', tag='27')
# III bob xulosalari va umumiy xulosa
P = 'Indeks vaznlarini to‘rtta empirik usul'
D.replace(P, 'barcha usullarda aktivlar sifati blokining yetakchi ahamiyatini tasdiqladi (vazni 0,27–0,57)',
          'aktivlar sifati blokining ekspert vazni (0,30) empirik baholar oralig‘ida (0,27–0,57) joylashishini va bu blok bashorat jihatidan yagona ahamiyatli blok ekanini ko‘rsatdi, tarqoqlik usullari esa daromadlilik blokiga eng yuqori vazn berdi', tag='25')
D.replace(P, 'korrelyatsiya 0,95–0,99 ni', 'korrelyatsiya 0,95–0,98 ni', tag='25')
D.replace(P, 'Barqarorlik sinflari chegaralari ham empirik jihatdan tasdiqlandi', 'Barqarorlik sinflari chegaralari tahlil davri ma’lumotlariga mos keldi', tag='27')
D.insert(P, 'mos ravishda 2,0; 13,3 va 90,0 foizni tashkil etdi', ' (95 foizlik ishonch oraliqlari: 0,8–5,0; 7,9–21,4 va 59,6–98,2 foiz)', tag='27')
P = 'banklar tartibi esa vaznlarning o'
D.replace(P, 'Indeks vaznlari va sinflar chegaralari empirik jihatdan asoslandi: entropiya, variatsiya koeffitsiyenti, bosh komponentlar va logit usullarining barchasi aktivlar sifati blokining yetakchi ahamiyatini tasdiqladi',
          'Indeks vaznlari va sinflar chegaralari empirik jihatdan tekshirildi: aktivlar sifati blokining ekspert vazni entropiya, variatsiya koeffitsiyenti, bosh komponentlar va logit usullari bo‘yicha olingan baholar oralig‘ida (0,27–0,57) joylashadi, bashorat jihatidan esa bu blok yagona ahamiyatli blok hisoblanadi', tag='25')
D.replace(P, '(tartib korrelyatsiyasi 0,95–0,99)', '(tartib korrelyatsiyasi 0,95–0,98)', tag='25')
D.insert(P, 'mos ravishda 2,0; 13,3 va 90,0 foizni tashkil etdi', ' (95 foizlik ishonch oraliqlari: 0,8–5,0; 7,9–21,4 va 59,6–98,2 foiz)', tag='27')

# 26-band: chegaralarni asoslash – L1 manbasi, A2 tashqi tekshiruvi, sinflar yarim ochiq oraliqlar
P = 'Aktivlar sifati blokida 3-bosqich kreditlari ulushi uchun 5 foiz chegarasi'
D.insert(P, 'chegaralar ikkinchi bobda aniqlangan qiymatlar oralig‘i asosida belgilandi.',
         ' Bu chegaralar tashqi ma’lumot bilan ham tekshirildi: Yevropa bank boshqarmasi ma’lumotlariga ko‘ra, Yevropa Ittifoqi va Yevropa iqtisodiy hududi banklarida 2-bosqich kreditlari ulushi 2023-yil oxirida 9,6 foizni, 2024-yil oxirida 9,7 foizni tashkil etgan, mamlakatlar kesimida esa eng yuqori qiymat 15,9 foizga teng bo‘lgan[[FN:' + FN['EBARD25'] + ']]. Demak, 4 ball chegarasi (10 foiz) Yevropa banklarining o‘rtacha darajasiga, 3 ball chegarasi (15 foiz) esa ularning eng yuqori darajasiga mos keladi, 25 va 35 foizlik chegaralar esa Yevropa amaliyotida mamlakat darajasida kuzatilmaydigan, ya’ni sezilarli kredit riski to‘planganini ko‘rsatuvchi darajalardir.', tag='26')
D.replace(P, 'uning 2 ball chegarasi sifatida Markaziy bank tomonidan belgilangan 10 foizlik minimal talab qabul qilindi',
          'uning 2 ball chegarasi sifatida Markaziy bank tomonidan aynan shu ko‘rsatkich – «yuqori likvid aktivlarning jami aktivlardagi ulushi» prudensial me’yori uchun belgilangan 10 foizlik minimal talab qabul qilindi (u lahzali likvidlik normativi bilan bir xil emas)[[FN:' + FN['SQB22'] + ']]', tag='26')
t = D.table('Sinf | MBII qiymati | Tavsifi')
D.set_cell(t, 1, 1, '3,25 ≤ MBII ≤ 4,00', tag='26')
D.set_cell(t, 2, 1, '2,50 ≤ MBII < 3,25', tag='26')
D.set_cell(t, 3, 1, '1,75 ≤ MBII < 2,50', tag='26')
D.set_cell(t, 4, 1, '1,00 ≤ MBII < 1,75', tag='26')
D.insert('Integral indeks qiymati 1 dan 4 gacha bo‘lgan oraliqda bo‘ladi.', 'nazorat choralari belgilandi (3.1.4-jadval).',
         ' Teng oraliqlar sinflar chegaralarini natijalarga moslashtirmasdan, oldindan belgilash imkonini beradi, bu esa ularni keyinchalik mustaqil ravishda empirik tekshirishga (3.2-§) imkon yaratadi.', tag='26')
t = D.table('Barqarorlik sinfi | Integral indeks, ball')
D.set_cell(t, 1, 1, '3,25 ≤ MBII', tag='26'); D.set_cell(t, 2, 1, '2,50 ≤ MBII < 3,25', tag='26')
D.set_cell(t, 3, 1, '1,75 ≤ MBII < 2,50', tag='26'); D.set_cell(t, 4, 1, 'MBII < 1,75', tag='26')

# 27-band: Wilson oraliqlari (3.2.4-jadval va matn)
WL = RB['wilson']
D.add_col(t, 4, ['95 % ishonch oralig‘i (Wilson), %'] + [f"{f2(WL[s]['wilson'][0],1)}–{f2(WL[s]['wilson'][1],1)}" if WL[s]['wilson'] else '–' for s in ('barqaror', 'qoniqarli', 'zaif', 'nobarqaror')], tag='27')
D.insert('Jadvaldan ko‘rinib turibdiki, sinflar chegaralari real ma’lumotlarda aniq farqlanadi', 'III sinfda esa 90,0 foizni tashkil etgan.',
         ' Bu ulushlarning 95 foizlik Wilson ishonch oraliqlari[[FN:' + FN['WILSON'] + ']] mos ravishda 0,8–5,0; 7,9–21,4 va 59,6–98,2 foizni tashkil etib, bir-biri bilan kesishmaydi, Fisherning aniq testi bo‘yicha esa I va II sinflar hamda II va III sinflar o‘rtasidagi farqlar statistik ahamiyatli (har ikki holatda p < 0,001). Shu bilan birga, III sinfda atigi 10 ta kuzatuv (5 ta bank) mavjudligi va IV sinfda kuzatuv yo‘qligi sababli bu natija chegaralarning yakuniy tasdig‘i emas, balki ularning tahlil davri ma’lumotlariga mosligini ko‘rsatadi.', tag='27')
D.replace('Jadvaldan ko‘rinib turibdiki, sinflar chegaralari real ma’lumotlarda aniq farqlanadi', 'sinflar chegaralari real ma’lumotlarda aniq farqlanadi',
          'sinflar chegaralari tahlil davri ma’lumotlarida aniq farqlanadi', tag='27')

# 29-band: H2 – 8 chorak lag bilan baholash
D.insert('Kreditlar o‘sishining ta’siri spetsifikatsiyaga bog‘liq bo‘lgan.', 'integral indeksda aks etmaydi.',
         ' Ikkinchi gipotezada ko‘zda tutilgan sakkiz chorak lag bilan baholash (194 ta kuzatuv, 2024–2025-yillar) ham shu manzarani beradi: Pooled modelda sakkiz chorak oldingi kreditlar o‘sishi indeksni 1 foiz darajada ahamiyatli pasaytiradi (−0,0007), bank ichida esa salbiy ta’sir aniqlanmadi (qayd etilgan effektli modelda koeffitsiyent +0,0005, p = 0,05), Z-score bo‘yicha esa sakkiz chorak lagli ta’sir barcha modellarda ahamiyatsiz. Demak, jadal kreditlashning Z-score ga ta’siri asosan bir yil ichida namoyon bo‘ladi.', tag='29')
t = D.table('Gipoteza | Integral indeks (MBII) | Z-score')
D.cell_replace(t, 2, 1, 'Pooled modelda manfiy***, bank ichida ahamiyatsiz', 'Pooled modelda manfiy*** (4 va 8 chorak lag), bank ichida ahamiyatsiz', tag='29')
D.cell_replace(t, 2, 2, 'barcha modellarda manfiy (**/***)', 'barcha modellarda manfiy (**/***) (4 chorak lag); 8 chorak lag – ahamiyatsiz', tag='29')

# 30-band: L1 – milliy balansdan (yuqorida 3.1.5-jadval izohiga qo'shildi), metodika cheklovlari
D.insert('Jadvaldan ko‘rinib turibdiki, taklif etilayotgan metodika mavjud yondashuvlarning afzalliklarini birlashtiradi',
         'ularga yaqin ko‘rsatkichlar bilan baholanadi',
         ', MHXS variantida esa yuqori likvid aktivlar ulushi milliy balansdan olinadi', tag='30')

# 31-band: mualliflik ta'rifi – ilmiy yangilik qismiga ko'chirildi
src = D.para('tijorat banklari moliyaviy barqarorligining kapital, aktivlar sifati, likvidlik va daromadlilik o‘rtasidagi muvozanatni')
D.new_para_after('Tadqiqotning ilmiy yangiligi quyidagilardan iborat:',
                 'tijorat banklari moliyaviy barqarorligining kapital, aktivlar sifati, likvidlik va daromadlilik o‘rtasidagi muvozanatni hamda xalqaro standartlarda belgilangan chegaralarga nisbatan zaxirani hisobga oluvchi mualliflik ta’rifi ishlab chiqilgan;',
                 tag='31', style_from=src)
D.delete_para(src, tag='31')

# 32-band: LCR ni Bazel III bo'yicha qayta hisoblash (2.3.4-jadval va matn)
t = D.table_after('2.3.4-jadval')
D.add_row(t, len(t.rows) - 1, ['Likvidlik', 'Likvidlikni qoplash koeffitsiyenti, %', '222,9', '154,1', '−68,8 p.p. (banklardagi mablag‘lar yuqori likvid aktivlardan chiqarilganda)'], tag='32')
D.replace('Ammo likvidlikni qoplash koeffitsiyenti tarkibida xalqaro standart bo‘yicha yuqori likvid aktiv hisoblanmaydigan mablag‘lar ulushi 30,9 foizni tashkil etadi, demak, bu blokda ham yashirin risk mavjud.',
          'demak, bu blokda ham yashirin risk mavjud.',
          'Ularni Bazel III talabiga muvofiq yuqori likvid aktivlar tarkibidan chiqarib, kelgusi 30 kunlik sof chiqimni o‘zgarishsiz qoldirgan holda qayta hisoblanganda likvidlikni qoplash koeffitsiyenti 222,9 foizdan 154,1 foizgacha, ya’ni 68,8 foiz punktga pasayadi[[FN:' + FN['SQB16'] + ']]. Bazel III garovga qo‘yilgan aktivlarni ham yuqori likvid aktivlar tarkibiga kiritmasligini hisobga olsak (598,3 mlrd so‘m), koeffitsiyent 146,5 foizgacha pasayadi. Koeffitsiyent minimal talabdan yuqori qolsa-da, likvidlik buferining qariyb uchdan bir qismi xalqaro standart bo‘yicha bufer hisoblanmaydi, demak, bu blokda ham yashirin risk mavjud.', tag='32')

# 33-band: ekspert baholash usuli aniqlashtirildi
D.replace('Tadqiqotning usullari.', 'ekspert baholash usullaridan',
          'ekspert baholash (vaznlarni muallif tomonidan belgilash va ularni entropiya, variatsiya koeffitsiyenti, bosh komponentlar va logit usullari bilan tekshirish), bootstrap va Monte-Karlo sezgirlik tahlili usullaridan', tag='33')

# 35-band: ma'lumotlardan foydalanish va takrorlanuvchanlik
D.insert('Har ikki bankning 2021–2025-yillar yakunidagi Markaziy bankka taqdim etilgan prudensial hisobot shakllari', 'moliyaviy hisobotlari bilan solishtirildi.',
         ' Prudensial hisobot shakllari ochiq e’lon qilinmaydi va tadqiqotda ulardan banklarning ruxsati bilan faqat ilmiy maqsadda foydalanildi, shu sababli ularga asoslangan hisob-kitoblarni (kapital monandligi, NSFR va LCR ni qayta hisoblash) dastlabki ma’lumotlar bilan takrorlash uchun bu shakllarga kirish huquqi talab etiladi. Ochiq ma’lumotlarga asoslangan qism – MHXS bo‘yicha moliyaviy hisobotlar, Markaziy bankning banklar kesimidagi statistikasi va «SNS» reyting agentligi bazasi asosidagi integral indeks hamda uning 25 ta bank bo‘yicha panel varianti – esa har qanday tadqiqotchi tomonidan to‘liq takrorlanishi mumkin.', tag='35')

# 36-band: ssenariylarga kapital kanalini qo'shish (4.2-§)
KK = json.load(open(HK + 'kapkanal_kam.json'))
D.insert('Prognozni talqin qilishda uning cheklovlarini hisobga olish zarur.', 'natijalarni yanada salbiy tomonga o‘zgartiradi.',
         ' Bu cheklovning ta’sirini baholash uchun ssenariylarga kapital kanali qo‘shilib, ya’ni qo‘shimcha zaxira ajratmalari natijasida aktivlar rentabelligining pasayishi har yili kapitalni kamaytiradi deb faraz qilingan holda hisob-kitob takrorlandi. Bunda integral indeks ko‘pi bilan ' + f2(max(v['max_farq'] for v in KK.values())) + ' ballga pasaydi, keskin salbiy ssenariyda 2027-yilda qizil zonadagi banklar soni o‘zgarmadi (' + str(KK['Keskin salbiy|2027']['qizil_kapital']) + ' ta), faqat salbiy ssenariyda 2028-yilda «Mikrokreditbank» ATB sariq zonadan qizil zonaga o‘tdi (2,755 dan 2,747 ballgacha). Demak, natijalar kapital kanaliga nisbatan barqaror, biroq likvidlik kanali hisobga olinmagani sababli ular shok ta’sirining quyi bahosi hisoblanadi va prognoz aniq bashorat emas, balki sezgirlik (stress-ssenariy) tahlili sifatida talqin qilinishi lozim.', tag='36')

# 37/38 – taqrizchi tuzatgan; 1.3 va 4.3: FFIEC (S&C sharhi) va EBA risk dashbordi
D.insert('AQSh baholash tizimi hozir ham o‘zgarishda davom etmoqda.', 'mavjud bo‘lgandagina qo‘yiladi',
         None or '', tag='x') if False else None
P = 'AQSh baholash tizimi hozir ham o‘zgarishda davom etmoqda.'
p = D.para(P)
i = p.text.find('Stress-testlar ham isloh qilindi')
D.insert(P, p.text[:i].rstrip()[-40:],
         ' Federal rezerv tizimining nazorat bo‘yicha rais o‘rinbosari M. Bowman bu o‘zgarishni nazoratning shaffoflik, miqdoriy omillar va bashorat qilinuvchanlik tomon qat’iy burilishi deb baholagan, taklif bo‘yicha fikrlar esa 2026-yil 17-avgustgacha qabul qilinadi[[FN:' + FN['SC26'] + ']].', tag='SC26')
D.insert('Ikki darajani bog‘lashda yagona ma’lumotlar bazasi muhim rol o‘ynaydi.', '53 ta asosiy risk indikatorini yig‘ib keladi',
         ' va ular asosida bank sektorining risk dashbordini muntazam e’lon qiladi[[FN:' + FN['EBARD'] + ']]', tag='EBARD')
D.insert('Markaziy bank uchun birinchi taklif integral indeksni masofaviy monitoring amaliyotiga joriy etishdan iborat.', 'uni sifat omillari bilan to‘ldiradi.',
         ' Bu yondashuv xalqaro amaliyotdagi yangi tendensiyaga ham mos keladi: AQShda 2026-yil mayida e’lon qilingan CAMELS reyting tizimini qayta ko‘rib chiqish taklifi baholashni bankning moliyaviy holatiga sezilarli ta’sir ko‘rsatuvchi miqdoriy omillarga yo‘naltirish va reytinglar shaffofligini oshirishni maqsad qilgan[[FN:' + FN['SC26'] + ']].', tag='SC26')
D.insert('To‘rtinchi taklif oshkoralikni kengaytirishdan iborat.', 'qayta tuzilgan va imtiyozli kreditlar ulushi.',
         ' Bundan tashqari, Markaziy bank Yevropa bank boshqarmasining risk dashbordi tajribasidan foydalanib[[FN:' + FN['EBARD'] + ']], bank tizimi bo‘yicha integral indeks va uning bloklari natijalarini guruhlar kesimida choraklik e’lon qilishi mumkin.', tag='EBARD')
# FSSA havolasi: birinchi izoh
D.insert(D.fn_para(10), 'Washington, DC: IMF, 2025.', FSSA_URL[:-1], tag='FSSA')
# Kilde blogi o'rniga rasmiy manba (2.2-§, aktivlar va kreditlar)
fp = D.fn_para(contains='Kilde')
D.replace(fp, ''.join(x.text or '' for x in fp._p.iter(K.qn('w:t'))).strip(), FN['MBS25A'], tag='47')

# =====================================================================
# C. TIL, MANTIQ VA RASMIYLASHTIRISH
# =====================================================================
# 39-band: 2.1-§ tuzilishini boshqa paragraflarga moslashtirish
for h in ('1. Baholashning me’yoriy-huquqiy va institutsional asoslari', '2. Amaldagi baholash tizimining tarkibi',
          '3. Bank tizimi moliyaviy barqarorligining asosiy ko‘rsatkichlari', '4. Xalqaro standartlarga moslashtirish muammolari', 'Paragraf bo‘yicha xulosa'):
    D.delete_para(D.para(h), tag='39')
P = 'O‘zbekistonda banklar moliyaviy barqarorligini baholash uch darajada – prudensial normativlar'
D.replace(P, 'O‘zbekistonda banklar moliyaviy barqarorligini baholash uch darajada',
          'O‘zbekistonda banklar moliyaviy barqarorligini baholashning amaldagi holati va xalqaro standartlarga moslashtirish muammolarini tadqiq etgan holda xulosa qilish mumkinki, baholash uch darajada', tag='39')

# 45-band: Kirishdagi to'ldirilmagan joylar – muallif to'ldiradi (qizil bilan belgilandi)
for s in ('[ilmiy-tadqiqot mavzusi nomi]',):
    D.replace(D.para(s), s, s, tag='45')

# 47-band: adabiyotlar ro'yxati
D.replace('MDH mamlakatlari olimlaridan O.I.Lavrushin', 'I.V.Larionova, V.V.Ivanov, O.N.Afanasyeva', 'I.V.Larionova, O.N.Afanasyeva', tag='47')
D.insert('70. Регулирование банковской сферы: учебник', '– М.: КНОРУС.', ' – [nashr yilini kiriting]', tag='47')
D.insert('71. Соколинская Н.Э. Анализ и оценка', '– М.: РУСАЙНС.', ' – [nashr yilini kiriting]', tag='47')
D.insert('75. Шаталова Е.П. Банковские рейтинги', 'учебное пособие.', ' – [shahar va nashriyotni kiriting]', tag='47')
p = D.para('164. Non-Bank Financial Institutions & Private Credit in Uzbekistan')
D.replace(p, p.text[5:], 'O‘zbekiston Respublikasi Markaziy banki. Moliyaviy barqarorlik sharhi: 2025-yil. – Toshkent, 2026.', tag='47')
D.insert('37. International Monetary Fund. Republic of Uzbekistan: Financial Sector Assessment Program – Detailed', '', '', tag='') if False else None
for s in ('169. https://www.bis.org', '170. https://www.imf.org', '171. https://www.worldbank.org', '172. https://www.ifrs.org', '173. https://cbu.uz', '174. https://lex.uz'):
    D.delete_para(D.para(s), tag='47')
# yangi normativ hujjatlar (I bo'lim oxiriga)
NEW_I = [
    'O‘zbekiston Respublikasi Prezidentining 2023-yil 11-sentabrdagi PF-158-son «“O‘zbekiston – 2030” strategiyasi to‘g‘risida»gi Farmoni // Qonunchilik ma’lumotlari milliy bazasi.',
    'O‘zbekiston Respublikasi Prezidentining 2015-yil 6-maydagi PQ-2344-son «Tijorat banklarining moliyaviy barqarorligini yanada oshirish va ularning resurs bazasini rivojlantirish chora-tadbirlari to‘g‘risida»gi qarori // O‘zbekiston Respublikasi qonun hujjatlari to‘plami. – 2015.',
    'O‘zbekiston Respublikasi Prezidentining 2017-yil 12-sentabrdagi PQ-3270-son «Respublika bank tizimini yanada rivojlantirish va barqarorligini oshirish chora-tadbirlari to‘g‘risida»gi qarori // Qonunchilik ma’lumotlari milliy bazasi.',
    'Tijorat banklarida korporativ boshqaruv to‘g‘risidagi nizom (ro‘yxat raqami 3254, 30.06.2020). – https://lex.uz.',
    'Tijorat banklarida ichki audit to‘g‘risidagi nizom (ro‘yxat raqami 3302, 07.05.2021). – https://lex.uz.',
    'Banklarda risklarni boshqarish tizimi to‘g‘risidagi nizom (ro‘yxat raqami 3427, 18.04.2023). – https://lex.uz.',
    'O‘zbekiston Respublikasi Markaziy banki Boshqaruvining 2023-yil 18-fevraldagi 4/13-son qarori bilan tasdiqlangan tizimli ahamiyatga molik banklarni aniqlash tartibi.',
    'O‘zbekiston Respublikasi Markaziy banki Boshqaruvining 2023-yil 23-sentabrdagi 34/3-son qarori bilan tasdiqlangan Markaziy bankning asoslangan mulohazasini shakllantirish tartibi.',
    'O‘zbekiston Respublikasi Markaziy banki Boshqaruvining 2023-yil 23-dekabrdagi 42/29-son qarori bilan tasdiqlangan risklarga asoslangan nazorat bo‘yicha qo‘llanma.',
]
NEW_II = [
    'European Banking Authority. EBA updates its risk dashboard for EU banking sector: press release. – London, 11 December 2014. – https://www.eba.europa.eu/eba-updates-its-risk-dashboard-for-eu-banking-sector.',
    'European Banking Authority. Risk Dashboard: Data as of Q4 2024. – Paris: EBA, March 2025.',
    'Sullivan & Cromwell LLP. FFIEC Proposes Revisions to CAMELS Rating System. – New York, 20 May 2026. – https://www.sullcrom.com/insights/memo/2026/May/FFIEC-Proposes-Revisions-CAMELS-Rating-System.',
]
NEW_IV = [
    'Cameron A.C., Gelbach J.B., Miller D.L. Bootstrap-Based Improvements for Inference with Clustered Errors // The Review of Economics and Statistics. – 2008. – Vol. 90, No. 3. – P. 414–427.',
    'Nickell S. Biases in Dynamic Models with Fixed Effects // Econometrica. – 1981. – Vol. 49, No. 6. – P. 1417–1426.',
    'Wilson E.B. Probable Inference, the Law of Succession, and Statistical Inference // Journal of the American Statistical Association. – 1927. – Vol. 22, No. 158. – P. 209–212.',
]
anchor = D.para('11. O‘zbekiston Respublikasi Markaziy banki. Moliya sektorini baholash dasturi (FSAP)')
for s in reversed(NEW_I):
    D.new_para_after(anchor, '0. ' + s, tag='47')
anchor = D.para('II. Xalqaro standartlar, xalqaro tashkilotlar')
# II bo'lim oxiriga: III bo'lim sarlavhasidan oldingi paragrafdan keyin
hdr3 = [p for p in D.paras() if p.text.startswith('III. Monografiyalar')][0]
prev = hdr3._p.getprevious()
from docx.text.paragraph import Paragraph
pp = Paragraph(prev, hdr3._parent)
for s in reversed(NEW_II):
    D.new_para_after(pp, '0. ' + s, tag='47')
hdr5 = [p for p in D.paras() if p.text.startswith('V. Dissertatsiyalar')]
if hdr5:
    pp = Paragraph(hdr5[0]._p.getprevious(), hdr5[0]._parent)
else:
    pp = D.para('IV. Ilmiy maqolalar')
for s in reversed(NEW_IV):
    D.new_para_after(pp, '0. ' + s, tag='47')
# FSSA havolasi adabiyotlar ro'yxatida
D.insert([p for p in D.paras() if '25/145' in (p.text or '') and p.text[:2].isdigit()][0],
         'Washington, DC: IMF, 2025.', FSSA_URL[:-1], tag='FSSA')
# qayta raqamlash (raqamlar o'zgaradi, matn qora qoladi)
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
            runs = D._runs(p)
            r0 = runs[0].find(K.qn('w:t'))
            if r0.text.startswith(m.group(1) + '.'):
                r0.text = str(n) + r0.text[len(m.group(1)):]
            elif r0.text == m.group(1):
                r0.text = str(n)
print('adabiyotlar soni:', n)
D.replace('Dissertatsiyaning tuzilishi va hajmi.', 'adabiyotlar ro‘yxati 174 ta manbani', f'adabiyotlar ro‘yxati {n} ta manbani', tag='47')

# 48-band: formulalarni raqamlash
for s, num in (('Ktaf = (XMHXS – Xmilliy) / Xmilliy × 100%', '(2.1)'), ('Kykr = (S3 / LMHXS) / (NPL / Lmilliy)', '(2.2)'),
               ('si = k + (xi – ak) / (ak+1 – ak)', '(3.1)'), ('MBII = Σ wj · Bj, Bj = (1 / nj)', '(3.2)'),
               ('MBIIit = α + β1·ln(Aktivlar)it', '(3.3)')):
    D.insert(s, '', '  ' + num, tag='48')

D.save(DST)
json.dump(K.LOG, open('/home/user/tarjima/manbalar/hisob_kitob/kam_tuzat_log.json', 'w'), ensure_ascii=False, indent=0)
print('saqlandi:', DST, len(K.LOG), 'ta amal')
