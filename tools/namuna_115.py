"""1.1.5-rasmning yangi varianti (alohida namuna fayl): AQSh va “Katta yettilik” mamlakatlarida
banklar barqarorligi yo‘qotilishi sabablari – uch davr bo‘yicha taqqoslash."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
import docx
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from lxml import etree
import wordart as W
from pardoz import pardoz

DST = '/home/user/tarjima/Namuna_1.1.5-rasm_yangi.docx'

# --- 2020–2025: rasmiy hisobotlar bo‘yicha kodlash (M – boshqaruv va kengash, A – aktivlar sifati,
# IC – ichki nazorat, F – firibgarlik, L – likvidlik va moliyalashtirish, R – foiz stavkasi riski, D – daromadlilik / biznes modeli)
BANKS = [
    ('Ericson State Bank', 'AQSh', 2020, 'Kengash va boshqaruvning o‘ta zaifligi; bank rahbarining aloqador shaxslarga hujjatsiz kreditlari bo‘yicha zararlar; broker depozitlariga tayanish natijasida likvidlik tanqisligi', 'MAL', 'FDIC OIG, FBR-20-002'),
    ('The First State Bank', 'AQSh', 2020, 'Kengash nazoratining o‘ta zaifligi va risk menejmentining sustligi; xodim tomonidan kredit firibgarligi; kredit berish standartlarining zaifligi', 'MAF', 'FDIC OIG, FBR-21-001'),
    ('First City Bank of Florida', 'AQSh', 2020, 'Sifatsiz aktivlarning katta hajmi va kreditlash amaliyotining zaifligi; 12 yil davomida zarar bilan ishlash', 'AD', 'FDIC OIG, FBR-21-002'),
    ('Almena State Bank', 'AQSh', 2020, 'Korporativ boshqaruv va boshqaruv nazoratining yetishmasligi; tez o‘sish strategiyasi, kreditlarni rasmiylashtirishdagi kamchiliklar, qimmat ulgurji resurslarga tayanish', 'MAL', 'FDIC OIG, FBR-21-003'),
    ('Silicon Valley Bank', 'AQSh', 2023, 'Boshqaruv tez o‘sish va konsentratsiya risklarini nazorat qilmagan; foiz stavkasi riski xedjlari olib tashlangan; 40 mlrd dollarlik depozitlar oqimi', 'MLR', 'FRT OIG, 2023-SR-B-013'),
    ('Signature Bank', 'AQSh', 2023, 'Likvidlik va favqulodda moliyalashtirish mexanizmlarining yetarli emasligi; boshqaruvning risk menejmenti amaliyoti zaif', 'ML', 'FDIC OIG, EVAL-24-02'),
    ('First Republic Bank', 'AQSh', 2023, 'Sug‘urtalanmagan depozitlar konsentratsiyasi, aktiv va passivlar muddatlarining nomutanosibligi, foiz stavkasi riski; boshqa banklar inqirozining yuqumli ta’siri', 'LR', 'FDIC OIG, EVAL-24-03'),
    ('Heartland Tri-State Bank', 'AQSh', 2023, 'Bosh direktorning firibgarlik bilan bog‘liq 47,1 mln dollarlik pul o‘tkazmalari; ichki nazoratning buzilishi va rahbarning hukmron mavqei', 'MIF', 'FRT OIG, 2024-SR-B-004'),
    ('Citizens Bank', 'AQSh', 2023, 'Kengash va boshqaruvning sust kreditlash amaliyoti; yirik yuk tashish kompaniyalari kreditlarining yomonlashuvi', 'MA', 'FDIC OIG, AEC 24-01'),
    ('Republic First Bank', 'AQSh', 2024, 'Muddatigacha saqlanadigan qimmatli qog‘ozlar bo‘yicha zararlarni tan olish (foiz stavkasi riski); kengash va boshqaruvning samarasizligi', 'MR', 'FDIC OIG, EVAL-25-01'),
    ('First National Bank of Lindsay', 'AQSh', 2024, 'Ichki nazoratning keskin buzilishi natijasida kredit portfeli va likvid aktivlarga ta’sir qilgan firibgarlik', 'IF', 'G‘aznachilik OIG, OIG-25-024'),
    ('Pulaski Savings Bank', 'AQSh', 2025, 'Kapitalning qadrsizlanishi: buxgalteriya tizimida aks ettirilmagan 20,7 mln dollarlik depozit sertifikatlari; kengash va boshqaruv nazoratining zaifligi', 'MI', 'FDIC OIG, REV-25-02, EVAL-26-01'),
    ('Santa Anna National Bank', 'AQSh', 2025, 'Ichki nazoratning keskin buzilishi natijasida kredit portfeliga ta’sir qilgan firibgarlik; kengash nazoratining zaifligi', 'MIF', 'G‘aznachilik OIG, OIG-26-023'),
    ('Greensill Bank AG', 'Germaniya', 2021, 'Balansdagi GFG Alliance guruhi debitorlik qarzlari mavjudligini tasdiqlay olmaslik (kredit konsentratsiyasi); balansni soxtalashtirish bo‘yicha jinoiy ariza', 'AF', 'BaFin'),
    ('North Channel Bank', 'Germaniya', 2023, 'Surunkali zarar va barqaror biznes modelining yo‘qligi', 'D', 'BaFin'),
    ('SVB UK', 'Buyuk Britaniya', 2023, 'Bosh bank inqirozi sababli depozitlar oqimi va likvidlik tanqisligi; mijozlar bazasining texnologiya sohasida konsentratsiyasi', 'L', 'Angliya banki'),
]
KEYS = 'MAIFLRD'
n = len(BANKS)
cnt = {k: sum(k in b[4] for b in BANKS) for k in KEYS}
pct = {k: int(100 * cnt[k] / n + 0.5) for k in KEYS}
print(n, cnt, pct)

CATS = ['Boshqaruv va kengash\nnazoratining zaifligi', 'Aktivlar sifati va kredit\nkonsentratsiyasi', 'Ichki nazorat va\nauditning zaifligi',
        'Firibgarlik', 'Likvidlik va\nmoliyalashtirish', 'Foiz stavkasi\nriski', 'Daromadlilik va\nbiznes modeli']
S1 = [90, 98, 25, 11, None, None, None]                 # 1979–1987, OCC (Lavrushin, 44-bet)
S2 = [100, 100, 80, None, 50, None, None]               # 2008–2010, FRT OIG, 20 bank
S3 = [pct[k] for k in KEYS]                             # 2020–2025, G-7, 16 bank
SER = [('1979–1987-yillar (AQSh)', S1, '7F7F7F'), ('2008–2010-yillar (AQSh)', S2, '2A78D6'), (f'2020–2025-yillar (“Katta yettilik”, {n} ta bank)', S3, 'EB6834')]

xml, rows = W.bar_chart(CATS, SER, ymax=100, ymin=0, major=20, label_fmt='0', gap=50)
xml = xml.replace('<c:barDir val="col"/>', '<c:barDir val="bar"/>', 1)
xml = xml.replace('<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/>', '<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="maxMin"/>', 1)
xml = xml.replace('<c:axPos val="b"/>', '<c:axPos val="l"/>', 1)
xml = xml.replace('<c:crossAx val="111"/><c:crosses val="autoZero"/>', '<c:crossAx val="111"/><c:crosses val="max"/>', 1).replace('<c:axPos val="l"/><c:majorGridlines>', '<c:axPos val="b"/><c:majorGridlines>', 1)

d = docx.Document()
sec = d.sections[0]
sec.page_width, sec.page_height = Cm(21), Cm(29.7)
sec.left_margin, sec.right_margin, sec.top_margin, sec.bottom_margin = Cm(3), Cm(1.5), Cm(2), Cm(2)
st = d.styles['Normal']
st.font.name = 'Times New Roman'; st.font.size = Pt(14)
st.element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
st.paragraph_format.line_spacing = 1.5
st.paragraph_format.space_after = Pt(0)


def par(text, bold=False, italic=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY, indent=True, size=14, color=None):
    p = d.add_paragraph()
    p.alignment = align
    p.paragraph_format.first_line_indent = Cm(1.25) if indent else Cm(0)
    r = p.add_run(text); r.bold = bold; r.italic = italic; r.font.size = Pt(size)
    if color:
        r.font.color.rgb = RGBColor.from_string(color)
    return p


FN = []


def fn(text):
    FN.append(text)
    return f'[{len(FN)}]'


par('1.1.5-rasmning yangi varianti (namuna)', bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False)
par('Ma’qul bo‘lsa, quyidagi matn dissertatsiyaning 1.1-§ dagi “Buni AQSh bank tizimi tajribasi ham tasdiqlaydi (1.1.5-rasm).” jumlasi, 1.1.5-rasm va undan keyingi izohning birinchi qismi o‘rniga qo‘yiladi; 1.1.x-jadval 1-ilovaga kiritiladi. Kvadrat qavsdagi raqamlar – sahifa osti snoskalari (oxirida keltirilgan).',
    italic=True, size=12, indent=False, color='595959')

par('Buni rivojlangan mamlakatlar bank tizimlari tajribasi ham tasdiqlaydi. 1.1.5-rasmda AQSh va “Katta yettilik” mamlakatlarida banklar moliyaviy barqarorligi yo‘qotilishining sabablari uch davr bo‘yicha taqqoslangan: 1979–1987-yillardagi AQSh Valyuta nazorati boshqarmasi (OCC) tadqiqoti, 2008–2010-yillardagi global moliyaviy inqiroz davri hamda 2020–2025-yillar.')
p = d.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.first_line_indent = Cm(0)
rid = W.add_chart(d, (xml, rows))
p._p.append(etree.fromstring(W.chart_run_xml(rid, 16.0, 12.0, '1.1.5-rasm')))
par('1.1.5-rasm. Rivojlangan mamlakatlarda banklar moliyaviy barqarorligi yo‘qotilishining sabablari, tegishli sabab qayd etilgan banklar ulushi, %' + fn(
    'Tadqiqot natijasida muallif tomonidan quyidagi manbalar asosida ishlab chiqildi: 1979–1987-yillar – Оценка финансовой устойчивости кредитной организации: учебник / под ред. О.И. Лаврушина, И.Д. Мамоновой. – М.: КНОРУС, 2011. – С. 44; 2008–2010-yillar – Board of Governors of the Federal Reserve System, Office of Inspector General. Summary Analysis of Failed Bank Reviews. – Washington, September 2011. – P. 8–10, 47 (Federal rezerv tizimi nazoratidagi 20 ta bank, “strategik rejalashtirish va qarorlar” hamda “aktivlar konsentratsiyasi” ustunlari bo‘yicha); 2020–2025-yillar – 1-ilovaning 1.x-jadvalida keltirilgan rasmiy hisobotlar asosida. Izoh: bir bankda bir nechta sabab qayd etilishi mumkin; bo‘sh qolgan ustunlar tegishli tadqiqotda alohida baholanmaganini bildiradi.'),
    bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False)

par('Rasm ma’lumotlari shuni ko‘rsatadiki, uchala davrda ham boshqaruv va kengash nazoratining zaifligi banklar barqarorligi yo‘qotilishining eng ko‘p qayd etiladigan sababi bo‘lib qolmoqda: 1979–1987-yillarda u o‘rganilgan muammoli banklarning 90 foizida, 2008–2010-yillarda Federal rezerv tizimi nazoratidagi faoliyati to‘xtatilgan 20 ta bankning barchasida, 2020–2025-yillarda esa “Katta yettilik” mamlakatlarida faoliyati to‘xtatilgan ' + str(n) + ' ta bankning ' + str(pct['M']) + ' foizida qayd etilgan. Aktivlar sifati va kredit konsentratsiyasi ham doimiy omil hisoblanadi, biroq so‘nggi davrda uning ulushi ' + str(pct['A']) + ' foizgacha pasaygan. Buning o‘rniga yangi xavflar kuchaygan. 2023-yilda Silicon Valley Bank, Signature Bank va First Republic Bank misolida sug‘urtalanmagan depozitlarning yuqori konsentratsiyasi va foiz stavkasi riski bir necha kun ichida likvidlik inqiroziga aylanishi mumkinligi namoyon bo‘ldi' + fn(
    'Board of Governors of the Federal Reserve System, Office of Inspector General. Material Loss Review of Silicon Valley Bank. Report 2023-SR-B-013. – Washington, September 25, 2023. – Executive Summary; FDIC Office of Inspector General. Material Loss Review of First Republic Bank. Report EVAL-24-03. – Arlington, November 2023. – Executive Summary.') + ': ' + str(n) + ' ta bankning ' + str(pct['L']) + ' foizida likvidlik va moliyalashtirish, ' + str(pct['R']) + ' foizida esa foiz stavkasi riski asosiy sabablar qatorida bo‘lgan. Kichik banklarda esa firibgarlik (' + str(pct['F']) + ' foiz) va ichki nazorat tizimining buzilishi (' + str(pct['I']) + ' foiz) yetakchi sabablardan biriga aylangan.')
par('Ommaviy depozitlar oqimi bankning faoliyatini to‘xtatishiga bevosita sabab bo‘lgan hollarda ham uning ortida odatda fundamental zaifliklar turadi. S. Korreya, S. Lak va E. Verner AQShning 1863–2024-yillardagi bank tizimini o‘rganib, OCC 1865–1937-yillarda faoliyati to‘xtatilgan banklarning 2 foizdan kamida ommaviy mablag‘ talab qilishni asosiy sabab deb topganini, banklar inqirozlarining asosiy sababi esa deyarli har doim va hamma joyda bank fundamental ko‘rsatkichlarining yomonlashuvi ekanini ko‘rsatgan' + fn(
    'Correia S., Luck S., Verner E. Failing Banks // Quarterly Journal of Economics. – 2026. – Vol. 141, No. 1. – P. 147–204; ishchi varianti: arXiv:2506.06082. – 2025. – P. 1, 52.') + '. Bundan kelib chiqadiki, moliyaviy barqarorlikni baholashda aktivlar sifati va boshqaruv sifati bilan bir qatorda likvidlik hamda bozor risklariga, O‘zbekiston sharoitida birinchi navbatda valyuta riskiga sezgirlik bloklari ham bo‘lishi lozim.')
par('(Keyingi jumla – “Shu bilan birga, mualliflar erta ogohlantirish signallari sifatida …” – o‘zgarishsiz qoladi.)', italic=True, size=12, indent=False, color='595959')

# --- ilova jadvali
d.add_page_break()
par('1-ilovaga taklif', bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False)
par('1.x-jadval. 2020–2025-yillarda “Katta yettilik” mamlakatlarida faoliyati to‘xtatilgan banklar va ularning sabablari' + fn(
    'Tadqiqot natijasida muallif tomonidan quyidagi rasmiy hisobotlar asosida tuzildi: FDIC Office of Inspector General. Failed Bank Review Memoranda FBR-20-002, FBR-21-001, FBR-21-002, FBR-21-003 (2020–2021), AEC Memorandum 24-01 (2024), REV-25-02 (2025); In-Depth Review of Pulaski Savings Bank, EVAL-26-01 (2026); Material Loss Reviews EVAL-24-02 (Signature Bank), EVAL-24-03 (First Republic Bank), EVAL-25-01 (Republic First Bank) – https://www.fdicoig.gov/reports-publications/bank-failures; Board of Governors of the Federal Reserve System, Office of Inspector General. Material Loss Reviews 2023-SR-B-013 (Silicon Valley Bank), 2024-SR-B-004 (Heartland Tri-State Bank); U.S. Department of the Treasury, Office of Inspector General. Failed Bank Limited Reviews OIG-25-024 (First National Bank of Lindsay), OIG-26-023 (Santa Anna National Bank); BaFin press releases on Greensill Bank AG (03.03.2021) va North Channel Bank (01.2023); Bank of England. Statement on Silicon Valley Bank UK Limited (13.03.2023). Belgilar: B – boshqaruv va kengash nazorati, A – aktivlar sifati va kredit konsentratsiyasi, N – ichki nazorat, F – firibgarlik, L – likvidlik va moliyalashtirish, S – foiz stavkasi riski, D – daromadlilik va biznes modeli.'),
    bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, size=12)
LAB = {'M': 'B', 'A': 'A', 'I': 'N', 'F': 'F', 'L': 'L', 'R': 'S', 'D': 'D'}
hdr = ['№', 'Bank, mamlakat', 'Yil', 'Rasmiy hisobotda qayd etilgan sabablar'] + [LAB[k] for k in KEYS] + ['Manba']
t = d.add_table(rows=1, cols=len(hdr)); t.style = 'Table Grid'
for c, h in zip(t.rows[0].cells, hdr):
    c.text = h
for i, b in enumerate(BANKS, 1):
    r = t.add_row().cells
    vals = [str(i), f'{b[0]} ({b[1]})', str(b[2]), b[3]] + ['+' if k in b[4] else '' for k in KEYS] + [b[5]]
    for c, v in zip(r, vals):
        c.text = v
r = t.add_row().cells
for c, v in zip(r, ['', 'Jami, %', '', f'{n} ta bank bo‘yicha ulush'] + [str(pct[k]) for k in KEYS] + ['']):
    c.text = v
widths = [0.8, 3.2, 1.2, 5.1] + [0.6] * 7 + [2.2]
tbl = t._tbl; tp = tbl.tblPr
lay = etree.SubElement(tp, qn('w:tblLayout')); lay.set(qn('w:type'), 'fixed')
tw = tp.find(qn('w:tblW'))
if tw is None:
    tw = etree.SubElement(tp, qn('w:tblW'))
tw.set(qn('w:w'), str(int(sum(widths) * 567))); tw.set(qn('w:type'), 'dxa')
for gc, w in zip(tbl.tblGrid.findall(qn('w:gridCol')), widths):
    gc.set(qn('w:w'), str(int(w * 567)))
for row in t.rows:
    for c, w in zip(row.cells, widths):
        c.width = Cm(w)
        for p in c.paragraphs:
            p.paragraph_format.first_line_indent = Cm(0)
            p.paragraph_format.line_spacing = 1.0
            for run in p.runs:
                run.font.size = Pt(10)
    for c in row.cells[:1] + row.cells[2:3] + row.cells[4:11]:
        for p in c.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
for c in t.rows[0].cells + t.rows[-1].cells:
    for p in c.paragraphs:
        for run in p.runs:
            run.bold = True

d.add_page_break()
par('Sahifa osti snoskalari', bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False)
for i, f in enumerate(FN, 1):
    par(f'[{i}] {f}', size=11, indent=False)
d.save(DST)
print(pardoz(DST), DST)
