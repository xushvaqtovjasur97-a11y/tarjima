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

# --- 2021–2025: rasmiy hisobotlar bo‘yicha kodlash (M – boshqaruv va kengash, A – aktivlar sifati,
# IC – ichki nazorat, F – firibgarlik, L – likvidlik va moliyalashtirish, R – foiz stavkasi riski, D – daromadlilik / biznes modeli)
BANKS = [
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

CATN = {'M': 'Boshqaruv va kengash\nnazoratining zaifligi', 'A': 'Aktivlar sifati va kredit\nkonsentratsiyasi', 'I': 'Ichki nazorat va\nauditning zaifligi',
        'F': 'Firibgarlik', 'L': 'Likvidlik va\nmoliyalashtirish', 'R': 'Foiz stavkasi\nriski', 'D': 'Daromadlilik va\nbiznes modeli'}
ORD = sorted(KEYS, key=lambda k: (-cnt[k], KEYS.index(k)))      # kamayish tartibida
CATS = [CATN[k] for k in ORD]
SER = [('2021–2025-yillar', [pct[k] for k in ORD], '2A78D6')]

xml, rows = W.bar_chart(CATS, SER, ymax=70, ymin=0, major=10, label_fmt='0', gap=60, legend=False)
xml = xml.replace('<c:barDir val="col"/>', '<c:barDir val="bar"/>', 1)
xml = xml.replace('<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/>', '<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="maxMin"/>', 1)
xml = xml.replace('<c:axPos val="b"/>', '<c:axPos val="l"/>', 1)
xml = xml.replace('<c:crossAx val="111"/><c:crosses val="autoZero"/>', '<c:crossAx val="111"/><c:crosses val="max"/>', 1).replace('<c:axPos val="l"/><c:majorGridlines>', '<c:axPos val="b"/><c:majorGridlines>', 1)


SRC_TAB = ('Manba: Tadqiqot natijasida muallif tomonidan quyidagi rasmiy hisobotlar asosida tuzildi: FDIC Office of Inspector General. Material Loss Reviews EVAL-24-02 (Signature Bank), EVAL-24-03 (First Republic Bank), EVAL-25-01 (Republic First Bank); Failed Bank Reviews AEC Memorandum 24-01 (Citizens Bank), REV-25-02 va EVAL-26-01 (Pulaski Savings Bank) – https://www.fdicoig.gov/reports-publications/bank-failures; Board of Governors of the Federal Reserve System, Office of Inspector General. Material Loss Reviews 2023-SR-B-013 (Silicon Valley Bank), 2024-SR-B-004 (Heartland Tri-State Bank) – https://oig.federalreserve.gov; U.S. Department of the Treasury, Office of Inspector General. Failed Bank Limited Reviews OIG-25-024 (First National Bank of Lindsay), OIG-26-023 (Santa Anna National Bank) – https://oig.treasury.gov; BaFin: Greensill Bank AG bo‘yicha moratoriy (03.03.2021), North Channel Bank bo‘yicha moratoriy (2023-yil yanvar); Bank of England: Silicon Valley Bank UK Limited bo‘yicha qaror (13.03.2023).')
IZOH_TAB = ('Izoh: B – boshqaruv va kengash nazoratining zaifligi; A – aktivlar sifati va kredit konsentratsiyasi; N – ichki nazorat va auditning zaifligi; F – firibgarlik; L – likvidlik va moliyalashtirish; S – foiz stavkasi riski; D – daromadlilik va biznes modeli. Bir bankda bir nechta sabab qayd etilishi mumkin; sabablar rasmiy hisobotlarda qayd etilgan omillar asosida muallif tomonidan guruhlangan.')
LAB = {'M': 'B', 'A': 'A', 'I': 'N', 'F': 'F', 'L': 'L', 'R': 'S', 'D': 'D'}


def texts(fn, ilova='9-ilova'):
    """fn(matn) – izoh belgisi (namunada [n], dissertatsiyada [[FN:...]])."""
    intro = ('Buni rivojlangan mamlakatlar bank tizimlari tajribasi ham tasdiqlaydi. 1.1.5-rasmda “Katta yettilik” mamlakatlarida 2021–2025-yillarda faoliyati to‘xtatilgan banklar misolida moliyaviy barqarorlik yo‘qotilishining sabablari keltirilgan.')
    caption = '“Katta yettilik” mamlakatlarida 2021–2025-yillarda faoliyati to‘xtatilgan banklarda moliyaviy barqarorlik yo‘qotilishining sabablari, tegishli sabab qayd etilgan banklar ulushi, %'
    cap_fn = ('Tadqiqot natijasida muallif tomonidan ' + ilova + 'da keltirilgan rasmiy hisobotlar asosida ishlab chiqildi (' + str(n) + ' ta bank). Bir bankda bir nechta sabab qayd etilishi mumkin.')
    a1 = ('Rasm ma’lumotlari shuni ko‘rsatadiki, boshqaruv va kengash nazoratining zaifligi banklar barqarorligi yo‘qotilishining eng ko‘p qayd etilgan sababi bo‘lib, u ' + str(n) + ' ta bankning ' + str(pct['M']) + ' foizida kuzatilgan (' + ilova + '). Undan keyingi o‘rinlarda likvidlik va moliyalashtirish muammolari, ichki nazorat tizimining zaifligi va firibgarlik turadi (har biri ' + str(pct['L']) + ' foiz). 2023-yilda Silicon Valley Bank va First Republic Bank misolida sug‘urtalanmagan depozitlarning yuqori konsentratsiyasi va uzoq muddatli aktivlar bo‘yicha foiz stavkasi riski bir necha kun ichida likvidlik inqiroziga aylanishi mumkinligi namoyon bo‘ldi'
          + fn('Board of Governors of the Federal Reserve System, Office of Inspector General. Material Loss Review of Silicon Valley Bank: Report 2023-SR-B-013. – Washington, DC, September 25, 2023. – Executive Summary; FDIC Office of Inspector General. Material Loss Review of First Republic Bank: Report EVAL-24-03. – Arlington, VA, November 2023. – Executive Summary.')
          + ', foiz stavkasi riski esa ' + str(pct['R']) + ' foiz bankda asosiy sabablar qatorida bo‘lgan. Aktivlar sifati va kredit konsentratsiyasi faqat ' + str(pct['A']) + ' foiz holatda asosiy sabab bo‘lgan, chunki so‘nggi yillardagi yirik banklar muvaffaqiyatsizliklari kredit zararlari bilan emas, balki likvidlik va foiz stavkasi riski bilan bog‘liq bo‘ldi. Kichik banklarda esa firibgarlik va ichki nazorat tizimining buzilishi yetakchi sabablarga aylangan.')
    a2 = ('Ommaviy depozitlar oqimi bank faoliyatining to‘xtatilishiga bevosita sabab bo‘lgan hollarda ham uning ortida odatda fundamental zaifliklar turadi. S. Korreya, S. Lak va E. Verner AQSh bank tizimining 1863–2024-yillardagi ma’lumotlarini o‘rganib, AQSh Valyuta nazorati boshqarmasi 1865–1937-yillarda faoliyati to‘xtatilgan banklarning 2 foizdan kamida ommaviy mablag‘ talab qilishni asosiy sabab deb topganini, banklar inqirozlarining asosiy sababi esa deyarli har doim bank fundamental ko‘rsatkichlarining yomonlashuvi ekanini ko‘rsatgan'
          + fn('Correia S., Luck S., Verner E. Failing Banks // Quarterly Journal of Economics. – 2026. – Vol. 141, No. 1. – P. 147–204; ishchi varianti: arXiv:2506.06082. – 2025. – P. 1, 52.'))
    a2_end = 'aktivlar sifati va boshqaruv sifati bilan bir qatorda likvidlik hamda bozor risklariga, O‘zbekiston sharoitida birinchi navbatda valyuta riskiga sezgirlik bloklari ham bo‘lishi lozim'
    return dict(intro=intro, caption=caption, cap_fn=cap_fn, a1=a1, a2=a2, a2_end=a2_end)


TAB_TITLE = '2021–2025-yillarda “Katta yettilik” mamlakatlarida faoliyati to‘xtatilgan banklar va ularning sabablari'
HDR = ['№', 'Bank, mamlakat', 'Yil', 'Rasmiy hisobotlarda qayd etilgan sabablar'] + [LAB[k] for k in KEYS] + ['Manba']
ROWS = [[str(i), f'{b[0]} ({b[1]})', str(b[2]), b[3]] + ['+' if k in b[4] else '' for k in KEYS] + [b[5]] for i, b in enumerate(sorted(BANKS, key=lambda b: (b[2], b[1] != 'AQSh')), 1)]
ROWS.append(['', 'Jami, %', '', f'{n} ta bank bo‘yicha ulush'] + [str(pct[k]) for k in KEYS] + [''])

if __name__ == '__main__':
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

    FN.append(texts(lambda x: '')['cap_fn'])   # rasm izohi birinchi bo‘lib chiqadi
    T = texts(fn)
    par('1.1.5-rasmning yangi varianti (namuna)', bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False)
    par(T['intro'])
    p = d.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.first_line_indent = Cm(0)
    rid = W.add_chart(d, (xml, rows))
    p._p.append(etree.fromstring(W.chart_run_xml(rid, 16.0, 12.0, '1.1.5-rasm')))
    par('1.1.5-rasm. ' + T['caption'] + '[1]', bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False)
    par(T['a1'])
    par(T['a2'] + '. Bundan kelib chiqadiki, moliyaviy barqarorlikni baholashda ' + T['a2_end'] + '.')
    d.add_page_break()
    par('9-ilova', bold=True, italic=True, align=WD_ALIGN_PARAGRAPH.RIGHT, indent=False, size=12)
    par(TAB_TITLE, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, size=12)
    t = d.add_table(rows=0, cols=len(HDR)); t.style = 'Table Grid'
    for vals in [HDR] + ROWS:
        cells = t.add_row().cells
        for c, v in zip(cells, vals):
            c.text = v
            for q in c.paragraphs:
                q.paragraph_format.first_line_indent = Cm(0); q.paragraph_format.line_spacing = 1.0
                for run in q.runs:
                    run.font.size = Pt(9)
    par(IZOH_TAB, italic=True, size=10, indent=False)
    par(SRC_TAB, italic=True, size=10, indent=False)
    d.add_page_break()
    par('Sahifa osti snoskalari', bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False)
    for i, f in enumerate(FN, 1):
        par(f'[{i}] {f}', size=11, indent=False)
    d.save(DST)
    print(pardoz(DST), DST)
