"""2.3.4-rasm: SQB da MHXS 9 bo‘yicha 3-bosqich kreditlarining milliy tasnif toifalari bo‘yicha tarkibi, 2021–2025 (namuna).
Manba: SQB MHXS hisobotlari – 2023 y. (2021–2023, S. 53–55), 2025 y. (2024–2025, P. 37–38). Mlrd so‘m."""
import re, sys
sys.path.insert(0, '/home/user/tarjima/tools')
import docx
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from lxml import etree
import wordart as W
from pardoz import pardoz

DST = '/home/user/tarjima/Namuna_2.3.4-rasm.docx'
Y = ['2021', '2022', '2023', '2024', '2025']
CAT = [('Standart', [192.4, 238.3, 848.8, 1013.3, 746.4], '00B050', '000000'),
       ('Substandart', [797.7, 860.1, 518.6, 577.8, 2290.9], 'FFE14D', '000000'),
       ('Qoniqarsiz', [953.2, 490.5, 581.2, 1326.1, 243.9], 'FFA726', '000000'),
       ('Shubhali', [225.3, 988.9, 910.8, 628.2, 642.1], 'FF5050', '000000'),
       ('Umidsiz', [609.4, 1206.3, 45.0, 540.2, 1386.9], 'C00000', 'FFFFFF')]
TOT = [round(sum(c[1][i] for c in CAT), 1) for i in range(5)]
SHARE = [round(100 * (CAT[0][1][i] + CAT[1][1][i]) / TOT[i], 1) for i in range(5)]
LINE = ('Milliy tasnifda muammoli hisoblanmagan qismi (standart va substandart), % (o‘ng o‘q)', SHARE, '1F4E79')
CATS = list(Y)
TOTS = [f'{t:,.1f}'.replace(',', ' ').replace('.', ',') for t in TOT]


PCT = [(c[0], [round(100 * c[1][i] / TOT[i], 1) for i in range(5)], c[2], c[3]) for c in CAT]


def chart():
    xml, rows = W.bar_chart(CATS, PCT, ytitle='%', stacked=True, ymin=0, ymax=114, major=20, label_fmt='0.0', gap=60)
    sers = re.findall(r'<c:ser>.*?</c:ser>', xml, re.S)
    for s_, c in zip(sers, PCT):
        dl = ''.join(f'<c:dLbl><c:idx val="{i}"/><c:layout><c:manualLayout><c:x val="0.075"/><c:y val="-0.012"/></c:manualLayout></c:layout>'
                     f'<c:numFmt formatCode="0.0" sourceLinked="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>'
                     f'<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>{W._rpr(1050, True, "000000")}</a:pPr><a:endParaRPr lang="uz-Latn-UZ"/></a:p></c:txPr>'
                     f'<c:dLblPos val="ctr"/><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/>'
                     f'<c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbl>' for i, v in enumerate(c[1]) if v < 4)
        if dl:
            xml = xml.replace(s_, s_.replace('<c:dLbls>', '<c:dLbls>' + dl, 1), 1)
    pass
    n = len(CATS)
    lbl = ''.join(f'<c:dLbl><c:idx val="{i}"/><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr>{W._rpr(1050, True, "1F3864")}</a:pPr>'
                  f'<a:r><a:rPr lang="uz-Latn-UZ" sz="1050" b="1"><a:solidFill><a:srgbClr val="1F3864"/></a:solidFill><a:latin typeface="Times New Roman"/></a:rPr>'
                  f'<a:t>{t}</a:t></a:r></a:p></c:rich></c:tx><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr><c:dLblPos val="ctr"/>'
                  f'<c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/>'
                  f'<c:showBubbleSize val="0"/></c:dLbl>' for i, t in enumerate(TOTS))
    line = (f'<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/><c:ser><c:idx val="5"/><c:order val="5"/>'
            f'<c:tx>{W._strref("Sheet1!$G$1", ["Jami, mlrd so‘m"])}</c:tx><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>'
            f'<c:marker><c:symbol val="none"/></c:marker><c:dLbls>{lbl}<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>'
            f'<c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>'
            f'<c:cat>{W._strref(f"Sheet1!$A$2:$A${n+1}", CATS)}</c:cat><c:val>{W._numref(f"Sheet1!$G$2:$G${n+1}", [109] * n)}</c:val><c:smooth val="0"/></c:ser>'
            f'<c:marker val="1"/><c:axId val="111"/><c:axId val="222"/></c:lineChart>')
    xml = xml.replace('<c:catAx>', line + '<c:catAx>', 1)
    # legendadan “Jami” ni olib tashlash
    xml = xml.replace('<c:legend><c:legendPos val="b"/>', '<c:legend><c:legendPos val="b"/><c:legendEntry><c:idx val="5"/><c:delete val="1"/></c:legendEntry>', 1)
    rows = [r + [x] for r, x in zip(rows, ['Jami, mlrd so‘m'] + TOT)]
    return xml, rows


OLD = '2024-yilda bu ulush 38,9 foizni tashkil etgan edi, ya’ni bir yil ichida ikki tasnif o‘rtasidagi nomuvofiqlik sezilarli kuchaygan.'
NEW = (f'Bu ulush 2021-yilda {SHARE[0]:.1f} foizni, 2022-yilda {SHARE[1]:.1f} foizni, 2023-yilda {SHARE[2]:.1f} foizni, 2024-yilda {SHARE[3]:.1f} foizni '
       'tashkil etgan, ya’ni ikki tasnif o‘rtasidagi nomuvofiqlik besh yil davomida muntazam saqlanib, 2025-yilda eng yuqori darajaga yetgan.').replace('.', ',', 0)
SRC23 = ('Акционерный коммерческий банк “Узбекский промышленно-строительный банк” и его дочерние организации. Консолидированная финансовая отчетность '
         'и аудиторское заключение за годы, закончившиеся 31 декабря 2023, 2022 и 2021 гг. – Ташкент, 2024. – С. 53–55')

if __name__ == '__main__':
    NEW = NEW.replace('35.6', '35,6')
    for a in SHARE:
        NEW = NEW.replace(f'{a:.1f}', f'{a:.1f}'.replace('.', ','))
    print(TOT, SHARE)
    dd = docx.Document()
    sec = dd.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin, sec.right_margin, sec.top_margin, sec.bottom_margin = Cm(3), Cm(1.5), Cm(2), Cm(2)
    st = dd.styles['Normal']; st.font.name = 'Times New Roman'; st.font.size = Pt(14)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    st.paragraph_format.line_spacing = 1.5; st.paragraph_format.space_after = Pt(0)

    def par(text, bold=False, italic=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY, size=14, color=None, strike=False, indent=False):
        q = dd.add_paragraph(); q.alignment = align
        if indent:
            q.paragraph_format.first_line_indent = Cm(1.25)
        r = q.add_run(text); r.bold = bold; r.italic = italic; r.font.size = Pt(size); r.font.strike = strike
        if color:
            r.font.color.rgb = RGBColor.from_string(color)
        return q

    par('2.3.4-rasm – 2021–2025-yillar (namuna)', bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    q = dd.add_paragraph(); q.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rid = W.add_chart(dd, chart())
    q._p.append(etree.fromstring(W.chart_run_xml(rid, 15.0, 7.5, '2.3.4-rasm')))
    par('2.3.4-rasm. “O‘zsanoatqurilishbank” ATB da MHXS 9 bo‘yicha 3-bosqich kreditlarining milliy tasnif toifalari bo‘yicha tarkibi, 2021–2025-yillar, %',
        bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    par('Izoh: ustunlar ustida – 3-bosqich kreditlarining jami summasi, mlrd so‘m.', italic=True, size=12, align=WD_ALIGN_PARAGRAPH.LEFT)
    par('Rasm nomida “mlrd so‘m” o‘rniga “2021–2025-yillar, %” (qizil); ustunlar ustida 3-bosqich kreditlarining jami summasi berilgan. Rasm manbasi (snoska) ga qo‘shiladi:', size=12)
    par(SRC23 + '.', size=12, color='FF0000')
    par('Rasmdan keyingi xatboshida:', size=12)
    par('Hozir: ' + OLD, size=12, color='7F7F7F', strike=True)
    par('Yangi: ' + NEW, size=12, color='FF0000')
    dd.save(DST)
    print(pardoz(DST), DST)
