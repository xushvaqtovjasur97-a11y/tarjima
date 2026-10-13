"""3.1.2-rasm: chegara nuqtalarida ball qiymatlari, ikki misol bank nuqtasi; formula izohi (namuna)."""
import re, sys, copy
sys.path.insert(0, '/home/user/tarjima/tools')
import docx
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from lxml import etree
import wordart as W

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
DST = '/home/user/tarjima/Namuna_3.1.2-rasm.docx'
X1, Y1 = [3, 4.5, 5.75, 7, 8, 9.25, 10.5, 13], [1, 1, 1.5, 2, 3, 3.5, 4, 4]
LB1 = {1: ('1,0', 't'), 3: ('2,0', 'l'), 4: ('3,0', 'l'), 6: ('4,0', 't')}
HIDE = [2, 5]   # 1,5 va 3,5 ball nuqtalari – markersiz
X3, Y3 = [8.1, 10.4], [3.04, 3.96]
LB3 = {0: ('8,1% → 3,04 ball', 'r', 0.0, 0.07), 1: ('10,4% → 3,96 ball', 'l')}
XT, YT = [4.5, 7, 8, 10.5], [1, 2, 3, 4]
LBT = {0: ('1,0', 't'), 1: ('2,0', 'l'), 2: ('3,0', 'l'), 3: ('4,0', 't')}
RED = 'C00000'
OLD_K = '(k = 1, 2, 3)'
NEW_K = '(k = 1, 2, 3, 4)'
OLD_M = 'bir xil – 3 ball oladi'
NEW_M = 'bir xil – 3 ball oladi, uzluksiz shkalada esa mos ravishda 3,96 va 3,04 ball oladi'


def _num(tag, ref, vals):
    pts = ''.join(f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i, v in enumerate(vals))
    return (f'<c:{tag}><c:numRef><c:f>{ref}</c:f><c:numCache><c:formatCode>General</c:formatCode>'
            f'<c:ptCount val="{len(vals)}"/>{pts}</c:numCache></c:numRef></c:{tag}>')


def _dlbl(i, text, pos, col, dx=0.0, dy=0.0):
    lay = (f'<c:layout><c:manualLayout><c:x val="{dx}"/><c:y val="{dy}"/></c:manualLayout></c:layout>'
           if dx or dy else '')
    return (f'<c:dLbl><c:idx val="{i}"/>{lay}<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr>{W._rpr(1000, True, col)}</a:pPr>'
            f'<a:r><a:rPr lang="uz-Latn-UZ" sz="1000" b="1"><a:solidFill><a:srgbClr val="{col}"/></a:solidFill>'
            f'<a:latin typeface="Times New Roman"/><a:cs typeface="Times New Roman"/></a:rPr><a:t>{text}</a:t></a:r></a:p></c:rich></c:tx>'
            f'<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr><c:dLblPos val="{pos}"/><c:showLegendKey val="0"/><c:showVal val="1"/>'
            f'<c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbl>')


def _dlbls(items, col):
    return ('<c:dLbls>' + ''.join(_dlbl(i, *v[:2], col, *v[2:]) for i, v in items.items()) +
            '<c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/>'
            '<c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>')


def chart(xml):
    sers = re.findall(r'<c:ser>.*?</c:ser>', xml, re.S)
    s1 = sers[0]
    n1 = re.sub(r'<c:xVal>.*?</c:xVal>', _num('xVal', 'Sheet1!$A$2:$A$9', X1), s1, flags=re.S)
    n1 = re.sub(r'<c:yVal>.*?</c:yVal>', _num('yVal', 'Sheet1!$B$2:$B$9', Y1), n1, flags=re.S)
    n1 = re.sub(r'<c:marker>.*?</c:marker>', '<c:marker><c:symbol val="none"/></c:marker>', n1, count=1, flags=re.S)
    s4 = ('<c:ser><c:idx val="3"/><c:order val="3"/><c:tx><c:strRef><c:f>Sheet1!$H$1</c:f><c:strCache><c:ptCount val="1"/>'
          '<c:pt idx="0"><c:v>Chegaralar</c:v></c:pt></c:strCache></c:strRef></c:tx>'
          '<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr>'
          '<c:marker><c:symbol val="circle"/><c:size val="6"/><c:spPr><a:solidFill><a:srgbClr val="2A78D6"/></a:solidFill>'
          '<a:ln><a:noFill/></a:ln></c:spPr></c:marker>' + _dlbls(LBT, '1F4E79') +
          _num('xVal', 'Sheet1!$G$2:$G$5', XT) + _num('yVal', 'Sheet1!$H$2:$H$5', YT) + '<c:smooth val="0"/></c:ser>')
    s3 = ('<c:ser><c:idx val="2"/><c:order val="2"/><c:tx><c:strRef><c:f>Sheet1!$F$1</c:f><c:strCache><c:ptCount val="1"/>'
          '<c:pt idx="0"><c:v>Misol: ikki bank</c:v></c:pt></c:strCache></c:strRef></c:tx>'
          '<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr>'
          f'<c:marker><c:symbol val="diamond"/><c:size val="13"/><c:spPr><a:solidFill><a:srgbClr val="{RED}"/></a:solidFill>'
          '<a:ln><a:noFill/></a:ln></c:spPr></c:marker>' + _dlbls(LB3, RED) +
          _num('xVal', 'Sheet1!$E$2:$E$3', X3) + _num('yVal', 'Sheet1!$F$2:$F$3', Y3) + '<c:smooth val="0"/></c:ser>')
    xml = xml.replace(s1, n1, 1).replace(sers[1], sers[1] + s3 + s4, 1)
    xml = re.sub(r'(<c:legend><c:legendPos val="[a-z]+"/>)', r'\1<c:legendEntry><c:idx val="3"/><c:delete val="1"/></c:legendEntry>', xml, count=1)
    rows = [['x (uzluksiz)', 'Uzluksiz shkala (taklif)', 'x (pog‘onali)', 'Pog‘onali shkala', 'x (misol)', 'Misol: ikki bank', 'x (chegara)', 'Chegaralar']]
    X2 = [3, 7, 7, 8, 8, 10.5, 10.5, 13]; Y2 = [1, 1, 2, 2, 3, 3, 4, 4]
    for i in range(8):
        rows.append([X1[i], Y1[i], X2[i], Y2[i],
                     X3[i] if i < 2 else None, Y3[i] if i < 2 else None,
                     XT[i] if i < 4 else None, YT[i] if i < 4 else None])
    return xml, rows


def asl_xml():
    d = docx.Document(F)
    cap = [p for p in d.paragraphs if p.text.startswith('3.1.2-rasm. Birinchi')][0]
    rid = [e.get(qn('r:id')) for e in cap._p.getprevious().iter('{http://schemas.openxmlformats.org/drawingml/2006/chart}chart')][0]
    return d.part.related_parts[rid].blob.decode('utf-8')


if __name__ == '__main__':
    xml, rows = chart(asl_xml())
    dd = docx.Document()
    sec = dd.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin, sec.right_margin, sec.top_margin, sec.bottom_margin = Cm(3), Cm(1.5), Cm(2), Cm(2)
    st = dd.styles['Normal']; st.font.name = 'Times New Roman'; st.font.size = Pt(14)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    st.paragraph_format.space_after = Pt(0)

    def par(parts, align=WD_ALIGN_PARAGRAPH.JUSTIFY, size=14, bold=False):
        q = dd.add_paragraph(); q.alignment = align
        for t, col, strike in parts:
            r = q.add_run(t); r.font.size = Pt(size); r.bold = bold; r.font.strike = strike
            if col: r.font.color.rgb = RGBColor.from_string(col)
        return q
    par([('3.1.2-rasm (namuna)', None, False)], WD_ALIGN_PARAGRAPH.CENTER, bold=True)
    q = dd.add_paragraph(); q.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rid = W.add_chart(dd, (xml, rows))
    q._p.append(etree.fromstring(W.chart_run_xml(rid, 16.0, 8.5, '3.1.2-rasm')))
    par([('3.1.2-rasm. Birinchi darajali asosiy kapital monandligini ballga aylantirish funksiyasi', None, False)],
        WD_ALIGN_PARAGRAPH.CENTER, bold=True)
    dd.add_paragraph()
    par([('Matndagi tuzatishlar:', None, False)], size=12, bold=True)
    par([('1) (3.1) formula ikki qatorga bo‘linadi, raqami qatorda qoladi:', None, False)], size=12)
    par([('si = k + (xi – ak) / (ak+1 – ak), agar ak < xi ≤ ak+1;', 'FF0000', False)], WD_ALIGN_PARAGRAPH.CENTER, 13)
    par([('si = 1, agar xi ≤ a1; si = 4, agar xi ≥ a4,          (3.1)', 'FF0000', False)], WD_ALIGN_PARAGRAPH.CENTER, 13)
    par([('2) Formuladan keyingi izohda: ', None, False), (OLD_K, '7F7F7F', True), (' → ', None, False), (NEW_K, 'FF0000', False)], size=12)
    par([('3) Formuladan oldingi xatboshida: …', None, False), (OLD_M, '7F7F7F', True), (' → ', None, False),
         (NEW_M, 'FF0000', False), ('…', None, False)], size=12)
    dd.save(DST)
    from pardoz import pardoz
    print(pardoz(DST), DST)
