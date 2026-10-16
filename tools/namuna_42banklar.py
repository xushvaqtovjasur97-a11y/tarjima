"""4.2-§: tijorat banklari darajasidagi prognoz – namuna.
1) “O‘zsanoatqurilishbank”, “Mikrokreditbank” va “Aloqabank” ATB: muammoli kreditlar ulushining haqiqiy (choraklik) va ssenariy prognozi;
2) bank darajasidagi model (2-model, bank paneli) bilan tekshiruv: proporsional taqsimlash va bevosita bank prognozi taqqoslanadi.
Hisob-kitob: manbalar/hisob_kitob/bank42.py."""
import os, re, sys
sys.path.insert(0, '/home/user/tarjima/tools')
HK = '/home/user/tarjima/manbalar/hisob_kitob'
_cwd = os.getcwd(); os.chdir(HK); sys.path.insert(0, HK)
import bank42 as B
os.chdir(_cwd)
import numpy as np
import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from lxml import etree
import wordart as W

DST = '/home/user/tarjima/Namuna_4.2_banklar_prognozi.docx'
NM = {'O‘zsanoatqurilishbank': '“O‘zsanoatqurilishbank” ATB', 'Mikrokreditbank': '“Mikrokreditbank” ATB', 'Aloqa bank': '“Aloqabank” ATB'}
COL = {'Asosiy': '2A78D6', 'Ijobiy': '70AD47', 'Salbiy': 'FFC000', 'Keskin salbiy': 'C00000'}
ORD = ['Asosiy', 'Ijobiy', 'Salbiy', 'Keskin salbiy']
# yorliqlar: 2027 (12-nuqta – indeks 8) va 2028 (indeks 12) – 4.2.3-rasm bilan bir xil joylashuv
POS = {'Asosiy': {8: 't', 12: 'r'}, 'Ijobiy': {8: 'b', 12: 'r'}, 'Salbiy': {8: 't', 12: 'r'}, 'Keskin salbiy': {8: 't', 12: 't'}}


def c(x, k=2):
    return f'{x:.{k}f}'.replace('.', ',').replace('-', '−')


def qx(davr):
    y, q = davr.split('_'); return int(y) - 1 + {'I': .25, 'II': .5, 'III': .75, 'IV': 1.0}[q]


def _lbl(i, pos, color):
    return (f'<c:dLbl><c:idx val="{i}"/><c:numFmt formatCode="0.00" sourceLinked="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>'
            f'<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>{W._rpr(1000, True, color)}</a:pPr><a:endParaRPr lang="uz-Latn-UZ"/></a:p></c:txPr>'
            f'<c:dLblPos val="{pos}"/><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/>'
            '<c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbl>')


def labels(xml, spec):
    for s in re.findall(r'<c:ser>.*?</c:ser>', xml, re.S):
        name = re.search(r'<c:tx>.*?<c:v>([^<]*)</c:v>', s, re.S).group(1)
        if name not in spec:
            continue
        n = int(re.search(r'<c:yVal>.*?<c:ptCount val="(\d+)"/>', s, re.S).group(1))
        col = re.search(r'<c:spPr><a:ln[^>]*><a:solidFill><a:srgbClr val="(\w+)"/>', s).group(1)
        sp = spec[name]
        body = ''.join(_lbl(i, sp[i], col) if i in sp else f'<c:dLbl><c:idx val="{i}"/><c:delete val="1"/></c:dLbl>' for i in range(n))
        d = (f'<c:dLbls>{body}<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr><c:showLegendKey val="0"/><c:showVal val="0"/>'
             '<c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>')
        xml = xml.replace(s, s.replace('<c:xVal>', d + '<c:xVal>', 1), 1)
    return xml


def bank_chart(b, legend):
    h = B.HIST[b]
    xs = [qx(dv) for dv, _ in h]; ys = [round(v, 2) for _, v in h]
    S = [('Haqiqiy (2021–2025)', xs, ys, '000000', 28575, None, 'circle')]
    fx = [2025.0] + [2025 + .25 * (k + 1) for k in range(12)]
    allv = list(ys)
    for s in ORD:
        v = [ys[-1]] + B.PATH[('P', s, b)]; allv += v
        S.append((f'{s} ssenariy', fx, [round(x, 2) for x in v], COL[s], 22225, 'dash', None))
    lo, hi = np.floor(min(allv) - .3), np.ceil(max(allv) + .5)
    xml, rows = W.scatter_chart(S, '', 'Ulush, %', 2021, 2028.6, 1, float(lo), float(hi), 1,
                                xfmt='0', yfmt='0.0', legend=legend)
    xml = xml.replace('<c:size val="7"/>', '<c:size val="5"/>')
    spec = {f'{s} ssenariy': POS[s] for s in ORD}
    spec['Haqiqiy (2021–2025)'] = {len(ys) - 1: 'b'}
    return labels(xml, spec), rows


if __name__ == '__main__':
    dd = docx.Document()
    sec = dd.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin, sec.right_margin, sec.top_margin, sec.bottom_margin = Cm(3), Cm(1.5), Cm(2), Cm(2)
    st = dd.styles['Normal']; st.font.name = 'Times New Roman'; st.font.size = Pt(14)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    st.paragraph_format.space_after = Pt(0); st.paragraph_format.line_spacing = 1.5

    def par(text, bold=False, center=False, size=None, indent=True, keep=False):
        p = dd.add_paragraph(); r = p.add_run(text); r.bold = bold
        if size: r.font.size = Pt(size); p.paragraph_format.line_spacing = 1.0
        if center: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            if indent: p.paragraph_format.first_line_indent = Cm(1.25)
        p.paragraph_format.keep_with_next = keep
        return p

    def table(rows, widths, size=11, hdr=1):
        tb = dd.add_table(rows=len(rows), cols=len(rows[0])); tb.style = 'Table Grid'; tb.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i, r in enumerate(rows):
            for j, v in enumerate(r):
                cell = tb.cell(i, j); cell.width = Cm(widths[j]); p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if (i < hdr or j > 0) else WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.line_spacing = 1.0
                run = p.add_run(str(v)); run.font.size = Pt(size); run.bold = i < hdr
        return tb

    T, PK = B.T, B.PEAK
    par('4.2-§ uchun tijorat banklari darajasidagi prognoz (namuna)', bold=True, center=True)
    par('1. Uch bank bo‘yicha prognoz rasmi (4.2.3-rasmdan keyin, yangi 4.2.4-rasm)', bold=True, indent=False)
    pk = {b: {s: PK[('P', s, b)] for s in ORD} for b in B.BANKS}
    s25 = {b: B.HIST[b][-1][1] for b in B.BANKS}
    par('Tizim darajasidagi prognoz har bir tijorat bankiga qo‘llanilganda, ssenariylarning ta’siri banklar kesimida turlicha namoyon bo‘ladi. '
        'Ikkinchi bobda batafsil tahlil qilingan uch bank bo‘yicha natijalar 4.2.4-rasmda keltirilgan.')
    for k, b in enumerate(B.BANKS):
        par(f'{"abv"[k]}) {NM[b]}', bold=True, center=True, size=12, keep=True)
        q = dd.add_paragraph(); q.alignment = WD_ALIGN_PARAGRAPH.CENTER; q.paragraph_format.keep_with_next = True
        rid = W.add_chart(dd, bank_chart(b, legend=(k == 2)))
        q._p.append(etree.fromstring(W.chart_run_xml(rid, 16.0, 8.0 if k == 2 else 6.4, f'4.2.4-rasm {k}')))
    par('4.2.4-rasm. Tanlangan tijorat banklarida muammoli kreditlar ulushining haqiqiy dinamikasi (2021–2025-yillar) va ssenariylar bo‘yicha '
        'prognozi (2026–2028-yillar), %', bold=True, center=True)
    par('Izoh: haqiqiy qiymatlar – choraklik, milliy hisobot; prognoz – tizim modeli bo‘yicha og‘ishning bankning 2025-yil IV choragidagi ulushiga '
        'proporsional taqsimlanishi.', size=12, indent=False)
    sq, mk, al = B.BANKS
    par(f'Rasmdan ko‘rinib turibdiki, “O‘zsanoatqurilishbank” va “Aloqabank” ATBlarda muammoli kreditlar ulushi 2025-yil oxirida deyarli bir xil '
        f'({c(s25[sq])} va {c(s25[al])} foiz) bo‘lib, keskin salbiy ssenariyda 2027-yil oxirida {c(pk[sq]["Keskin salbiy"])}–{c(pk[al]["Keskin salbiy"])} '
        f'foizgacha oshadi, ya’ni 5 foizlik erta ogohlantirish chegarasidan past qoladi. “Mikrokreditbank” ATB da esa ulush 2025-yilning o‘zida '
        f'{c(s25[mk])} foizni tashkil etib, salbiy ssenariyda {c(pk[mk]["Salbiy"])} foizga, keskin salbiy ssenariyda {c(pk[mk]["Keskin salbiy"])} '
        'foizga yetadi va chegaradan sezilarli oshadi. Asosiy va ijobiy ssenariylarda esa uch bankning birortasida ham ulush 5 foizdan oshmaydi. '
        'Demak, bir xil makroiqtisodiy shok kredit portfeli sifati past bo‘lgan bankda ancha kuchli ta’sir ko‘rsatadi.')

    par('2. Bank darajasidagi model bilan tekshiruv (4.2.4-jadvaldan keyin)', bold=True, indent=False)
    par('Prognoz natijalari tanlangan usulga bog‘liq emasligini tekshirish uchun muammoli kreditlar ulushi har bir bank uchun bevosita '
        '4.2.2-jadvaldagi 2-model (25 ta bank paneli) asosida ham prognoz qilindi. Bu modelda YaIM o‘sishining ta’siri o‘sha chorakning o‘zida '
        'namoyon bo‘ladi va barcha banklarda bir xil foiz punktga teng bo‘ladi, asosiy usulda esa tizim bo‘yicha og‘ish to‘rt chorakdan keyin '
        'namoyon bo‘lib, banklarga ularning ulushiga proporsional taqsimlanadi. Natijalar 4.2.5-jadvalda keltirilgan.')
    par('4.2.5-jadval. Bank darajasidagi prognozning ikki usul bo‘yicha natijalari: salbiy va keskin salbiy ssenariylar', bold=True, center=True, size=12)
    rows = [['Ssenariy, yil', 'Muammoli kreditlar ulushi (25 bank, o‘rtacha tortilgan), %: asosiy usul / bank modeli',
             'Qizil zonadagi banklar soni: asosiy usul / bank modeli', 'Qizil zonadagi banklar aktivlari ulushi, %: asosiy usul / bank modeli']]
    for s in ('Salbiy', 'Keskin salbiy'):
        for y in (2026, 2027, 2028):
            p_, b_ = T.loc[('P', s, y)], T.loc[('B', s, y)]
            rows.append([f'{s}, {y}', f'{c(p_.S3_w)} / {c(b_.S3_w)}', f'{int(p_.qizil)} / {int(b_.qizil)}', f'{c(p_.qizil_aktiv, 1)} / {c(b_.qizil_aktiv, 1)}'])
    rows.append(['Uch bank bo‘yicha eng yuqori ulush (keskin salbiy), %',
                 '; '.join(f'{x}: {c(PK[("P", "Keskin salbiy", b)])} / {c(PK[("B", "Keskin salbiy", b)])}' for x, b in zip(('SQB', 'MKB', 'Aloqabank'), B.BANKS)), '', ''])
    tb = table(rows, [3.4, 4.6, 3.6, 4.4], size=11)
    last = tb.rows[-1].cells; last[1].merge(last[3])
    par('Izoh: asosiy usul – tizim modeli bo‘yicha og‘ishning banklarga proporsional taqsimlanishi (4.2.4-jadval); bank modeli – 2-model '
        '(muammoli kreditlar ulushi = bank effekti + 0,684 × oldingi chorak ulushi − 0,153 × YaIM o‘sishi); asosiy va ijobiy ssenariylarda ikkala '
        'usulda ham qizil zonada bank yo‘q.', size=12, indent=False)
    r27p, r26b = T.loc[('P', 'Keskin salbiy', 2027)], T.loc[('B', 'Keskin salbiy', 2026)]
    par('Jadvaldan ko‘rinib turibdiki, ikki usul bir xil asosiy natija beradi. Keskin salbiy ssenariyda qizil zonaga ikkala usulda ham aynan '
        f'bir xil to‘rt bank – “Agrobank”, “Asaka bank”, “Mikrokreditbank” va “Turonbank” ATBlar tushadi, ularning aktivlari ulushi esa '
        f'{c(r27p.qizil_aktiv, 1)} foizni tashkil etadi. Salbiy ssenariyda ham ikkala usulda “Mikrokreditbank” va “Turonbank” ATBlar qizil zonaga '
        'tushadi. Usullar o‘rtasidagi farq asosan vaqtda namoyon bo‘ladi. Bank modelida shok o‘sha yilning o‘zida (2026-yil), asosiy usulda esa '
        'bir yildan keyin (2027-yil) eng yuqori ta’sir ko‘rsatadi. Bundan tashqari, bank modelida og‘ish barcha banklarga bir xil foiz punktda '
        f'qo‘shilgani sababli muammoli kreditlar ulushi past bo‘lgan banklarda ta’sir kuchliroq, masalan, “O‘zsanoatqurilishbank” ATB da eng yuqori '
        f'ulush {c(PK[("B", "Keskin salbiy", sq)])} foizni tashkil etadi ({c(PK[("P", "Keskin salbiy", sq)])} foizga nisbatan). Demak, erta ogohlantirish '
        'zonalari bo‘yicha xulosalar prognoz usuliga bog‘liq emas, ya’ni barqaror hisoblanadi.')
    dd.save(DST)
    from pardoz import pardoz
    print(pardoz(DST), DST)
