"""4.2-§ uchun model chizmalari – namuna:
A) muammoli kreditlar ulushi va 4 chorak oldingi YaIM o‘sishi: nuqtalar va regressiya chizig‘i (1-model asosi);
B) 25 ta bank: muammoli kreditlar ulushi va ROA – nuqtalar bulutи va regressiya chizig‘i (3-model asosi);
C) orqaga tekshiruv: haqiqiy qiymatlar va bir yil oldinga prognozlar (4.2.3-jadval ma’lumotlari)."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
import numpy as np, pandas as pd
import statsmodels.api as sm
import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from lxml import etree
import wordart as W

HK = '/home/user/tarjima/manbalar/hisob_kitob/'
DST = '/home/user/tarjima/Namuna_4.2_model_chizmalari.docx'


def f2(x, n=2):
    return f'{x:.{n}f}'.replace('.', ',').replace('-', '−')


# ---------------- A) NPL(t) ~ YaIM(t−4)
def chart_a():
    D = pd.read_csv(HK + 'tizim42.csv', index_col=0)
    q = pd.concat([D.NPL, D.YaIM.shift(4).rename('G')], axis=1).loc['2021_I':]
    m = sm.OLS(q.NPL, sm.add_constant(q.G)).fit()
    a, b, r2 = m.params.const, m.params.G, m.rsquared
    cols = {'2021': 'C00000', '2022': 'EB6834', '2023': 'BF9000', '2024': '70AD47', '2025': '2A78D6'}
    S = []
    for y, c in cols.items():
        z = q.loc[[k for k in q.index if k.startswith(y)]]
        S.append((f'{y}-yil', [round(v, 2) for v in z.G], [round(v, 2) for v in z.NPL], c, 0, None, 'circle'))
    xs = [0.0, 9.0]
    S.append(('Regressiya chizig‘i', xs, [round(a + b * x, 3) for x in xs], '000000', 19050, None, None))
    xml, rows = W.scatter_chart(S, 'YaIM real o‘sish sur’ati (4 chorak oldin), %', 'Muammoli kreditlar ulushi, %',
                                0, 10, 2, 2, 6, 1, xfmt='0', yfmt='0.0', legend=True)
    xml = xml.replace('<c:size val="7"/>', '<c:size val="8"/>').replace('<c:legendPos val="b"/>', '<c:legendPos val="r"/>', 1)
    info = dict(a=a, b=b, r2=r2, p=m.pvalues.G, n=int(m.nobs), r=q.NPL.corr(q.G))
    return (xml, rows), info


# ---------------- B) bank paneli: ROA ~ NPL
def chart_b():
    p = pd.read_csv(HK + 'panel33.csv')
    d = p[p.tanlov].copy()

    def w(x, al=.01):
        lo, hi = x.quantile([al, 1 - al]); return x.clip(lo, hi)
    for c in ('S3', 'ROAq'):
        d[c] = w(d[c])
    d = d.dropna(subset=['S3', 'ROAq'])
    m = sm.OLS(d.ROAq, sm.add_constant(d.S3)).fit()
    a, b = m.params.const, m.params.S3
    g1 = d[d.DAVLAT == 1]; g0 = d[d.DAVLAT != 1]
    grp = sorted([('Davlat ulushi mavjud banklar', g1, 'C00000'), ('Boshqa banklar', g0, '2A78D6')], key=lambda t: -len(t[1]))
    S = [(f'{nm} ({len(g)} ta kuzatuv)', [round(v, 2) for v in g.S3], [round(v, 2) for v in g.ROAq], c, 0, None, 'circle') for nm, g, c in grp]
    xmax = float(np.ceil(d.S3.max() / 2) * 2)
    xs = [0.0, xmax]
    S.append(('Regressiya chizig‘i', xs, [round(a + b * x, 3) for x in xs], '000000', 19050, None, None))
    ymin, ymax = float(np.floor((d.ROAq.min() - 0.5) / 5) * 5), float(np.ceil((d.ROAq.max() + 0.5) / 5) * 5)
    xml, rows = W.scatter_chart(S, 'Muammoli kreditlar ulushi, %', 'Aktivlar rentabelligi (ROA), %',
                                0, xmax, 2, ymin, ymax, 5, xfmt='0', yfmt='0', legend=True)
    xml = xml.replace('<c:size val="7"/>', '<c:size val="5"/>')
    info = dict(a=a, b=b, r2=m.rsquared, n=int(m.nobs), r=d.S3.corr(d.ROAq), n1=len(g1), n0=len(g0),
                xmax=xmax, ymin=ymin, ymax=ymax)
    return (xml, rows), info


# ---------------- C) orqaga tekshiruv (4.2.3-jadval)
def chart_c():
    D = pd.read_csv(HK + 'tizim42.csv', index_col=0)
    act = D.NPL.loc['2021_I':]
    X = [round(2020 + 0.25 * (i + 1), 2) for i in range(len(act))]          # 2021-y. I chorak = 2020,25 … IV chorak = 2021 (4.2.1-rasm bilan bir xil)
    FC = {2023: [3.71, 3.84, 3.77, 3.65], 2024: [3.50, 3.40, 3.43, 3.48], 2025: [3.99, 3.98, 3.99, 3.98]}
    yend = {y: round(float(act.iloc[X.index(float(y))]), 2) for y in (2022, 2023, 2024)}
    S = [('Haqiqiy qiymat', X, [round(float(v), 2) for v in act], '000000', 25400, None, 'circle')]
    hide = []
    for j, (y, v) in enumerate(FC.items()):
        o = y - 1
        S.append(('Satellit model prognozi (o‘rtacha xato 0,21–0,73 f.p.)', [float(o)] + [o + 0.25 * k for k in (1, 2, 3, 4)],
                  [yend[o]] + v, 'C00000', 22225, 'dash', None))
        S.append(('O‘zgarishsiz holat prognozi (0,12–0,65 f.p.)', [float(o), float(y)], [yend[o]] * 2, '808080', 19050, 'sysDot', None))
        S.append(('Prognoz boshlanish nuqtasi', [float(o), float(o)], [2.0, 6.0], 'A6A6A6', 9525, 'dash', None))
        if j:
            hide += [len(S) - 3, len(S) - 2, len(S) - 1]
    xml, rows = W.scatter_chart(S, 'Yillar (choraklar)', 'Muammoli kreditlar ulushi, %', 2020, 2025.25, 1, 2, 6, 1,
                                xfmt='[&lt;2020.5]&quot;&quot;;0', yfmt='0.0', legend=True)
    xml = xml.replace('<c:size val="7"/>', '<c:size val="5"/>')
    xml = xml.replace('<c:legendPos val="b"/>', '<c:legendPos val="b"/>' + ''.join(
        f'<c:legendEntry><c:idx val="{k}"/><c:delete val="1"/></c:legendEntry>' for k in hide), 1)
    return (xml, rows), {}


if __name__ == '__main__':
    dd = docx.Document()
    sec = dd.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin, sec.right_margin, sec.top_margin, sec.bottom_margin = Cm(3), Cm(1.5), Cm(2), Cm(2)
    st = dd.styles['Normal']; st.font.name = 'Times New Roman'; st.font.size = Pt(14)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    st.paragraph_format.space_after = Pt(0); st.paragraph_format.line_spacing = 1.5

    def par(text, bold=False, center=False, size=None, indent=True):
        p = dd.add_paragraph(); r = p.add_run(text); r.bold = bold
        if size: r.font.size = Pt(size)
        if center: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            if indent: p.paragraph_format.first_line_indent = Cm(1.25)
        return p

    def chart(xr, h, name):
        q = dd.add_paragraph(); q.alignment = WD_ALIGN_PARAGRAPH.CENTER
        q.paragraph_format.keep_with_next = True
        rid = W.add_chart(dd, xr)
        q._p.append(etree.fromstring(W.chart_run_xml(rid, 16.0, h, name)))

    A, ia = chart_a(); B, ib = chart_b(); C, _ = chart_c()
    print('A', {k: round(v, 3) for k, v in ia.items()}); print('B', {k: round(v, 3) for k, v in ib.items()})
    par('4.2-§ uchun model chizmalari (namuna)', bold=True, center=True)

    par('A. Satellit modelning asosi – nuqtalar va regressiya chizig‘i (4.2.2-jadvaldan oldin)', bold=True, indent=False)
    par(f'Muammoli kreditlar ulushi va to‘rt chorak oldingi YaIM o‘sishi o‘rtasidagi bog‘liqlik 4.2.1-rasmda keltirilgan. '
        f'Rasmdan ko‘rinib turibdiki, nuqtalar manfiy qiyalikka ega regressiya chizig‘i atrofida joylashgan '
        f'(korrelyatsiya koeffitsiyenti {f2(ia["r"])}, 1 foiz punkt sekinroq o‘sish keyingi yili ulushni {f2(-ia["b"])} foiz punktga oshiradi). '
        f'Eng yuqori qiymatlar 2020-yildagi pandemiya shokidan keyingi 2021-yil choraklariga to‘g‘ri keladi.')
    chart(A, 9.0, 'A')
    par('4.2.1-rasm. Muammoli kreditlar ulushi va to‘rt chorak oldingi YaIM o‘sish sur’ati o‘rtasidagi bog‘liqlik, 2021–2025-yillar', bold=True, center=True)
    par(f'Izoh: har bir nuqta – bir chorak; regressiya chizig‘i y = {f2(ia["a"])} − {f2(-ia["b"], 3)}x (R² = {f2(ia["r2"])}, n = {ia["n"]}); '
        f'oldingi chorak darajasi hisobga olingan baho (−0,151) 4.2.2-jadvalda.', size=12, indent=False)

    par('B. Bank darajasida kredit riski va daromadlilik – nuqtalar buluti (4.2.2-jadvaldan keyin)', bold=True, indent=False)
    par(f'Bank darajasidagi uchinchi model asosidagi bog‘liqlik 4.2.2-rasmda keltirilgan. Unda 25 ta bankning {ib["n"]} ta choraklik kuzatuvi '
        f'aks ettirilgan. Nuqtalar keng sochilgan bo‘lsa-da, muammoli kreditlar ulushi yuqori bo‘lgan banklarda aktivlar rentabelligi '
        f'past ekani yaqqol ko‘rinadi (korrelyatsiya koeffitsiyenti {f2(ib["r"])}).')
    chart(B, 9.5, 'B')
    par('4.2.2-rasm. Tanlangan 25 ta bankda muammoli kreditlar ulushi va aktivlar rentabelligi o‘rtasidagi bog‘liqlik, 2022–2025-yillar', bold=True, center=True)
    par(f'Izoh: har bir nuqta – bir bankning bir choragi; ko‘rsatkichlar 1 va 99 foizlik chegaralarda vinzorlangan; regressiya chizig‘i barcha kuzatuvlar bo‘yicha '
        f'y = {f2(ib["a"])} − {f2(-ib["b"])}x (R² = {f2(ib["r2"])}); bank effektlari hisobga olingan baho (−0,201) 4.2.2-jadvalda.', size=12, indent=False)

    par('C. Orqaga tekshiruv – haqiqiy va prognoz qiymatlar (4.2.3-jadvaldan keyin)', bold=True, indent=False)
    chart(C, 9.0, 'C')
    par('4.2.3-rasm. Bank tizimida muammoli kreditlar ulushining haqiqiy qiymatlari va orqaga tekshiruvdagi bir yil oldinga prognozlar, %', bold=True, center=True)
    par('Izoh: har bir prognoz oldingi yil oxirigacha bo‘lgan ma’lumotlar asosida baholangan model va haqiqiy YaIM o‘sishi bilan tuzilgan; '
        'qavs ichida 2023–2025-yillardagi o‘rtacha mutlaq xato (4.2.3-jadval).', size=12, indent=False)
    dd.save(DST)
    from pardoz import pardoz
    print(pardoz(DST), DST)
