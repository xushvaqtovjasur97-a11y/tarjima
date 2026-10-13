"""1.1.6-rasmning yangi varianti (namuna): Z-score va integral indeksning O‘zbekiston banklari
muammolarini bashorat qilish aniqligi, 2021–2025-yillar (4.1-§ hodisa ta’rifi bilan)."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
import numpy as np, pandas as pd, statsmodels.api as sm, warnings
import docx
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from lxml import etree
import wordart as W
from pardoz import pardoz
warnings.filterwarnings('ignore')

HK = '/home/user/tarjima/manbalar/hisob_kitob/'
DST = '/home/user/tarjima/Namuna_1.1.6-rasm_yangi.docx'

p = pd.read_csv(HK + 'panel33.csv').sort_values(['bank', 't']); g = p.groupby('bank')
p['lossq'] = -p.NPq / g.Kapital.shift(1) * 100; p['EV'] = (p.lossq >= 5).astype(int)
for k in range(1, 5):
    p[f'ev{k}'] = g.EV.shift(-k).where(g.t.shift(-k) == p.t + k)
p['Y'] = p[[f'ev{k}' for k in range(1, 5)]].max(axis=1)
d = p[p.tanlov & (p.t <= 2024 * 4 + 3) & (p.EV == 0)].dropna(subset=['Y', 'MBII', 'lnZ'])
N, POS, NB = len(d), int(d.Y.sum()), d.bank.nunique()


def best(y, s):
    """Youden mezoni bo‘yicha eng yaxshi chegara: (muammoli to‘g‘ri, sog‘lom to‘g‘ri), %."""
    out = None
    for th in np.unique(s):
        pr = s >= th
        tpr = (pr & (y == 1)).sum() / (y == 1).sum(); tnr = (~pr & (y == 0)).sum() / (y == 0).sum()
        if out is None or tpr + tnr > out[0] + out[1]:
            out = (tpr, tnr)
    return [int(100 * x + 0.5) for x in out]


y = d.Y.values
RZ = best(y, -d.lnZ.values)                      # past Z-score – xavfli
RM = best(y, -d.MBII.values)                     # past MBII – xavfli
lc = sm.Logit(d.Y, sm.add_constant(d[['MBII', 'lnZ']])).fit(disp=0)
RC = best(y, lc.predict(sm.add_constant(d[['MBII', 'lnZ']])).values)
lz = sm.Logit(d.Y, sm.add_constant(d[['lnZ']])).fit(disp=0)
PZ, PZC = lz.pvalues['lnZ'], lc.pvalues['lnZ']
print(N, POS, NB, RZ, RM, RC, round(PZ, 3), round(PZC, 3), d.davr.min(), d.davr.max())
f = lambda x: f'{x:.2f}'.replace('.', ',')

CATS = ['Faqat Z-score', 'Integral indeks (MBII)', 'MBII + Z-score']
SER = [('Muammoli banklar to‘g‘ri aniqlangan', [RZ[0], RM[0], RC[0]], 'EB6834'),
       ('Sog‘lom banklar to‘g‘ri aniqlangan', [RZ[1], RM[1], RC[1]], '2A78D6')]
xml, rows = W.bar_chart(CATS, SER, ytitle='To‘g‘ri tasniflash ulushi, %', ymax=100, ymin=0, major=20, label_fmt='0', gap=80)


def texts(fn):
    s139 = ('Bu ko‘rsatkichning amaliy samaradorligini L. Chiaramonte va F. Poli tadqiqoti tasdiqlaydi: ular 12 mamlakatdagi 3125 ta bankning 2001–2011-yillardagi ma’lumotlari asosida faqat Z-score ko‘rsatkichining o‘zi muammoga uchragan banklarning 75 foizini, sog‘lom banklarning esa 81 foizini to‘g‘ri aniqlaganini ko‘rsatgan'
            + fn('Chiaramonte L., Poli F. Predicting European Bank Distress: Evidence from the Recent Financial Crisis // Governance, Regulation and Bank Stability / ed. by T. Lindblom, S. Sjögren, M. Willesson. – Basingstoke: Palgrave Macmillan, 2014. – P. 83–95.')
            + '. Biroq O‘zbekiston bank tizimining 2021–2025-yillardagi ma’lumotlari bu xulosani tasdiqlamaydi (1.1.6-rasm).')
    caption = 'Z-score ko‘rsatkichi va integral indeksning O‘zbekiston tijorat banklari muammolarini bashorat qilish aniqligi, 2021–2025-yillar, %'
    cap_fn = (f'Tadqiqot natijasida muallif tomonidan {NB} ta tijorat bankining 2021–2025-yillardagi choraklik hisobotlari asosida hisoblandi ({N} ta kuzatuv, {POS} ta muammo hodisasi). Muammo hodisasi – keyingi to‘rt chorak ichida choraklik zararning kapitalning 5 foizidan oshishi (4.1-§); Z-score aktivlar rentabelligining sakkiz chorak bo‘yicha standart og‘ishi asosida hisoblangani sababli kuzatuvlar 2022-yil IV choragidan boshlanadi; har bir model uchun chegaraviy qiymat Youden mezoni bo‘yicha tanlangan.')
    a = (f'Rasm ma’lumotlari shuni ko‘rsatadiki, O‘zbekiston banklarida faqat Z-score ko‘rsatkichining o‘zi muammoga uchragan banklarning {RZ[0]} foizini, sog‘lom banklarning {RZ[1]} foizini to‘g‘ri aniqlaydi, ya’ni uning bashorat kuchi tasodifiy tasniflashdan deyarli farq qilmaydi. Logit modelida ham Z-score logarifmi statistik ahamiyatsiz (p = {f(PZ)}), integral indeksga qo‘shilganda esa uning bashorat kuchini oshirmaydi (p = {f(PZC)}). 3.1-paragrafda taklif etilgan integral indeks esa xuddi shu ma’lumotlarda muammoli banklarning {RM[0]} foizini, sog‘lom banklarning {RM[1]} foizini to‘g‘ri aniqlaydi. Bu farqning sababi shundaki, O‘zbekiston banklarida, ayniqsa davlat ishtirokidagi banklarda, kapitalning aktivlarga nisbati davlat tomonidan kapitallashtirish hisobiga yuqori saqlanadi, aktivlar sifatidagi muammolar esa aktivlar rentabelligi o‘zgaruvchanligida kechikib aks etadi. Natijada faqat kapital va rentabellikka asoslangan Z-score yashirin kredit riskini sezmaydi. Bundan kelib chiqadiki, aksiyalar bozori rivojlanmagan O‘zbekiston sharoitida buxgalteriya ma’lumotlariga asoslangan baholash bitta ko‘rsatkich bilan cheklanmasligi, balki aktivlar sifati blokini o‘z ichiga olgan integral baholash usuliga tayanishi lozim.')
    return dict(s139=s139, caption=caption, cap_fn=cap_fn, a=a)


if __name__ == '__main__':
    dd = docx.Document()
    sec = dd.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin, sec.right_margin, sec.top_margin, sec.bottom_margin = Cm(3), Cm(1.5), Cm(2), Cm(2)
    st = dd.styles['Normal']
    st.font.name = 'Times New Roman'; st.font.size = Pt(14)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    st.paragraph_format.line_spacing = 1.5; st.paragraph_format.space_after = Pt(0)

    def par(text, bold=False, italic=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY, indent=True, size=14, color=None):
        q = dd.add_paragraph(); q.alignment = align
        q.paragraph_format.first_line_indent = Cm(1.25) if indent else Cm(0)
        r = q.add_run(text); r.bold = bold; r.italic = italic; r.font.size = Pt(size)
        if color:
            r.font.color.rgb = RGBColor.from_string(color)

    FN = []

    def fn(t):
        FN.append(t); return f'[{len(FN)}]'

    T = texts(fn)
    par('1.1.6-rasmning yangi varianti (namuna)', bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False)
    par('Ma’qul bo‘lsa: 1.1-§ dagi “Bu ko‘rsatkichning amaliy samaradorligini L. Chiaramonte va F. Poli tadqiqoti yaqqol tasdiqlaydi (1.1.6-rasm).” jumlasi quyidagi birinchi xatboshi bilan, 1.1.6-rasm va undan keyingi “Mualliflar 12 mamlakatdagi …” xatboshi esa quyidagi rasm va tahlil bilan almashtiriladi.',
        italic=True, size=12, indent=False, color='595959')
    par('… ' + T['s139'])
    q = dd.add_paragraph(); q.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rid = W.add_chart(dd, (xml, rows))
    q._p.append(etree.fromstring(W.chart_run_xml(rid, 16.0, 8.5, '1.1.6-rasm')))
    FN.append(T['cap_fn']); par('1.1.6-rasm. ' + T['caption'] + f'[{len(FN)}]', bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False)
    par(T['a'])
    par('Sahifa osti snoskalari', bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False)
    for i, t in enumerate(FN, 1):
        par(f'[{i}] {t}', size=11, indent=False)
    dd.save(DST)
    print(pardoz(DST), DST)
