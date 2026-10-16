"""Safaraliyev B. avtoreferati (“O‘zsanoatqurilishbank” ATB muammoli aktivlari prognozi) bo‘yicha:
1) prognozni qayta hisoblash, 2) taqriz uchun fikr-mulohazalar, 3) tuzatilgan prognoz namunasi (ulush + makro ssenariy)."""
import json, sys
sys.path.insert(0, '/home/user/tarjima/tools')
import numpy as np
import statsmodels.api as sm
import docx
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from lxml import etree
import wordart as W

DST = '/home/user/tarjima/Namuna_Safaraliyev_prognoz.docx'
YRS = list(range(2018, 2026))
NPA = [558093000, 437583650, 456482825, 1723656083, 1764995550, 1917453217, 2616418810, 2615903007]   # ming so‘m, 4-jadval
AUTH = {2026: (4109805792, 3103632008, 2097458224), 2027: (4535610459, 3457478395, 2379346331),
        2028: (4969922393, 3811324782, 2652727171)}                                               # 6-jadval: pess., dinamik, optim.
AKTIV = {2021: 56.5, 2022: 64.3, 2023: 74.6, 2024: 87.6, 2025: 102.4}                               # trln so‘m (bank hisobotlari)
AKTIV28 = 128.4                                                                                     # muallif prognozi (8-xulosa)
FC = json.load(open('/home/user/tarjima/manbalar/hisob_kitob/fc42.json'))
SYS = {k: v[3::4] for k, v in FC['tizim_NPL_yo‘li'].items()}                                        # tizim NPL, 2026–2028 yil oxiri
SYS25 = 2.99
GDP = FC['ssenariy']


def c(x, k=2):
    return f'{x:,.{k}f}'.replace(',', ' ').replace('.', ',').replace('-', '−')


# ---------------- 1) qayta hisoblash
y = np.array(NPA, float); t = np.arange(1, 9)
m = sm.OLS(y, sm.add_constant(t)).fit()
pr = m.get_prediction(sm.add_constant(np.array([9, 10, 11]), has_constant='add')).summary_frame(alpha=0.05)
fit25 = m.fittedvalues[-1] / 1e9
ly = sm.OLS(np.log(y), sm.add_constant(t)).fit()
log28 = float(np.exp(ly.predict([[1, 11]]))[0]) / 1e9
m2 = sm.OLS(y[3:], sm.add_constant(t[3:])).fit()
tr2128 = float(m2.predict([[1, 11]])[0]) / 1e9

# ---------------- 3) tuzatilgan prognoz
g = (AKTIV28 / AKTIV[2025]) ** (1 / 3) - 1
akt = {2025: AKTIV[2025], 2026: AKTIV[2025] * (1 + g), 2027: AKTIV[2025] * (1 + g) ** 2, 2028: AKTIV28}
r25 = NPA[-1] / 1e9 / AKTIV[2025] * 100
ratio_hist = {yy: NPA[YRS.index(yy)] / 1e9 / AKTIV[yy] * 100 for yy in AKTIV}
SC = ['Ijobiy', 'Asosiy', 'Salbiy', 'Keskin salbiy']
RAT = {s: [r25 * SYS[s][i] / SYS25 for i in range(3)] for s in SC}                  # proporsional taqsimlash
AMT = {s: [RAT[s][i] / 100 * akt[2026 + i] for i in range(3)] for s in SC}
const28 = r25 / 100 * AKTIV28


def chart():
    xs_a = YRS; ys_a = [round(v / 1e9, 3) for v in NPA]
    S = [('Haqiqiy (4-jadval)', xs_a, ys_a, '000000', 28575, None, 'circle'),
         ('Avtoreferat: chiziqli trend (dinamik)', [2025, 2026, 2027, 2028], [round(fit25, 3)] + [round(AUTH[k][1] / 1e9, 3) for k in (2026, 2027, 2028)], '808080', 22225, 'sysDash', 'triangle'),
         ('Avtoreferat: 95 % oraliq chegaralari', [2026, 2027, 2028], [round(AUTH[k][0] / 1e9, 3) for k in (2026, 2027, 2028)], 'BFBFBF', 12700, 'dash', None),
         ('', [2026, 2027, 2028], [round(AUTH[k][2] / 1e9, 3) for k in (2026, 2027, 2028)], 'BFBFBF', 12700, 'dash', None)]
    col = {'Ijobiy': '70AD47', 'Asosiy': '2A78D6', 'Salbiy': 'FFC000', 'Keskin salbiy': 'C00000'}
    for s in ('Keskin salbiy', 'Salbiy', 'Asosiy', 'Ijobiy'):
        S.append((f'Tuzatilgan: {s.lower()} ssenariy', [2025, 2026, 2027, 2028], [round(NPA[-1] / 1e9, 3)] + [round(v, 3) for v in AMT[s]],
                  col[s], 25400, None, None))
    xml, rows = W.scatter_chart(S, 'Yillar', 'Muammoli aktivlar, trln so‘m', 2018, 2028, 1, 0, 6, 1, xfmt='0', yfmt='0.0', legend=True)
    xml = xml.replace('<c:size val="7"/>', '<c:size val="5"/>')
    xml = xml.replace('<c:legendPos val="b"/>', '<c:legendPos val="b"/><c:legendEntry><c:idx val="3"/><c:delete val="1"/></c:legendEntry>', 1)
    return xml, rows


def table(dd, rows, widths, hdr_rows=1, size=11):
    tb = dd.add_table(rows=len(rows), cols=len(rows[0])); tb.style = 'Table Grid'; tb.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, r in enumerate(rows):
        for j, v in enumerate(r):
            cell = tb.cell(i, j); cell.width = Cm(widths[j]); p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if (i < hdr_rows or j > 0) else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.line_spacing = 1.0; p.paragraph_format.first_line_indent = Cm(0)
            run = p.add_run(str(v)); run.font.size = Pt(size); run.bold = i < hdr_rows
    return tb


if __name__ == '__main__':
    dd = docx.Document()
    sec = dd.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin, sec.right_margin, sec.top_margin, sec.bottom_margin = Cm(3), Cm(1.5), Cm(2), Cm(2)
    st = dd.styles['Normal']; st.font.name = 'Times New Roman'; st.font.size = Pt(14)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    st.paragraph_format.space_after = Pt(0); st.paragraph_format.line_spacing = 1.5

    def par(text, bold=False, center=False, size=None, indent=True, italic=False):
        p = dd.add_paragraph(); r = p.add_run(text); r.bold = bold; r.italic = italic
        if size: r.font.size = Pt(size); p.paragraph_format.line_spacing = 1.0
        if center: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            if indent: p.paragraph_format.first_line_indent = Cm(1.25)
        return p

    par('B. Safaraliyev avtoreferatidagi “O‘zsanoatqurilishbank” ATB muammoli aktivlari prognozi bo‘yicha fikr-mulohazalar va tuzatilgan prognoz (namuna)',
        bold=True, center=True)

    # 1
    par('1. Prognozni qayta hisoblash natijalari', bold=True, indent=False)
    par('Avtoreferatning 4-jadvalidagi ma’lumotlar asosida chiziqli trend modeli qayta baholandi. Barcha ko‘rsatkichlar avtoreferatdagi natijalarga '
        'to‘liq mos keladi, ya’ni hisob-kitoblarda arifmetik xato yo‘q. Shu bilan birga, “pessimistik” va “optimistik” ssenariylar prognozning '
        '95 foizlik ishonch oralig‘i chegaralari ekani aniqlandi (1-jadval).')
    par('1-jadval. Avtoreferatdagi va qayta hisoblangan natijalar', bold=True, center=True, size=12)
    rows = [['Ko‘rsatkich', 'Avtoreferatda', 'Qayta hisoblangan'],
            ['Trend koeffitsiyenti, mlrd so‘m', '354', c(m.params[1] / 1e6, 1)],
            ['Konstanta, mlrd so‘m', '−81', c(m.params[0] / 1e6, 1)],
            ['R² / F-statistika', '0,8928 / 49,98', f'{c(m.rsquared, 4)} / {c(m.fvalue, 2)}'],
            ['Durbin–Watson', '2,4622', c(2.4622, 4)]]
    for k in (2026, 2027, 2028):
        i = k - 2026
        rows.append([f'{k}-yil: dinamik / pessimistik / optimistik, trln so‘m',
                     ' / '.join(c(v / 1e9) for v in (AUTH[k][1], AUTH[k][0], AUTH[k][2])),
                     f'{c(pr["mean"].iloc[i] / 1e9)} / {c(pr["obs_ci_upper"].iloc[i] / 1e9)} / {c(pr["obs_ci_lower"].iloc[i] / 1e9)}'])
    table(dd, rows, [7.5, 4.2, 4.3])
    par('Izoh: qayta hisoblashda pessimistik va optimistik qiymatlar chiziqli trend prognozining 95 foizlik ishonch oralig‘ining yuqori va quyi '
        'chegaralari sifatida olingan.', size=12, indent=False)

    # 2
    par('2. Taqriz uchun fikr-mulohazalar', bold=True, indent=False)
    R = [
        ('Ssenariylar tavsifi.', ' 6-jadvaldagi “pessimistik” va “optimistik” ssenariylar chiziqli trend prognozining 95 foizlik ishonch oralig‘i '
         f'chegaralariga aynan mos keladi (masalan, 2026-yil uchun {c(AUTH[2026][2] / 1e9)}–{c(AUTH[2026][0] / 1e9)} trln so‘m). Ular YaIM o‘sishi, '
         'foiz stavkasi yoki almashuv kursi kabi makroiqtisodiy farazlarga asoslanmagan, shu sababli ssenariy emas, balki statistik noaniqlik oralig‘idir. '
         'Buni matnda ochiq ko‘rsatish yoki ssenariylarni makroiqtisodiy shartlar bilan bog‘lash tavsiya etiladi.'),
        ('Model spetsifikatsiyasi.', ' 2018–2025-yillar ma’lumotlari chiziqli emas, balki pog‘onali o‘zgarishni aks ettiradi. Muammoli aktivlar '
         '2018–2020-yillarda 0,44–0,56 trln so‘m, 2021-yilda 1,72 trln so‘m (3,8 barobar), 2024–2025-yillarda 2,62 trln so‘m bo‘lgan. '
         'Determinatsiya koeffitsiyentining yuqoriligi (0,89) asosan 2021-yildagi bir martalik sakrash hisobiga shakllangan. Tarkibiy siljishni '
         'hisobga olish, masalan, 2021-yildan keyingi davr uchun alohida baholash yoki fiktiv o‘zgaruvchi kiritish maqsadga muvofiq.'),
        ('Kuzatuvlar soni.', ' Model 8 ta yillik kuzatuvga asoslangan. Bunday kichik tanlamada Shapiro–Wilk, Breusch–Godfrey va Ramsey RESET '
         'testlarining quvvati past, shu sababli ularning natijasi modelning prognoz uchun ishonchliligini tasdiqlash uchun yetarli dalil bo‘la olmaydi.'),
        ('Prognozning boshlang‘ich nuqtasi.', f' Prognoz 2025-yilning haqiqiy qiymatidan ({c(NPA[-1] / 1e9)} trln so‘m) emas, balki trend '
         f'chizig‘idagi qiymatdan ({c(fit25)} trln so‘m) davom ettirilgan. Natijada 2026-yil uchun dinamik ssenariy 2025-yilga nisbatan '
         f'{c((AUTH[2026][1] / NPA[-1] - 1) * 100, 1)} foiz o‘sishni ko‘rsatadi, holbuki 2024–2025-yillarda muammoli aktivlar deyarli o‘zgarmagan. '
         'Bu uzilish 8-rasmda ham ko‘rinadi.'),
        ('Nominal ko‘rsatkich.', ' Muammoli aktivlarning nominal hajmi bank balansi va inflyatsiya bilan birga o‘sadi. Aktivlar sifatini baholash uchun '
         'nisbiy ko‘rsatkich, ya’ni muammoli aktivlarning jami aktivlar yoki kreditlardagi ulushi prognoz qilinishi lozim. Muallifning o‘z prognozlari '
         f'bo‘yicha bu ulush 2025-yildagi {c(r25)} foizdan 2028-yilda {c(AUTH[2028][1] / 1e9 / AKTIV28 * 100)} foizga oshadi, xolos. Bu 8-xulosadagi '
         '“muammoli aktivlarning keskin ortishi” degan fikrni to‘liq tasdiqlamaydi.'),
        ('Muqobil modellar bilan taqqoslanmagan.', ' Chiziqli trend “maqbul model” deb baholangan, biroq u boshqa modellar bilan solishtirilmagan. '
         'Model o‘zgartirilganda 2028-yil prognozi keskin farq qiladi (2-jadval), ya’ni natija ma’lumotlardan ko‘ra model tanloviga ko‘proq bog‘liq.'),
        ('Texnik kamchiliklar.', ' 26-betda “Quyida berilgan 6-rasmda” deyilgan, biroq rasm 8-rasm. 8-rasmning vertikal o‘qidagi qiymatlar bir-birining '
         'ustiga tushib o‘qib bo‘lmaydi, ishonch oralig‘i esa rasmda ko‘rsatilmagan. 4-jadval dastlabki ma’lumotlardan iborat, biroq uning manbasi '
         '“Stata dasturida muallif hisob-kitoblari” deb ko‘rsatilgan, holbuki manba bank hisobotlari bo‘lishi kerak. 8-xulosadagi aktivlar (128,4 trln so‘m) '
         'va soliqqacha foyda (2,49 trln so‘m) prognozlarining usuli avtoreferatda keltirilmagan.')]
    for k, (h, b) in enumerate(R, 1):
        p = dd.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY; p.paragraph_format.first_line_indent = Cm(1.25)
        r = p.add_run(f'{k}. {h}'); r.bold = True; p.add_run(b)
    par('2-jadval. 2028-yil uchun muammoli aktivlar prognozining modelga bog‘liqligi', bold=True, center=True, size=12)
    table(dd, [['Yondashuv', '2028-yil, trln so‘m'],
               ['Chiziqli trend, 2018–2025 (avtoreferat)', c(AUTH[2028][1] / 1e9)],
               ['Chiziqli trend, faqat 2021–2025', c(tr2128)],
               [f'Aktivlardagi ulush 2025-yil darajasida ({c(r25)} %), aktivlar 128,4 trln so‘m', c(const28)],
               [f'Logarifmik trend (yiliga {c((np.exp(ly.params[1]) - 1) * 100, 1)} % o‘sish)', c(log28)]], [11.0, 4.5])

    # 3
    par('3. Tuzatilgan prognoz namunasi', bold=True, indent=False)
    par('Tuzatilgan prognoz quyidagi tartibda tuziladi. Birinchidan, nominal hajm o‘rniga muammoli aktivlarning jami aktivlardagi ulushi prognoz qilinadi. '
        f'2021–2025-yillarda bu ulush {c(min(ratio_hist.values()))}–{c(max(ratio_hist.values()))} foiz oralig‘ida bo‘lib, o‘sish tendensiyasiga ega emas. '
        'Ikkinchidan, ulushning o‘zgarishi YaIM o‘sishi bo‘yicha makroiqtisodiy ssenariylarga bog‘lanadi. Buning uchun bank tizimi darajasida baholangan '
        'satellit model (muammoli kreditlar ulushi = 2,98 + 0,49 × oldingi chorak ulushi − 0,15 × to‘rt chorak oldingi YaIM o‘sishi) qo‘llanib, '
        'tizim bo‘yicha og‘ish bankka uning 2025-yildagi ulushiga proporsional taqsimlanadi. Uchinchidan, muammoli aktivlar hajmi ulushni aktivlarga '
        f'ko‘paytirish orqali topiladi, aktivlar esa muallifning o‘z prognozi bo‘yicha olinadi (2028-yilda 128,4 trln so‘m, yiliga {c(g * 100, 1)} foiz). '
        'To‘rtinchidan, prognoz 2025-yilning haqiqiy qiymatidan boshlanadi.')
    par('3-jadval. “O‘zsanoatqurilishbank” ATB muammoli aktivlarining makroiqtisodiy ssenariylar bo‘yicha tuzatilgan prognozi', bold=True, center=True, size=12)
    rows = [['Ssenariy', 'YaIM o‘sishi, %: 2026 / 2027 / 2028', 'Aktivlardagi ulush, %: 2026 / 2027 / 2028', 'Muammoli aktivlar, trln so‘m: 2026 / 2027 / 2028']]
    for s in SC:
        rows.append([s, ' / '.join(c(GDP[s][str(k)], 1) for k in (2026, 2027, 2028)), ' / '.join(c(v) for v in RAT[s]), ' / '.join(c(v) for v in AMT[s])])
    rows.append(['Avtoreferat (dinamik)', '–', ' / '.join(c(AUTH[k][1] / 1e9 / akt[k] * 100) for k in (2026, 2027, 2028)),
                 ' / '.join(c(AUTH[k][1] / 1e9) for k in (2026, 2027, 2028))])
    table(dd, rows, [3.3, 4.2, 4.2, 4.3], size=11)
    par(f'Izoh: 2025-yil – {c(NPA[-1] / 1e9)} trln so‘m, aktivlardagi ulushi {c(r25)} %; YaIM ssenariylari va satellit model – O.T. Maxmudov '
        'dissertatsiyasi, 4.2-paragraf; 2026-yilda ssenariylar bir xil, chunki YaIM shokining ta’siri to‘rt chorakdan keyin namoyon bo‘ladi.', size=12, indent=False)
    q = dd.add_paragraph(); q.alignment = WD_ALIGN_PARAGRAPH.CENTER; q.paragraph_format.keep_with_next = True
    rid = W.add_chart(dd, chart())
    q._p.append(etree.fromstring(W.chart_run_xml(rid, 16.0, 10.0, 'prognoz')))
    par('1-rasm. “O‘zsanoatqurilishbank” ATB muammoli aktivlari: haqiqiy qiymatlar, avtoreferat prognozi va tuzatilgan ssenariy prognozi', bold=True, center=True, size=12)
    a, k_, s_ = AMT['Asosiy'][2], AMT['Keskin salbiy'], AMT['Ijobiy'][2]
    par(f'Tuzatilgan prognozga ko‘ra, 2028-yilda muammoli aktivlar asosiy ssenariyda {c(a)} trln so‘mni, ijobiy ssenariyda {c(s_)} trln so‘mni tashkil etadi. '
        f'Keskin salbiy ssenariyda ular 2027-yilda {c(k_[1])} trln so‘mgacha oshib, 2028-yilda {c(k_[2])} trln so‘mgacha kamayadi. Avtoreferatdagi dinamik '
        f'prognoz ({c(AUTH[2028][1] / 1e9)} trln so‘m) asosiy va salbiy ssenariylar oralig‘ida joylashgan, biroq uning “pessimistik” chegarasi '
        f'({c(AUTH[2028][0] / 1e9)} trln so‘m) keskin salbiy ssenariydan ham yuqori, “optimistik” chegarasi ({c(AUTH[2028][2] / 1e9)} trln so‘m) esa '
        '2025-yil darajasidan deyarli farq qilmaydi. Demak, tuzatilgan yondashuv prognozni aniq makroiqtisodiy shartlar bilan bog‘laydi va '
        'muammoli aktivlarning o‘sishini bank balansining o‘sishidan ajratib ko‘rsatadi.')
    dd.save(DST)
    from pardoz import pardoz
    print(pardoz(DST), DST)
    print('r25', round(r25, 2), 'g', round(g * 100, 2), {s: [round(v, 2) for v in AMT[s]] for s in SC}, {s: [round(v, 2) for v in RAT[s]] for s in SC},
          {k: round(v, 2) for k, v in ratio_hist.items()})
