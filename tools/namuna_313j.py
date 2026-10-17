"""3.1.3-jadval: uch manba nomlandi, daromadlilik/likvidlik/kapital vaznlari o‘zgarishi asoslandi (namuna)."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
import docx
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
DST = '/home/user/tarjima/Namuna_3.1.3-jadval.docx'
# (paragraf boshi, after-matn, qo‘shiladigan matn)
INS1 = ('uch manba', ' – Bazel III va Xalqaro valyuta jamg‘armasi ko‘rsatkichlar tizimidagi bloklar ustuvorligi, ikkinchi bob '
        'natijalari hamda O‘zbekistonda CAMEL tizimi uchun T.M. Karaliyev va I.F. Sayfiddinov keltirgan vaznlar nisbati –')
REP2 = ('ikki jihatda', 'uch jihatda')
INS3 = ('dollarlashuv darajasi yuqoriligiga mos keladi.',
        ' Uchinchidan, daromadlilik blokining vazni 0,25 dan 0,15 ga kamaytirildi, bo‘shagan vazn esa kapital (0,23 dan 0,25 ga) '
        'va likvidlik (0,15 dan 0,20 ga) bloklariga taqsimlandi. Buning uch sababi bor: birinchidan, davlat banklarida foyda '
        'davlat ko‘magi va imtiyozli resurslar hisobiga shakllanib, bankning o‘z barqarorligini to‘liq aks ettirmaydi (3.3-§); '
        'ikkinchidan, daromadlilik ko‘rsatkichlari ikki hisobot o‘rtasida eng ko‘p farq qiladigan ko‘rsatkichlardan biri: '
        '“Aloqabank” ATB da besh yillik sof foyda milliy hisobotda MHXS dagidan 1,8 barobar ko‘p, “Mikrokreditbank” ATB da esa '
        'zarar ikki hisobotda turli yillarda tan olingan (2.3-§); uchinchidan, foyda taqsimlanmagan foyda orqali kapital blokida '
        'ham qisman aks etadi. Likvidlik vaznining oshirilishi esa tadqiqot obyektlarida kreditlarning depozitlarga nisbati '
        '200 foizdan yuqori bo‘lib, kreditlar asosan tashqi qarzlar hisobiga moliyalashtirilishi bilan izohlanadi.')
CELL = ('Kapitalni ichki manbadan to‘ldirish imkoniyati',
        '; davlat banklarida davlat ko‘magi ta’sirida va ikki hisobotda keskin farq qiladi')

if __name__ == '__main__':
    src = docx.Document(F)
    ps = src.paragraphs
    p1 = [p for p in ps if 'uch manba asosida asoslandi' in p.text][0].text
    p2 = [p for p in ps if p.text.startswith('Jadval ma’lumotlari shuni ko‘rsatadiki, taklif etilayotgan vaznlar')][0].text
    cap = [p for p in ps if p.text.startswith('Moliyaviy barqarorlikning integral indeksi bloklari vaznlari')][0]
    rows = [[c.text for c in r.cells] for r in docx.table.Table(cap._p.getnext(), src).rows]

    dd = docx.Document()
    sec = dd.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin, sec.right_margin, sec.top_margin, sec.bottom_margin = Cm(3), Cm(1.5), Cm(2), Cm(2)
    st = dd.styles['Normal']; st.font.name = 'Times New Roman'; st.font.size = Pt(14)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    st.paragraph_format.space_after = Pt(0)

    def run(q, t, size=14, bold=False, red=False, strike=False):
        r = q.add_run(t); r.font.size = Pt(size); r.bold = bold; r.font.strike = strike
        if red: r.font.color.rgb = RGBColor(0xFF, 0, 0)
        if strike: r.font.color.rgb = RGBColor(0x7F, 0x7F, 0x7F)

    def body():
        q = dd.add_paragraph(); q.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        q.paragraph_format.first_line_indent = Cm(1.25); q.paragraph_format.line_spacing = 1.5
        return q

    def note(t):
        q = dd.add_paragraph(); run(q, t, 12, bold=True)

    note('1) Jadvaldan oldingi xatboshi (oxirgi qismi):')
    q = body(); i = p1.index('Shu sababli'); a, b = p1[i:].split(INS1[0], 1)
    run(q, '…' + a + INS1[0]); run(q, INS1[1], red=True); run(q, b)
    note('2) 3.1.3-jadval, “Asoslash” ustuni:')
    t = dd.add_table(rows=len(rows), cols=5); t.style = 'Table Grid'
    for i, row in enumerate(rows):
        for j, txt in enumerate(row):
            q = t.cell(i, j).paragraphs[0]
            if j in (1, 2, 3) or i == 0: q.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run(q, txt, 11, bold=(i == 0))
            if j == 4 and txt == CELL[0]:
                run(q, CELL[1], 11, red=True)
    note('3) Jadvaldan keyingi xatboshi:')
    q = body()
    a, b = p2.split(REP2[0], 1)
    run(q, a); run(q, REP2[0], strike=True); run(q, ' '); run(q, REP2[1], red=True)
    c, d = b.split(INS3[0], 1)
    run(q, c + INS3[0]); run(q, INS3[1], red=True); run(q, d)
    dd.save(DST)
    from pardoz import pardoz
    print(pardoz(DST), DST)
