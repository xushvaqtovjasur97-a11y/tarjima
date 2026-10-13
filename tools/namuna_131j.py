"""1.3.1-jadval tuzatish namunasi (BCBS WP4 bo‘yicha)."""
import copy, sys, shutil
sys.path.insert(0, '/home/user/tarjima/tools')
import docx
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import kamlib as K
from pardoz import pardoz

SRC = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
DST = '/home/user/tarjima/Namuna_1.3.1-jadval_yangi.docx'
BCBS = 'Sahajwala R., Van den Bergh P. Supervisory risk assessment and early warning systems: Basel Committee on Banking Supervision Working Papers No. 4. – Basel: BIS, December 2000. – P. 4–5 (Table 1)'
# (qator kaliti, ustun, eski, yangi)
EDITS = [
    ('CAMELS (1978–1979)', 1, 'CAMELS (1978–1979)', 'CAMEL (1979), CAMELS (1997)'),
    ('FIMS, SEER, SCOR', 1, 'FIMS, SEER, SCOR', 'SEER (1993, avval FIMS), SCOR (1995)'),
    ('RATE (1998)', 1, 'RATE (1998)', 'RATE (1998–2013)'),
    ('ORAP (1997), SAABA', 1, 'ORAP (1997), SAABA', 'ORAP (1997), SAABA (1997)'),
    ('ORAP (1997), SAABA', 2, 'Reyting va ekspert tizimi', 'Reyting va statistik model'),
    ('ORAP (1997), SAABA', 3, None, 'SAABA kredit portfeli bo‘yicha kutilgan zararni baholaydi; 3 yillik to‘lov qobiliyati bahosi; har 6 oyda yangilanadi'),
    ('PATROL (1993)', 3, None, 'Masofaviy nazorat; 5 komponent (kapital, daromadlilik, kredit sifati, tashkiliy tuzilma, likvidlik), 1–5 ballik shkala'),
    ('RAST', 1, 'RAST', 'RAST (1999)'),
    ('RAST', 3, None, '4 bosqichli tahlil: risklarni tahlil qilish va nazorat choralarini belgilash'),
]


def table(D):
    return [tb for tb in D.d.tables if tb.rows[0].cells[0].text.strip() == 'Mamlakat' and any('CAMELS' in r.cells[1].text for r in tb.rows)][0]


def apply(D, tag='131J'):
    t = table(D)
    for key, c, old, new in EDITS:
        ri = [i for i, r in enumerate(t.rows) if r.cells[1].text.strip().startswith(key)][0]
        if old is None:
            D.set_cell(t, ri, c, new, tag=tag)
        else:
            D.cell_replace(t, ri, c, old, new, tag=tag)
    return t


if __name__ == '__main__':
    tmp = '/tmp/claude-0/-home-user-tarjima/315c132a-88b7-5e49-bb67-7f6a05e43e7d/scratchpad/tmp131.docx'
    shutil.copy(SRC, tmp)
    D = K.Doc(tmp)
    t = apply(D)
    changed = {}
    for key, c, o, n in EDITS:
        ri = [i for i, r in enumerate(t.rows) if r.cells[1].text.strip().startswith(n if c == 1 else key) or r.cells[1].text.strip().startswith(key)][0]
        changed.setdefault(ri, {})[c] = (o, n)
    data = [[cl.text for cl in r.cells] for r in t.rows]
    dd = docx.Document()
    sec = dd.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin, sec.right_margin, sec.top_margin, sec.bottom_margin = Cm(3), Cm(1.5), Cm(2), Cm(2)
    st = dd.styles['Normal']; st.font.name = 'Times New Roman'; st.font.size = Pt(14)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    st.paragraph_format.line_spacing = 1.5; st.paragraph_format.space_after = Pt(0)

    def par(text, bold=False, italic=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY, size=12, color=None):
        q = dd.add_paragraph(); q.alignment = align
        r = q.add_run(text); r.bold = bold; r.italic = italic; r.font.size = Pt(size)
        if color:
            r.font.color.rgb = RGBColor.from_string(color)
        return q

    par('1.3.1-jadval bo‘yicha tuzatishlar (namuna, o‘zgarishlar qizil rangda)', bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, size=14)
    par('1.3.1-jadval', italic=True, align=WD_ALIGN_PARAGRAPH.RIGHT, size=14)
    par('Rivojlangan mamlakatlarda banklarni baholash tizimlarining qiyosiy tavsifi', bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, size=14)
    tb = dd.add_table(rows=len(data), cols=4); tb.style = 'Table Grid'
    W_ = [Cm(2.6), Cm(3.6), Cm(3.4), Cm(6.9)]
    tb.autofit = False
    for gc, w in zip(tb._tbl.tblGrid.findall(qn('w:gridCol')), W_):
        gc.set(qn('w:w'), str(int(w.twips)))
    for ri, row in enumerate(data):
        for ci, txt in enumerate(row):
            cl = tb.cell(ri, ci); cl.width = W_[ci]
            q = cl.paragraphs[0]; q.paragraph_format.line_spacing = 1.0
            ch = changed.get(ri, {}).get(ci)
            if ch and ch[0] is not None:
                # eski matn ustidan chizilgan, yangi qizil
                base = data[ri][ci]
                r0 = q.add_run(base); r0.font.size = Pt(11); r0.font.color.rgb = RGBColor(0xFF, 0, 0)
            elif ch:
                r0 = q.add_run(txt); r0.font.size = Pt(11); r0.font.color.rgb = RGBColor(0xFF, 0, 0)
            else:
                r0 = q.add_run(txt); r0.font.size = Pt(11); r0.bold = ri == 0
    par('')
    par('Jadval manbasi (snoska) – birlamchi manba boshiga qo‘shiladi:')
    par('Tadqiqot natijasida muallif tomonidan quyidagi manbalar asosida ishlab chiqildi: ' + BCBS + '; Karaliyev T.M., Sayfiddinov I.F. … (qolgani o‘zgarmaydi).', color='C00000')
    dd.save(DST)
    print(pardoz(DST), DST)
