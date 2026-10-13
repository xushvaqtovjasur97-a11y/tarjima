"""1.3.4-rasmning yangi varianti (namuna): Buyuk Britaniya PRA faol aralashuv tizimi (PIF) bosqichlari.
Manba: Bank of England, PRA. The PRA’s approach to banking supervision (July 2023), Box 5, P. 43–45; Table 1, P. 18."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
import docx
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from lxml import etree
import wordart as W
from pardoz import pardoz

DST = '/home/user/tarjima/Namuna_1.3.4-rasm_yangi.docx'
PRA = 'Bank of England, Prudential Regulation Authority. The Prudential Regulation Authority’s approach to banking supervision. – London: Bank of England, July 2023.'
BCBS = 'Sahajwala R., Van den Bergh P. Supervisory risk assessment and early warning systems: Basel Committee on Banking Supervision Working Papers No. 4. – Basel: BIS, December 2000.'

STAGES = [
    ('1-bosqich\nRisk past', '70AD47', 'Bank yashovchanligiga xavf past', 'Odatiy nazorat: risklarni baholash, tiklash va tugatish rejalari'),
    ('2-bosqich\nO‘rtacha risk', '9BC2E6', 'Moliyaviy holat, risk menejmenti yoki boshqaruvda zaifliklar aniqlangan', 'Nazorat kuchaytiriladi, qo‘shimcha hisobot talab qilinadi, kamchiliklar belgilangan muddatda bartaraf etiladi'),
    ('3-bosqich\nChora ko‘rilmasa xavf', 'FFC000', 'Bank barqarorligiga jiddiy tahdidlar aniqlangan', 'Rahbariyat yoki kengash tarkibini o‘zgartirish; dividendlar, o‘sish va faoliyatni cheklash; kapital va likvidlik talablarini oshirish'),
    ('4-bosqich\nXavf yaqin', 'ED7D31', 'Bank chegaraviy shartlarni bajara olmasligi xavfi real, biroq tuzatish imkoniyati saqlangan', 'Tiklash choralari qisqa muddatda bajariladi, aktivlar yoki bankning o‘zi sotiladi'),
    ('5-bosqich\nTugatish', 'C00000', 'Bank tugatish yoki faoliyatini to‘xtatish jarayonida', 'Angliya bankining tugatish vakolatlari qo‘llanadi, omonatlar qoplanadi'),
]


def figure():
    sh = []
    y = 0.0
    sh.append(W.box(0.0, y, 3.6, 0.6, 'PIF bosqichi', fill=None, line=None, sz=20, bold=True))
    sh.append(W.box(3.8, y, 4.9, 0.6, 'Bankning holati', fill=None, line=None, sz=20, bold=True))
    sh.append(W.box(8.9, y, 6.6, 0.6, 'Nazorat choralari', fill=None, line=None, sz=20, bold=True))
    y = 0.7
    for lab, col, st, act in STAGES:
        fc = 'FFFFFF' if col in ('C00000', 'ED7D31', '70AD47') else '000000'
        sh.append(W.box(0.0, y, 3.6, 1.45, lab, fill=col, line=col, sz=20, bold=True, color=fc))
        sh.append(W.box(3.8, y, 4.9, 1.45, st, fill='FFFFFF', line='7F7F7F', sz=19))
        sh.append(W.box(8.9, y, 6.6, 1.45, act, fill='FFFFFF', line=col, sz=19, italic=True))
        y += 1.6
    sh.append(W.line(15.85, 0.75, 15.85, y - 0.2, color='C00000', lw=28575, arrow=True))
    return W.group_run_xml(sh, 16.2, y, '1.3.4-rasm'), y


def texts(fn):
    s_uk = ('Angliya banki tomonidan ishlab chiqilgan va 1998-yildan Moliyaviy xizmatlar boshqarmasi tomonidan qo‘llangan RATE tizimi uch blokdan, ya’ni risklarni baholash, nazorat vositalari va samaradorlikni baholashdan iborat bo‘lgan'
            + fn(BCBS)
            + '. Risklarni baholashda CAMEL-B omillari (kapital, aktivlar, bozor riski, foyda, majburiyatlar, biznes) va COM omillari (nazorat, tashkiliy tuzilma, boshqaruv) asosida yagona ball chiqarilib, nazorat davri 6 oydan 3 yilgacha belgilangan. 2013-yildan banklar nazorati Angliya banki tarkibidagi Prudensial tartibga solish boshqarmasiga o‘tgan bo‘lib, u nazorat intensivligini ikki o‘lchov asosida belgilaydi: bankning moliya tizimiga potensial ta’siri bo‘yicha to‘rt toifa va bankning faoliyati to‘xtatilishiga yaqinligi bo‘yicha faol aralashuv tizimining besh bosqichi'
            + fn(PRA[:-1] + ' – P. 18, 43.') + ' (1.3.4-rasm).')
    caption = 'Angliya banki Prudensial tartibga solish boshqarmasining faol aralashuv tizimi bosqichlari'
    cap_fn = 'Tadqiqot natijasida muallif tomonidan quyidagi manba asosida ishlab chiqildi: ' + PRA[:-1] + ' – P. 43–45 (Box 5).'
    a = ('Rasmdan ko‘rinadiki, bank yuqori bosqichga o‘tgan sari nazorat organining choralari kuchayib boradi va 3-bosqichdan boshlab bank boshqaruvi, dividendlari va faoliyatiga bevosita aralashuvga aylanadi. Muhim jihati shundaki, nazorat organi bankning qaysi bosqichda ekanini bozorga ham, bankning o‘ziga ham muntazam ravishda oshkor qilmaydi, chunki bu stress sharoitida beqarorlikni kuchaytirishi mumkin'
         + fn(PRA[:-1] + ' – P. 43.')
         + '. Bu yondashuvning afzalligi cheklangan nazorat resurslarini eng riskli va moliya tizimi uchun eng ahamiyatli banklarga yo‘naltirishdadir. Boshqacha aytganda, baholash natijasi shunchaki reyting bo‘lib qolmaydi, balki aniq nazorat qaroriga aylanadi.')
    return dict(s_uk=s_uk, caption=caption, cap_fn=cap_fn, a=a)


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
        return q

    FN = []

    def fn(t):
        FN.append(t); return f'[{len(FN)}]'

    T = texts(fn)
    par('1.3.4-rasmning yangi varianti (namuna)', bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False)
    par('Ma’qul bo‘lsa: 1.3-§ dagi “Angliya banki 1997-yildan BAITE tizimini qo‘llab kelmoqda. … (1.3.4-rasm).” qismi quyidagi matn bilan, 1.3.4-rasm va undan keyingi “Matritsadan ko‘rinadiki, …” xatboshi esa quyidagi rasm va tahlil bilan almashtiriladi. Xatboshining boshi (“AQShdan farqli ravishda, … alohida o‘rin tutadi.”) va Lavrushin izohlari o‘zgarishsiz qoladi.',
        italic=True, size=12, indent=False, color='595959')
    par('… ' + T['s_uk'])
    q = dd.add_paragraph(); q.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run, h = figure()
    q._p.append(etree.fromstring(run))
    FN.append(T['cap_fn']); par('1.3.4-rasm. ' + T['caption'] + f'[{len(FN)}]', bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False)
    par(T['a'])
    par('Sahifa osti snoskalari', bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False)
    for i, t in enumerate(FN, 1):
        par(f'[{i}] {t}', size=11, indent=False)
    dd.save(DST)
    print(pardoz(DST), DST)
