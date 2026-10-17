"""1–3-boblar jadval va rasmlarini tekshirish natijalari bo‘yicha tuzatishlar (3.3-§ dan tashqari).

Har bir almashtirish aniq bir joyda bajarilishi tekshiriladi (assert). Oxirida qora.py bilan butun matn qora qilinadi.
"""
import re, sys
sys.path.insert(0, '/home/user/tarjima/tools')
from docx.oxml.ns import qn
from lxml import etree
import kamlib as K
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
DONE = []


def body_paras():
    return [p for p in D.d.paragraphs]


def R(old, new, n=1, where=None):
    """Asosiy matn paragraflaridagi almashtirish (n – kutilgan joylar soni)."""
    ps = [p for p in D.paras() if old in p.text and (where is None or where in p.text)]
    cnt = sum(p.text.count(old) for p in ps)
    assert cnt == n, (old, cnt)
    for p in ps:
        for _ in range(p.text.count(old)):
            D.replace(p, old, new)
    DONE.append(old[:50])


def fn_by_text(sub):
    out = []
    for f in D.fx.findall(qn('w:footnote')):
        t = ''.join(x.text or '' for x in f.iter(qn('w:t'))).replace('\xa0', ' ')
        if sub in t:
            out.append(f)
    return out


def FN(sub, old, new, n=1):
    fs = fn_by_text(sub)
    assert len(fs) == n, (sub, len(fs))
    from docx.text.paragraph import Paragraph
    for f in fs:
        p = Paragraph(f.find(qn('w:p')), None)
        D.replace(p, old, new)
    DONE.append('fn:' + old[:40])


def FN_all(old, new):
    """Barcha izohlardagi almashtirish."""
    from docx.text.paragraph import Paragraph
    k = 0
    for f in D.fx.findall(qn('w:footnote')):
        for pe in f.findall(qn('w:p')):
            p = Paragraph(pe, None)
            for _ in range(''.join(t.text or '' for t in pe.iter(qn('w:t'))).count(old)):
                D.replace(p, old, new); k += 1
    return k


def FN_tail(sub, new_tail, pre='Tadqiqot natijasida muallif tomonidan '):
    """Izohning “Tadqiqot natijasida muallif tomonidan ” dan keyingi qismini almashtirish."""
    fs = fn_by_text(sub)
    assert len(fs) == 1, (sub, len(fs))
    from docx.text.paragraph import Paragraph
    p = Paragraph(fs[0].find(qn('w:p')), None)
    old = ''.join(t.text or '' for t in fs[0].find(qn('w:p')).iter(qn('w:t')))
    i = old.index(pre) + len(pre)
    D.replace(p, old[i:], new_tail)
    DONE.append('fn-tail:' + sub[:40])


def FN_tail_cap(cap_start, new_tail, pre='Tadqiqot natijasida muallif tomonidan '):
    """Sarlavhasi cap_start bilan boshlanadigan paragrafdagi izoh."""
    ps = [p for p in D.d.paragraphs if p.text.startswith(cap_start)]
    assert len(ps) == 1, (cap_start, len(ps))
    ids = [r.get(qn('w:id')) for r in ps[0]._p.iter(qn('w:footnoteReference'))]
    assert len(ids) == 1, (cap_start, ids)
    f = [f for f in D.fx.findall(qn('w:footnote')) if f.get(qn('w:id')) == ids[0]][0]
    from docx.text.paragraph import Paragraph
    pe = f.find(qn('w:p'))
    old = ''.join(t.text or '' for t in pe.iter(qn('w:t')))
    i = old.index(pre) + len(pre)
    D.replace(Paragraph(pe, None), old[i:], new_tail)
    DONE.append('fncap:' + cap_start)


def CELL(old, new, n=1):
    k = 0
    for t in D.d.tables:
        seen = set()
        for row in t.rows:
            for c in row.cells:
                if c._tc in seen:
                    continue
                seen.add(c._tc)
                for p in c.paragraphs:
                    for _ in range(p.text.count(old)):
                        D.replace(p, old, new); k += 1
    assert k == n, (old, k)
    DONE.append('cell:' + old[:40])


def TX(old, new, nmin=1):
    """Sxema qutichalari (w:txbxContent) ichidagi matn – Choice va Fallback nusxalarida."""
    k = 0
    for tb in D.d.element.body.iter(qn('w:txbxContent')):
        for pe in tb.iter(qn('w:p')):
            ts = [t for t in pe.iter(qn('w:t'))]
            s = ''.join(t.text or '' for t in ts)
            if old in s:
                s2 = s.replace(old, new)
                ts[0].text = s2
                ts[0].set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                for t in ts[1:]:
                    t.text = ''
                k += 1
    assert k >= nmin, (old, k)
    DONE.append('tx:' + old[:40])
    return k


def TXBOX(contains, newtext, nmin=1):
    """Qutichaning butun matnini almashtirish (ko‘p qatorli qutichalar uchun)."""
    k = 0
    for tb in D.d.element.body.iter(qn('w:txbxContent')):
        pp = tb.findall(qn('w:p'))
        full = ' '.join(''.join(t.text or '' for t in p.iter(qn('w:t'))) for p in pp)
        if contains in full:
            ts = list(pp[0].iter(qn('w:t')))
            ts[0].text = newtext
            for t in ts[1:]:
                t.text = ''
            for p in pp[1:]:
                tb.remove(p)
            k += 1
    assert k >= nmin, (contains, k)


def CH_STR(name, old, new):
    x, part = D.chart_xml(name)
    k = 0
    for v in list(x.iter('{http://schemas.openxmlformats.org/drawingml/2006/chart}v')) + \
            list(x.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}t')):
        if v.text == old:
            v.text = new; k += 1
    assert k >= 1, (name, old)
    part._blob = etree.tostring(x, xml_declaration=True, encoding='UTF-8', standalone=True)


# ============================ 1-BOB ============================
# 1.1.1-rasm
R('Demak, moliyaviy barqarorlik bu tushunchalarning yig‘indisi emas, balki ularning o‘zaro uyg‘unligi natijasida yuzaga keladigan sifat holatidir.',
  'Demak, moliyaviy barqarorlik bu tushunchalarning yig‘indisi emas, balki ularning o‘zaro uyg‘unligi natijasida yuzaga keladigan sifat holatidir. '
  'Rasmda, shuningdek, moliyaviy holat muayyan sanadagi statik baho sifatida, rivojlanish esa barqarorlikning dinamik mazmuni sifatida ko‘rsatilgan.')
FN('Лаврушина, И.Д. Мамоновой. – М.: КНОРУС, 2011. – С. 14–15, 57–61', 'С. 14–15, 57–61', 'С. 14–15, 49, 57–61')

# 1.1.1-jadval: I.V. Larionova
LAR = ('Ларионова И.В. Стабильность банковской системы в условиях переходной экономики: дис. … д-ра экон. наук: 08.00.10. – М.: '
       'Финансовая академия при Правительстве Российской Федерации, 2001. – С. 35–37.')
CELL('Ichki va tashqi ta’sirlarga qarshi turish, uzoq davr davomida barqaror muvozanat va ishonchlilikni saqlash qobiliyati',
     'Bank tizimi stabilligi – ichki va tashqi ta’sirlarga qarshi turish, vaqt davomida barqaror muvozanat va ishonchlilikni saqlash qobiliyati; barqarorlik – undan torroq, statik tushuncha')
CELL('Uzoq muddatli muvozanat', 'Stabillik: muvozanat va ishonchlilik')
R('I.V. Larionova moliyaviy barqarorlikni ichki va tashqi ta’sirlarga qarshi turish, uzoq vaqt davomida barqaror muvozanat va ishonchlilikni saqlash qobiliyati sifatida, V.V. Ivanov esa',
  'I.V. Larionova esa bank tizimi stabilligini ichki va tashqi ta’sirlarga qarshi turish, vaqt davomida barqaror muvozanat va ishonchlilikni saqlash qobiliyati sifatida ta’riflab, '
  'barqarorlikni undan torroq, muayyan sanadagi statik tushuncha deb hisoblaydi[[FN:' + LAR + ']]. V.V. Ivanov')
# 56-izoh: Gerasimova jadvalda yo‘q; Fetisov sahifasi; Larionova; nuqta
G = ('Герасимова Е.Б. Феноменологический подход к анализу устойчивости организации // Сибирская финансовая школа. – 2011. – № 3. – '
     'Цит. по: Вотинцева Р.С. Современные теоретические подходы к определению понятия “Финансовая устойчивость коммерческих банков” '
     '// Вестник Удмуртского университета. Серия “Экономика и право”. – 2014. – Вып. 3. – С. 46; ')
FN('Т. 27, № 1. – С. 65; Houben', G, '')
FN('Т. 27, № 1. – С. 65; Houben',
   'Финансы и статистика, 1999; ', 'Финансы и статистика, 1999. – С. 17; ' + LAR[:-len(' – С. 35–37.')] + ' – С. 35; ')
FN('Фетисов Г.Г. Устойчивость коммерческого банка и рейтинговые системы ее оценки. – М.: Финансы и статистика, 1999; Герасимова',
   'Финансы и статистика, 1999; ', 'Финансы и статистика, 1999. – С. 17; ')
k = 0
for old in (' – 18 p;', ' – 10 p;', ' – 32 p;'):
    k += FN_all(old, old[:-1] + '.;'); print('p-fix', old, k)

# 1.1.2-rasm
TX('A. Krokett', 'A. Crockett', 1)
R('A. Krokett ham', 'A. Crockett ham')
R('Ta’kidlash lozimki, aynan to‘rtinchi yondashuv Bazel qo‘mitasi',
  'Rasmning quyi qismida esa ushbu yondashuvlarning sintezi – daraja, dinamika va muvozanatni birlashtiruvchi integral yondashuv ko‘rsatilgan bo‘lib, '
  'u tadqiqotda ishlab chiqilgan baholash metodikasiga asos qilib olindi. Ta’kidlash lozimki, aynan to‘rtinchi yondashuv Bazel qo‘mitasi')

# 1.1.2-jadval
CELL('Likvidlik va fondlash alohida blok sifatida baholanadi', 'Likvidlik va moliyalashtirish alohida blok sifatida baholanadi', 1)
R('Ushbu model likvidlik va fondlash barqarorligini', 'Ushbu model likvidlik va moliyalashtirish barqarorligini')
CELL('Moliyaviy beqarorlik gipotezasi (Minskiy)', 'Moliyaviy beqarorlik gipotezasi (Minsky)', 1)
R('H.P. Minskiyning', 'H.P. Minskyning')
R('Minskiy ta’riflagan', 'Minsky ta’riflagan')
R('Minskiy gipotezasi dinamik tahlilni', 'Minsky gipotezasi dinamik tahlilni')
# 1.1.2 va 1.2.1-jadvallardagi bir xil katak: 1.1.2-jadvaldagisi o‘zgartiriladi
t112 = D.table_after('Moliyaviy barqarorlikni tushuntiruvchi konseptual modellarning qiyosiy tavsifi')
c = [cc for row in t112.rows for cc in row.cells if 'Kapital markaziy blok; baho minimumdan ortiq zaxiraga qarab beriladi' in cc.text][0]
D.replace(c.paragraphs[0], 'Kapital markaziy blok; baho minimumdan ortiq zaxiraga qarab beriladi',
          'Kapital zaxirasi regulyativ minimumga nisbatan baholanadi')
FN('P. 401–419; Minsky H.P. The Financial Instability Hypothesis',
   'P. 401–419; Minsky', 'P. 401–419; Diamond D.W., Rajan R.G. Liquidity Risk, Liquidity Creation, and Financial Fragility: A Theory '
   'of Banking // Journal of Political Economy. – 2001. – Vol. 109, No. 2. – P. 287–327; Minsky')
FN('P. 2–17; Borio C. Towards a Macroprudential Framework',
   'P. 2–17; Borio', 'P. 2–17; Roy A.D. Safety First and the Holding of Assets // Econometrica. – 1952. – Vol. 20, No. 3. – P. 431–449; Borio')

# 1.1.7-rasm: “Daromad-lilik”
k = 0
for tb in D.d.element.body.iter(qn('w:txbxContent')):
    pp = tb.findall(qn('w:p'))
    tx = [''.join(t.text or '' for t in p.iter(qn('w:t'))) for p in pp]
    for i in range(len(pp) - 1):
        if tx[i].strip() == 'Daromad-' and tx[i + 1].strip() == 'lilik':
            ts = list(pp[i].iter(qn('w:t'))); ts[0].text = 'Daromadlilik'
            for t in ts[1:]:
                t.text = ''
            tb.remove(pp[i + 1]); k += 1
            break
    else:
        for i, s in enumerate(tx):
            if 'Daromad-lilik' in s or 'Daromad- lilik' in s:
                ts = list(pp[i].iter(qn('w:t'))); ts[0].text = s.replace('Daromad- lilik', 'Daromadlilik').replace('Daromad-lilik', 'Daromadlilik')
                for t in ts[1:]:
                    t.text = ''
                k += 1
assert k >= 1, 'Daromad-lilik'

# 1.2.1-jadval
CELL('Bitta bank inqirozi butun tizimga tarqaladi', 'Bitta bank inqirozi butun tizimga tarqalishi mumkin', 1)

# 1.2.2-jadval
t122 = D.table_after('Moliyaviy barqarorlikni baholash yondashuvlarining qiyosiy tavsifi')
cells = [cc for row in t122.rows for cc in row.cells]
c = [cc for cc in cells if cc.text.strip() == 'RATE'][0]
D.replace(c.paragraphs[0], 'RATE', 'RATE (Buyuk Britaniya), SREP (Yevropa Ittifoqi)')
c = [cc for cc in cells if cc.text.startswith('FIMS, logit')][0]
D.replace(c.paragraphs[0], 'FIMS,', 'SEER (avval FIMS),')
FN('B. 195–213; Chiaramonte L., Poli F. Predicting European Bank Distress',
   'Central Bank Policy. – Cambridge: Cambridge University Press, 2017. – P. 194.',
   'Central Bank Policy. – Cambridge: Cambridge University Press, 2017. – P. 194; '
   'Указание Банка России от 03.04.2017 № 4336-У “Об оценке экономического положения банков”. – https://www.cbr.ru.')
# “sub’ekt” → “subyekt” (amaldagi imlo)
for p in D.paras():
    for o, nw in (('Sub’ekt', 'Subyekt'), ('sub’ekt', 'subyekt')):
        while o in p.text:
            D.replace(p, o, nw)

# 1.2.1-rasm (Bazel bosqichlari)
TX('Tier 1 ≥ 4%', 'birinchi darajali kapital ≥ 4%', 1)
k = 0
for tb in D.d.element.body.iter(qn('w:txbxContent')):
    full = ''.join(t.text or '' for t in tb.iter(qn('w:t')))
    if 'Bazel III' in full and '(2010)' in full and 'CET1' not in full:
        for t in tb.iter(qn('w:t')):
            if t.text and '(2010)' in t.text:
                t.text = t.text.replace('(2010)', '(2010–2017)'); k += 1
assert k >= 1, 'Bazel III (2010)'
TXBOX('LCR, NSFR',
   'CET1 ≥ 4,5%; birinchi darajali kapital ≥ 6%; konservatsiya va kontrsiklik buferlar; leverej ≥ 3%; LCR, NSFR ≥ 100%; quyi chegara 72,5% (2017)', 1)
FN('P. 45–67; Tian W. Regulatory Capital Requirement in Basel III',
   'НИУ ВШЭ, 2015. – С. 20–24.',
   'НИУ ВШЭ, 2015. – С. 20–24; Basel Committee on Banking '
   'Supervision. Basel III: A global regulatory framework for more resilient banks and banking systems. – Basel: Bank for International '
   'Settlements, December 2010 (rev. June 2011); Basel Committee on Banking Supervision. Basel III: Finalising post-crisis reforms. – '
   'Basel: BIS, December 2017.')

# 1.2.2-rasm
R('asosiy kapital talabi 2 foizdan 4,5 foizga', 'birinchi darajali asosiy kapital talabi 2 foizdan 4,5 foizga')
R('2,5 foizlik kapitalni saqlash buferi va 0–2,5 foizlik', '2,5 foizlik kapitalni saqlash (konservatsiya) buferi va 0–2,5 foizlik')
CH_STR('/word/charts/chart3.xml', 'Asosiy kapital (CET1)', 'Birinchi darajali asosiy kapital (CET1)')

# 1.2.4-rasm va matn
R('MHXS 9 kredit riskini muammo yuzaga kelgandan keyin emas, balki kredit riski sezilarli oshgan paytdan boshlab, ya’ni 2-bosqichdan e’tiborga oladi.',
  'MHXS 9 kutilgan kredit zararlarini muammo yuzaga kelgandan keyin emas, balki aktiv tan olingan kundan boshlab (1-bosqichda 12 oylik) aks ettiradi, '
  'kredit riski sezilarli oshgan paytdan, ya’ni 2-bosqichdan boshlab esa butun muddat uchun zaxira talab qiladi.')
R('N.E. Sokolinskaya prognoziga ko‘ra, MHXS 9 ga o‘tish natijasida Rossiya banklarida zaxiralar kredit portfelining 19,1 foiziga yetishi mumkin bo‘lgan.',
  'N.E. Sokolinskaya hisob-kitobiga ko‘ra, MHXS 9 ga o‘tishda garovdan olingan aktivlar defolt deb hisoblansa, bitta Rossiya banki misolida '
  'zaxiralar kredit portfelining 19,1 foiziga yetib, deyarli ikki barobar oshadi.')
TXBOX('Kredit riski sezilarli oshgan aktivlar', 'Kredit riski sezilarli oshgan aktivlar (30 kundan ortiq kechikish)', 1)
TXBOX('Qadrsizlangan (defolt) aktivlar', 'Kredit qadrsizlangan aktivlar (90 kundan ortiq kechikish – defolt)', 1)
FN_all('IFRS 9 Financial Instruments. – London: International Accounting Standards Board, 2014; Соколинская',
       'IFRS 9 Financial Instruments. – London: International Accounting Standards Board, 2014. – Para. 5.5.3, 5.5.5, 5.5.11, B5.5.37; Соколинская')
for p in D.paras():
    while 'kutilayotgan kredit zarar' in p.text:
        D.replace(p, 'kutilayotgan kredit zarar', 'kutilgan kredit zarar')

# 1.2.4-jadval
t124 = D.table_after('Banklar moliyaviy barqarorligini baholash bo‘yicha xalqaro standartlar va tizimlarning qiyosiy ta')
cells = [cc for row in t124.rows for cc in row.cells]
c = [cc for cc in cells if cc.text.strip() == 'Kapital, likvidlik'][0]
D.replace(c.paragraphs[0], 'Kapital, likvidlik', 'Kapital, leverej, likvidlik')
c = [cc for cc in cells if cc.text.strip() == 'Faqat kredit riski'][0]
D.replace(c.paragraphs[0], 'Faqat kredit riski', 'Faqat kredit zararlari (qadrsizlanish); o‘zi baholash tizimi emas')
cap124 = [p for p in D.d.paragraphs if p.text.startswith('Banklar moliyaviy barqarorligini baholash bo‘yicha xalqaro standartlar va tizimlarning qiyosiy ta')][0]
fid = [r.get(qn('w:id')) for r in cap124._p.iter(qn('w:footnoteReference'))][0]
from docx.text.paragraph import Paragraph
fp = Paragraph([f for f in D.fx.findall(qn('w:footnote')) if f.get(qn('w:id')) == fid][0].find(qn('w:p')), None)
_ft = ''.join(t.text or '' for t in fp._p.iter(qn('w:t'))).strip()
assert _ft == 'Tadqiqot natijasida muallif tomonidan ishlab chiqildi.', _ft
D.replace(fp, 'ishlab chiqildi.',
          'quyidagi manbalar asosida ishlab chiqildi: Basel Committee on Banking Supervision. Basel III: A global regulatory framework for more '
          'resilient banks and banking systems. – Basel: Bank for International Settlements, December 2010 (rev. June 2011); IFRS 9 Financial '
          'Instruments. – London: International Accounting Standards Board, 2014; International Monetary Fund. Financial Soundness Indicators: '
          'Compilation Guide. – Washington, DC: IMF, 2019; Langohr H., Langohr P. The Rating Agencies and Their Credit Ratings: What They Are, '
          'How They Work and Why They Are Relevant. – Chichester: John Wiley & Sons, 2008.')
FN_all('International Monetary Fund. 2019 Financial Soundness Indicators Compilation Guide.',
       'International Monetary Fund. Financial Soundness Indicators: Compilation Guide.')

# ============================ 2.2-§ ============================
FN_tail_cap('2.2.3-rasm.', '2.2.1-jadval ma’lumotlari asosida ishlab chiqildi.')
R('Soliq to‘lagunga qadar aktivlar rentabelligi 2021-yilda 1,3 foiz, kapital rentabelligi 6,1 foiz bo‘lgan bo‘lsa',
  'Soliq to‘lagunga qadar foyda bo‘yicha aktivlar rentabelligi 2021-yilda 1,3 foiz, sof foyda bo‘yicha kapital rentabelligi esa 6,1 foiz bo‘lgan bo‘lsa')
R('nofoiz xarajatlarning', 'foizsiz xarajatlarning')
CH_STR('/word/charts/chart13.xml', 'ROE', 'Kapital rentabelligi (ROE)')
CH_STR('/word/charts/chart13.xml', 'ROA', 'Aktivlar rentabelligi, soliq to‘lagunga qadar (ROA)')
R('3-bosqich kreditlari ulushi 7,48 foizni', '3-bosqich kreditlari ulushi 7,5 foizni')
CELL('Asosiy kapital (CET1) monandligi, %', 'I darajali asosiy kapital (CET1) monandligi, %', 3)
CELL('Asosiy kapital monandligi, %', 'I darajali asosiy kapital (CET1) monandligi, %', 1)
R('Kapital monandligi yuqoriroq bo‘lsa-da, keskin tebranadi',
  'Kapital monandligi “O‘zsanoatqurilishbank” ATB nikiga nisbatan aksariyat yillarda yuqoriroq bo‘lsa-da, keskin tebranadi')

# ============================ 2.3-§ ============================
TX('Tafovut: 2–3 barobar', 'Tafovut: 1,7–4,9 barobar', 1)
TX('Tafovut: 10–20 foiz', 'Tafovut: 13–18 foiz', 1)
TX('Tafovut: 2–3 foiz punkt', 'Tafovut: 1–3 foiz punkt', 1)
SQB_GEN = ('“O‘zsanoatqurilishbank” ATB ning 2021–2025-yillardagi prudensial hisobot shakllari va MHXS bo‘yicha konsolidatsiyalangan '
           'moliyaviy hisobotlari asosida tayyorlandi.')
FN_tail_cap('2.3.2-rasm.', SQB_GEN)
FN_tail_cap('2.3.3-rasm.', SQB_GEN)
MKB_GEN = ('“Mikrokreditbank” ATB ning 2021–2025-yillardagi prudensial hisobot shakllari va MHXS bo‘yicha moliyaviy hisobotlari asosida '
           'tayyorlandi.')
FN_tail_cap('“Mikrokreditbank” ATB ning asosiy ko‘rsatkichlari: milliy hisobot va MHXS', MKB_GEN)
FN_tail_cap('2.3.6-rasm.', MKB_GEN)
x, part = D.chart_xml('/word/charts/chart22.xml')
s0 = x.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}ser')[0]
k = 0
for e in s0.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr'):
    if e.get('val') == '70AD47':
        e.set('val', '2A78D6'); k += 1
assert k >= 1
part._blob = etree.tostring(x, xml_declaration=True, encoding='UTF-8', standalone=True)
# Mikrokreditbank yashirin kredit riski koeffitsiyenti (yaxlitlanmagan ulushlardan)
t233 = D.table_after('“Mikrokreditbank” ATB ning asosiy ko‘rsatkichlari: milliy hisobot va MHXS')
row = [r for r in t233.rows if r.cells[0].text.startswith('Yashirin kredit riski koeffitsiyenti')][0]
for c in row.cells[1:]:
    for o, nw in (('2,95', '2,94'), ('2,81', '2,83'), ('3,30', '3,28')):
        if c.text.strip() == o:
            D.replace(c.paragraphs[0], o, nw)
assert [c.text.strip() for c in row.cells[1:]] == ['2,94', '2,83', '2,77', '3,28', '2,80'], [c.text for c in row.cells]
D.chart_set('/word/charts/chart21.xml', 1, [2.94, 2.83, 2.77, 3.28, 2.8])
K.chart_sync_xlsx(D, '/word/charts/chart21.xml')
R('2,77–3,30', '2,77–3,28', 3)
CELL('yashirin kredit riski koeffitsiyenti 1,65–3,30', 'yashirin kredit riski koeffitsiyenti 1,65–4,90', 1)
# 2.3.4-jadval: izoh va manba
t234 = D.table_after('ko‘rsatkichlarining milliy va xalqaro yondashuv bo‘yicha tafovuti')
izsrc = D.para('Izoh: milliy kredit zaxiralari – kreditlar bo‘yicha maxsus zaxiralar')
iz = D.new_para_after(izsrc, 'Izoh: kapital monandligi ko‘rsatkichlari Markaziy bank metodikasi bo‘yicha MHXS ma’lumotlari asosida qayta '
                             'hisoblandi; NSFR va LCR – Bazel III metodikasi bo‘yicha.')
t234._tbl.addnext(iz._p)
FN_tail_cap('“O‘zsanoatqurilishbank” ATB moliyaviy barqarorlik ko‘rsatkichlarining milliy va xalqaro',
        '“O‘zsanoatqurilishbank” ATB ning 2025-yil uchun prudensial hisobot shakllari, MHXS bo‘yicha konsolidatsiyalangan moliyaviy hisoboti '
        'va Bazel III standartlari asosida tayyorlandi.')
# 2.3.5-jadval
R('“Aloqabank” ATB ko‘rsatkichlarining milliy va MHXS hisobotlari bo‘yicha qiyosi, 2021–2025-yillar',
  '“Aloqabank” ATB ning asosiy ko‘rsatkichlari: milliy hisobot va MHXS, 2021–2025-yillar')
R('Bunga MHXS bo‘yicha kutilgan kredit zararlari zaxirasining 2024–2025-yillarda 156,2 mlrd so‘mdan 513,1 mlrd so‘mgacha oshishi sabab bo‘lgan.',
  'Bunga, birinchidan, bozor stavkasidan past stavkadagi joylashtirishlarni dastlabki tan olishdan ko‘rilgan 283,7 mlrd so‘m zarar '
  '(shu jumladan “O‘zbektelekom” AJ ga joylashtirilgan mablag‘lar bo‘yicha 260,5 mlrd so‘m), ikkinchidan, kutilgan kredit zararlari '
  'zaxirasining 2023-yil yakunidagi 156,2 mlrd so‘mdan 2025-yil yakunida 513,1 mlrd so‘mgacha oshishi sabab bo‘lgan.')
FN('“Aloqabank” ATB ning prudensial hisobot shakllari va MHXS bo‘yicha konsolidatsiyalangan moliyaviy hisobotlari asosida hisoblandi.',
   '“Aloqabank” ATB ning prudensial', '“Aloqabank” ATB ning 2021–2025-yillardagi prudensial', n=2)

# ============================ 3.1.6-jadval ============================
R('(2.3.4, 2.3.5-jadvallar va “Mikrokreditbank” ATB bo‘yicha jadval)', '(2.3.3–2.3.5-jadvallar; “Mikrokreditbank” ATB ning regulyativ kapital monandligi – 19,96 foiz – xuddi shu usulda hisoblandi)')
R('hamda MHXS va Bazel III metodikasi asosida (3.1.6-jadval)', 'hamda MHXS hisobotlari asosida (3.1.6-jadval)')
FN_tail_cap('Uch bank bo‘yicha integral indeks ko‘rsatkichlari va ballari: MHXS',
        '“O‘zsanoatqurilishbank”, “Mikrokreditbank” va “Aloqabank” ATBlarining 2025-yil uchun MHXS bo‘yicha moliyaviy hisobotlari va '
        'prudensial hisobot shakllari asosida hisoblandi.')

# ============================ Adabiyotlar ============================


def bib_paras():
    on = False; out = []
    for p in D.d.paragraphs:
        if p.text.strip() == 'FOYDALANILGAN ADABIYOTLAR RO‘YXATI':
            on = True; out = []; continue
        if on and p.text.strip() == 'ILOVALAR':
            break
        if on:
            out.append(p)
    return out


NEWB = [('Diamond D.W., Dybvig P.H. Bank Runs, Deposit Insurance, and Liquidity',
         'Diamond D.W., Rajan R.G. Liquidity Risk, Liquidity Creation, and Financial Fragility: A Theory of Banking // Journal of Political '
         'Economy. – 2001. – Vol. 109, No. 2. – P. 287–327.'),
        ('Карминский А.М. Кредитные рейтинги и их моделирование.',
         'Ларионова И.В. Стабильность банковской системы в условиях переходной экономики: дис. … д-ра экон. наук: 08.00.10. – М.: '
         'Финансовая академия при Правительстве Российской Федерации, 2001. – 415 с.')]
for anc, txt in NEWB:
    a = [p for p in bib_paras() if anc in p.text]
    assert len(a) == 1, anc
    D.new_para_after(a[0], '0. ' + txt)
n = 0
for p in bib_paras():
    m = re.match(r'^(\d+)\. ', p.text)
    if not m:
        continue
    n += 1
    r0 = D._runs(p)[0].find(qn('w:t'))
    if r0.text.startswith(m.group(1) + '.'):
        r0.text = str(n) + r0.text[len(m.group(1)):]
    elif r0.text == m.group(1):
        r0.text = str(n)
    else:
        raise RuntimeError(p.text[:60])
R('foydalanilgan adabiyotlar ro‘yxati 208 ta manbani', 'foydalanilgan adabiyotlar ro‘yxati %d ta manbani' % n)

D.save(F)
print(len(DONE), 'ta tuzatish; manbalar:', n, pardoz(F))
