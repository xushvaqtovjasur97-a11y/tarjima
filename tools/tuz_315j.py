"""3.1.5-jadval → ikki jadval (milliy / MHXS, uch bank) + tahlil; 3.1.3-rasm olib tashlandi; sezgirlik rasmi uch bankka;
bog‘liq xulosalar (83–86 → 56–86). Matn o‘zgarishlari qizil, diagramma odatiy ranglarda."""
import copy, io, re, sys
sys.path.insert(0, '/home/user/tarjima/tools')
import openpyxl
import docx
from docx.oxml.ns import qn
import kamlib as K
import namuna_315j as N
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
P = lambda s: D.para(s)
T = 'tab315'

# A) eski 3.1.6-jadval → 3.1.7-jadval
old316 = [p for p in D.d.paragraphs if p.text.strip() == '3.1.6-jadval'][0]
D.replace(old316, '3.1.6-jadval', '3.1.7-jadval', tag=T)
D.replace(P('keng qo‘llanadigan metodikalar bilan solishtirildi (3.1.6-jadval)'), '3.1.6-jadval', '3.1.7-jadval', tag=T)

# B) kirish gapi
intro = P('Taklif etilayotgan metodikaning amaliy qo‘llanilishini')
D.replace(intro, N.INTRO_OLD[0], N.INTRO_OLD[1], tag=T)
D.replace(intro, N.INTRO_OLD2[0], N.INTRO_OLD2[1], tag=T)

# C) 1-jadval nomi va manbasi
no1 = [p for p in D.d.paragraphs if p.text.strip() == '3.1.5-jadval'][0]
cap1 = no1._p.getnext(); cap1 = docx.text.paragraph.Paragraph(cap1, no1._parent)
fid = [r.get(qn('w:id')) for r in cap1._p.iter(qn('w:footnoteReference'))][0]
D.replace(cap1, cap1.text, N.CAP1, tag=T)
fnp = D.fn_para(fid=fid)
fnt = ''.join(x.text or '' for x in fnp._p.iter(qn('w:t'))).strip()
D.replace(fnp, fnt, N.FN1, tag=T)

# D) 1-jadval: 5 ustun → 4, 15 qator → 21
tbl = cap1._p.getnext()
t = docx.table.Table(tbl, cap1._parent)
for tr in tbl.findall(qn('w:tr')):
    tcs = tr.findall(qn('w:tc')); tr.remove(tcs[-1])
grid = tbl.find(qn('w:tblGrid')); gcs = grid.findall(qn('w:gridCol')); grid.remove(gcs[-1])
WD = [4082, 1758, 1758, 1758]
for g, w in zip(grid.findall(qn('w:gridCol')), WD): g.set(qn('w:w'), str(w))
trs = tbl.findall(qn('w:tr'))
mb = trs[-2]
for _ in range(5):
    mb.addprevious(copy.deepcopy(mb))
for tr in tbl.findall(qn('w:tr')):
    for tc, w in zip(tr.findall(qn('w:tc')), WD):
        tcw = tc.find(qn('w:tcPr') + '/' + qn('w:tcW'))
        if tcw is not None: tcw.set(qn('w:w'), str(w)); tcw.set(qn('w:type'), 'dxa')
t = docx.table.Table(tbl, cap1._parent)


def fill(tt, var):
    for i, row in enumerate(N.jadval_rows(var)):
        for j, v in enumerate(row):
            D.set_cell(tt, i, j, v, tag=T)
fill(t, 'milliy')

# E) izoh, F) tahlil
iz1 = docx.text.paragraph.Paragraph(tbl.getnext(), cap1._parent)
D.replace(iz1, iz1.text, N.IZOH1, tag=T)
an1 = docx.text.paragraph.Paragraph(iz1._p.getnext(), cap1._parent)
assert an1.text.startswith('Jadval ma’lumotlari metodikaning asosiy afzalligini')
T1 = N.T1 + (' Integral indeks bunday turli profillarni yagona shkalada solishtirish imkonini beradi, bloklar ballari esa har bir '
             'bank uchun aniq nazorat yo‘nalishini ko‘rsatadi.')
D.replace(an1, an1.text, T1, tag=T)

# G) 2-jadval (MHXS)
no2 = D.new_para_after(no1, '3.1.6-jadval', tag=T); no2 = no1._p.getnext()
cap2 = D.new_para_after(cap1, N.CAP2 + '[[FN:' + N.FN2 + ']]', tag=T); cap2 = cap1._p.getnext()
iz2 = D.new_para_after(iz1, N.IZOH2, tag=T); iz2 = iz1._p.getnext()
an2 = D.new_para_after(an1, N.T2, tag=T); an2 = an1._p.getnext()
tbl2 = copy.deepcopy(tbl)
an1._p.addnext(no2); no2.addnext(cap2); cap2.addnext(tbl2); tbl2.addnext(iz2); iz2.addnext(an2)
fill(docx.table.Table(tbl2, cap1._parent), 'MHXS')

# H) 3.1.3-rasm (bloklar ballari) olib tashlanadi: diagramma, nomi, ikki tahlil xatboshisi
c313 = P('3.1.3-rasm. “O‘zsanoatqurilishbank” ATB va “Mikrokreditbank” ATB da integral indeks bloklari')
fig = docx.text.paragraph.Paragraph(c313._p.getprevious(), c313._parent)
a1 = docx.text.paragraph.Paragraph(c313._p.getnext(), c313._parent)
a2 = docx.text.paragraph.Paragraph(a1._p.getnext(), c313._parent)
assert a1.text.startswith('Rasmdan ko‘rinib turibdiki, ikkala bankda ham indeks pasayishining') and a2.text.startswith('Shu bilan birga, bloklar ballari ikki bankning')
for p in (fig, c313, a1, a2):
    D.delete_para(p, tag=T)

# I) 3.1.4-rasm → 3.1.3-rasm
c314 = P('3.1.4-rasm. Moliyaviy barqarorlikning integral indeksi qiymatlarining vaznlar tizimiga sezgirligi')
D.replace(c314, '3.1.4-rasm', '3.1.3-rasm', tag=T)
D.replace(P('va O‘zbekistonda CAMEL tizimi uchun keltirilgan nisbat (3.1.4-rasm)'), '3.1.4-rasm', '3.1.3-rasm', tag=T)

# J) sezgirlik tahlili
s = P('Rasm ma’lumotlari muhim xulosaga olib keladi. MHXS asosidagi indeks bo‘yicha banklarning tartibi')
OLD_S = s.text[:s.text.index('Demak, MHXS asosidagi indeks vaznlar tanloviga')]
NEW_S = ('Rasm ma’lumotlari muhim xulosaga olib keladi. MHXS asosidagi indeks bo‘yicha banklarning tartibi uchala vaznlar tizimida '
         'ham o‘zgarmaydi: “O‘zsanoatqurilishbank” ATB birinchi, “Aloqabank” ATB ikkinchi, “Mikrokreditbank” ATB esa uchinchi o‘rinda '
         'bo‘lib, u taklif etilgan va CAMEL vaznlarida III sinfda, teng vaznlarda II sinfning quyi chegarasida (2,58) bo‘ladi. '
         '“Aloqabank” ATB ning indeksi CAMEL vaznlarida eng ko‘p pasayadi (2,56), chunki bu tizimda daromadlilik blokining vazni '
         'yuqori (0,25), bank esa MHXS bo‘yicha zarar ko‘rgan. Milliy hisobot asosidagi indeksda esa banklar tartibi vaznlarga '
         'bog‘liq: taklif etilgan va CAMEL vaznlarida “O‘zsanoatqurilishbank” ATB birinchi o‘rinda bo‘lsa, teng vaznlarda u '
         '“Aloqabank” (3,11) va “Mikrokreditbank” (3,07) ATBlaridan keyin oxirgi o‘ringa (3,06) tushadi. ')
D.replace(s, OLD_S, NEW_S, tag=T)

# K) sezgirlik diagrammasi: 6 toifa
sys.path.insert(0, '/home/user/tarjima/manbalar/hisob_kitob')
import mbii as M
O = N.O
CATS = [('SQB: milliy', 'SQB milliy'), ('SQB: MHXS', 'SQB MHXS'), ('MKB: milliy', 'MKB milliy'), ('MKB: MHXS', 'MKB MHXS'),
        ('Aloqabank: milliy', 'ALQ milliy'), ('Aloqabank: MHXS', 'ALQ MHXS')]
VAL = {w: [round(M.mbii(O[k]['ind'], M.W[w])[0], 2) for _, k in CATS] for w in ('taklif', 'teng', 'CAMEL')}
fig = docx.text.paragraph.Paragraph(c314._p.getprevious(), c314._parent)
rid = [e.get(qn('r:id')) for e in fig._p.iter('{http://schemas.openxmlformats.org/drawingml/2006/chart}chart')][0]
part = D.d.part.related_parts[rid]
x = part.blob.decode('utf-8')
sers = re.findall(r'<c:ser>.*?</c:ser>', x, re.S)
cat = ('<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$7</c:f><c:strCache><c:ptCount val="6"/>' +
       ''.join(f'<c:pt idx="{i}"><c:v>{c}</c:v></c:pt>' for i, (c, _) in enumerate(CATS)) + '</c:strCache></c:strRef></c:cat>')
for sr, (w, col) in zip(sers, [('taklif', 'B'), ('teng', 'C'), ('CAMEL', 'D')]):
    val = (f'<c:val><c:numRef><c:f>Sheet1!${col}$2:${col}$7</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="6"/>' +
           ''.join(f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i, v in enumerate(VAL[w])) + '</c:numCache></c:numRef></c:val>')
    n = re.sub(r'<c:cat>.*?</c:cat>', cat, sr, flags=re.S)
    n = re.sub(r'<c:val>.*?</c:val>', val, n, flags=re.S)
    x = x.replace(sr, n, 1)
part._blob = x.encode('utf-8')
for r in part.rels.values():
    if r.reltype.endswith('/package'):
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'Sheet1'
        ws.append(['', 'Taklif etilgan vaznlar', 'Teng vaznlar', 'CAMEL nisbati'])
        for i, (c, _) in enumerate(CATS):
            ws.append([c, VAL['taklif'][i], VAL['teng'][i], VAL['CAMEL'][i]])
        b = io.BytesIO(); wb.save(b); r.target_part._blob = b.getvalue()
print('sezgirlik:', VAL)

# L) bog‘liq xulosalar
p = P('Metodikani ikki bankka qo‘llash MHXS ma’lumotlari asosida')
D.replace(p, 'Metodikani ikki bankka qo‘llash MHXS ma’lumotlari asosida ikkala bank ham milliy hisobotdagiga nisbatan bir sinf pastga tushishini, ΔMBII ning 83–86 foizi',
          'Metodikani uch bankka qo‘llash MHXS ma’lumotlari asosida uchala bankning indeksi milliy hisobotdagiga nisbatan 0,46–0,66 ballga pasayishini, ulardan ikkitasi bir sinf pastga tushishini, ΔMBII ning 56–86 foizi', tag=T)
p = P('ΔMBII −0,46 va −0,66 ballga teng bo‘lib, uning 83–86 foizi')
D.insert(p, '3,13 va 2,47 ballni (II va III sinflar)', ', “Aloqabank” ATB niki 3,248 va 2,78 ballni (ikkala variantda II sinf)', tag=T)
D.replace(p, 'ΔMBII −0,46 va −0,66 ballga teng bo‘lib, uning 83–86 foizi', 'ΔMBII mos ravishda −0,46; −0,66 va −0,47 ballga teng bo‘lib, uning 56–86 foizi', tag=T)
p = P('“Mikrokreditbank” ATB da −0,66 ballni tashkil etib, uning 83–86 foizi')
D.replace(p, '“Mikrokreditbank” ATB da −0,66 ballni tashkil etib, uning 83–86 foizi', '“Mikrokreditbank” ATB da −0,66, “Aloqabank” ATB da −0,47 ballni tashkil etib, uning 56–86 foizi', tag=T)
D.replace(p, 'bir sinfga oshirib', 'bir sinfgacha oshirib', tag=T)
p = P('13 ko‘rsatkichli to‘liq indeks bo‘yicha milliy hisobot asosidagi qiymatlar 3.1.5-jadvalda berilgan')
D.insert(p, '“Mikrokreditbank” ATB – 3,13', ', “Aloqabank” ATB – 3,248', tag=T)
D.save(F)
print(pardoz(F))
# Keyin: 3.1.3-rasm (sezgirlik) yorliqlari 10,5 → 9 pt.
