"""Asosiy faylga yangi 1.1.5-rasm, uning tahlili va 9-ilovani qizil rangda kiritadi
(ma’lumotlar va matn – namuna_115.py dan)."""
import copy, re, sys
sys.path.insert(0, '/home/user/tarjima/tools')
from lxml import etree
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
import kamlib as K
import wordart as W
import namuna_115 as N
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
T = N.texts(lambda s: '[[FN:' + s + ']]')
P = lambda sub: D.para(sub)

# 1) kirish jumlasi
D.replace('Tashqi omillar orasida mualliflar ishonchni', 'Buni AQSh bank tizimi tajribasi ham tasdiqlaydi (1.1.5-rasm).', T['intro'], tag='115')

# 2) diagramma: chart1 qismini yangi XML va Excel bilan almashtirish, balandligi 12 sm
cap = P('1.1.5-rasm. AQSh banklarida moliyaviy barqarorlik pasayishining sabablari')
chp = cap._p.getprevious()
rid = [e.get(qn('r:id')) for e in chp.iter('{http://schemas.openxmlformats.org/drawingml/2006/chart}chart')][0]
part = D.d.part.related_parts[rid]
xrel = [r for r in part.rels.values() if r.reltype.endswith('/package')][0]
part._blob = N.xml.replace('r:id="rId1"', f'r:id="{xrel.rId}"').encode('utf-8')
xrel.target_part._blob = W._xlsx(N.rows)
ext = chp.find('.//{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent')
ext.set('cy', str(int(12.0 * W.EMU_CM)))

# 3) rasm nomi va manbasi
D.replace(cap, 'AQSh banklarida moliyaviy barqarorlik pasayishining sabablari, muammoli banklar sonidagi ulushi, %', T['caption'], tag='115')
fid = [r.get(qn('w:id')) for r in cap._p.iter(qn('w:footnoteReference'))][0]
fp = D.fn_para(fid)
old = ''.join(x.text or '' for x in fp._p.iter(qn('w:t'))).strip()
D.replace(fp, old, T['cap_fn'], tag='115')

# 4) tahlil: birinchi qism – yangi paragraf, ikkinchi qism – mavjud paragrafning boshi
an = P('Rasm ma’lumotlari muammoli banklarning deyarli barchasida (98 foiz)')
np_ = copy.deepcopy(an._p)
for e in list(np_):
    if e.tag != qn('w:pPr'):
        np_.remove(e)
tmpl = D._runs(an)[0]
for e in D._build(tmpl, T['a1']):
    np_.append(e)
an._p.addprevious(np_)
old1 = ('Rasm ma’lumotlari muammoli banklarning deyarli barchasida (98 foiz) aktivlar sifatining yomonligi, 90 foizida esa rejalashtirish va boshqaruvning zaifligi qayd etilganini ko‘rsatadi. Ichki auditning yo‘qligi va firibgarlik kamroq uchraydi, ammo ular ham muhim omil bo‘lib qolmoqda')
D.replace(an, old1, T['a2'], tag='115')
# eski 84-izoh (Lavrushin, 44-bet) endi rasm manbasida – havolani va izohni o‘chiramiz
refs = list(an._p.iter(qn('w:footnoteReference')))
old_ref = [r for r in refs if D.fn_para(r.get(qn('w:id'))) is not None and 'С. 44.' in ''.join(x.text or '' for x in D.fn_para(r.get(qn('w:id')))._p.iter(qn('w:t')))]
assert len(old_ref) == 1, len(old_ref)
ofid = old_ref[0].get(qn('w:id'))
run = old_ref[0].getparent(); run.getparent().remove(run)
for f in D.fx.findall(qn('w:footnote')):
    if f.get(qn('w:id')) == ofid:
        D.fx.remove(f)
D.replace(an, 'aktivlar sifati va boshqaruv sifati eng muhim bloklar qatorida bo‘lishi lozim', T['a2_end'], tag='115')
D.replace(an, 'Shu bilan birga, mualliflar erta ogohlantirish', 'Shu bilan birga, O.I. Lavrushin va I.D. Mamonova erta ogohlantirish', tag='115')

# 5) 9-ilova (8-ilova shablonidan)
body = D.d.element.body
def ptxt(e):
    return ''.join(x.text or '' for x in e.iter(qn('w:t')))


p8 = [Paragraph(e, D.d) for e in body.iterchildren(qn('w:p')) if ptxt(e).strip() == '8-ilova'][0]
brk = p8._p.getprevious()                      # pageBreakBefore bo‘sh paragraf
title8 = p8._p.getnext()
last = body[-2]                                # sectPr oldidagi oxirgi element (8-ilova manbasi)
manba8 = last


def clone_par(src, text, after):
    e = copy.deepcopy(src)
    for x in list(e):
        if x.tag != qn('w:pPr'):
            e.remove(x)
    if text:
        tm = [r for r in src.iter(qn('w:r')) if r.find(qn('w:t')) is not None][0]
        for r in D._build(tm, text):
            e.append(r)
    after.addnext(e)
    return e


e1 = clone_par(brk, '', manba8)
e2 = clone_par(p8._p, '9-ilova', e1)
e3 = clone_par(title8, N.TAB_TITLE, e2)
# jadval
WID = [600, 2700, 700, 6300] + [420] * 7 + [1330]
tbl8 = title8.getnext()
tb = etree.Element(qn('w:tbl'))
tb.append(copy.deepcopy(tbl8.find(qn('w:tblPr'))))
grid = etree.SubElement(tb, qn('w:tblGrid'))
for w in WID:
    etree.SubElement(grid, qn('w:gridCol')).set(qn('w:w'), str(w))
hdr_tc = tbl8.find(qn('w:tr')).find(qn('w:tc'))
for ri, vals in enumerate([N.HDR] + N.ROWS):
    tr = etree.SubElement(tb, qn('w:tr'))
    trpr = etree.SubElement(tr, qn('w:trPr')); etree.SubElement(trpr, qn('w:cantSplit'))
    if ri == 0:
        etree.SubElement(trpr, qn('w:tblHeader'))
    for ci, (v, w) in enumerate(zip(vals, WID)):
        tc = etree.SubElement(tr, qn('w:tc'))
        tcpr = etree.SubElement(tc, qn('w:tcPr'))
        tw = etree.SubElement(tcpr, qn('w:tcW')); tw.set(qn('w:w'), str(w)); tw.set(qn('w:type'), 'dxa')
        if ri == 0 or ri == len(N.ROWS):
            sh = etree.SubElement(tcpr, qn('w:shd')); sh.set(qn('w:val'), 'clear'); sh.set(qn('w:color'), 'auto'); sh.set(qn('w:fill'), 'DEEBF7')
        p = etree.SubElement(tc, qn('w:p'))
        ppr = etree.SubElement(p, qn('w:pPr'))
        if ci in (0, 2) or 4 <= ci <= 10 or ri == 0:
            etree.SubElement(ppr, qn('w:jc')).set(qn('w:val'), 'center')
        r = etree.SubElement(p, qn('w:r'))
        rpr = etree.SubElement(r, qn('w:rPr'))
        if ri == 0 or ri == len(N.ROWS):
            etree.SubElement(rpr, qn('w:b'))
        etree.SubElement(rpr, qn('w:color')).set(qn('w:val'), K.RED)
        etree.SubElement(rpr, qn('w:sz')).set(qn('w:val'), '18')
        t = etree.SubElement(r, qn('w:t')); t.text = v
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
e3.addnext(tb)
izoh8 = [e for e in body.iterchildren(qn('w:p')) if ptxt(e).startswith('Izoh: har bir katakda vaznlar ekspert')][0]
e4 = clone_par(izoh8, N.IZOH_TAB, tb)
e5 = clone_par(manba8, N.SRC_TAB, e4)
K.LOG.append(('115', 'ilova', '9-ilova', N.TAB_TITLE))
# yangi lxml elementlariga Python havolalari qolmasin (aks holda python-docx ularni oddiy element sifatida ko‘radi)
del r, t, p, tc, tr, rpr, ppr, tcpr, tw, trpr, tb, grid, e1, e2, e3, e4, e5, np_, run, e
import gc; gc.collect()

# 6) adabiyotlar
def add_after(anchor_sub, text):
    a = [p for p in D.paras() if re.match(r'^\d+\. ', p.text) and anchor_sub in p.text][-1]
    D.new_para_after(a, '0. ' + text, tag='115')

add_after('Revisions to the Large Financial Institution Rating System', 'Board of Governors of the Federal Reserve System, Office of Inspector General. Summary Analysis of Failed Bank Reviews. – Washington, DC, September 2011. – https://oig.federalreserve.gov/reports/Cross_Cutting_Final_Report_9-30-11.pdf.')
add_after('Summary Analysis of Failed Bank Reviews', 'Board of Governors of the Federal Reserve System, Office of Inspector General. Material Loss Review of Silicon Valley Bank: Report 2023-SR-B-013. – Washington, DC, September 25, 2023.')
add_after('Federal Deposit Insurance Corporation Improvement Act', 'Federal Deposit Insurance Corporation, Office of Inspector General. Material Loss Review of First Republic Bank: Report EVAL-24-03. – Arlington, VA, November 2023. – https://www.fdicoig.gov/reports-publications/bank-failures.')
add_after('Cornett M.M., Guo L.', 'Correia S., Luck S., Verner E. Failing Banks // Quarterly Journal of Economics. – 2026. – Vol. 141, No. 1. – P. 147–204.')
start = False; n = 0
for p in D.paras():
    if p.text.strip() == 'FOYDALANILGAN ADABIYOTLAR RO‘YXATI':
        start = True; continue
    if start and p.text.strip() == 'ILOVALAR':
        break
    if start:
        m = re.match(r'^(\d+)\. ', p.text)
        if m:
            n += 1
            r0 = D._runs(p)[0].find(qn('w:t'))
            if r0.text.startswith(m.group(1) + '.'):
                r0.text = str(n) + r0.text[len(m.group(1)):]
            elif r0.text == m.group(1):
                r0.text = str(n)
print('adabiyotlar:', n)
hp = P('Dissertatsiyaning tuzilishi va hajmi.')
D.replace(hp, re.search(r'ro‘yxati \d+ ta manbani', hp.text).group(0), f'ro‘yxati {n} ta manbani', tag='115')
D.save(F)
print('pardoz:', pardoz(F))

sys.argv = ['x']
_src = open('/home/user/tarjima/tools/kam_tuzat.py').read().split('# ---------------- mundarija sahifa raqamlarini yangilash')[1].split("print('mundarija:'")[0]
exec(_src.split('\n', 1)[1])
print('mundarija:', mundarija(F), mundarija(F))
print(len(K.LOG), 'amal')
