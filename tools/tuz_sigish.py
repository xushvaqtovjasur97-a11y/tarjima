"""Chizmalarda matn qutichalarga sig‘ishi uchun chizmalar balandligi oshiriladi (har bir qutida ≥15% zaxira).
1.3.1-rasm sarlavha qutichasi kengaytiriladi."""
import re, math, sys
sys.path.insert(0, '/home/user/tarjima/tools')
import docx
from docx.oxml.ns import qn
import sigish as S

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
LIMIT = 0.85
A, WPS, WPG = S.A, S.WPS, S.WPG
d = docx.Document(F)


def figures():
    for p in d.paragraphs:
        m = re.match(r'^(\d\.\d\.\d+)-rasm', p.text)
        if m and p._p.getprevious() is not None and p._p.getprevious().xpath('.//*[local-name()="wsp"]'):
            yield m.group(1), p._p.getprevious()


def ratio(el):
    rs = [S.need(sp) for sp in el.iter('{%s}wsp' % WPS)]
    rs = [r for r in rs if r and r['text']]
    return max(r['need'] / r['h'] for r in rs)


# 1) 1.3.1-rasm: sarlavha qutichasini kengaytirish
for k, el in figures():
    if k != '1.3.1':
        continue
    grp = el.find('.//{%s}wgp' % WPG)
    gx = grp.find('{%s}grpSpPr/{%s}xfrm' % (WPG, A))
    choff = int(gx.find('{%s}chOff' % A).get('x')); chext = int(gx.find('{%s}chExt' % A).get('cx'))
    for sp in el.iter('{%s}wsp' % WPS):
        t = ''.join(x.text or '' for x in sp.iter(qn('w:t')))
        if 'Bazel qo‘mitasi tasnifi asosida' in t:
            xf = sp.find('{%s}spPr/{%s}xfrm' % (WPS, A))
            ext = xf.find('{%s}ext' % A); off = xf.find('{%s}off' % A)
            new = int(chext * 0.66)
            ext.set('cx', str(new)); off.set('x', str(choff + (chext - new) // 2))
            print('1.3.1 sarlavha kengligi ->', new)

# 2) balandlikni oshirish
for k, el in figures():
    r = ratio(el)
    if r <= LIMIT:
        continue
    f = math.ceil(r / LIMIT * 100) / 100
    ext = el.xpath('.//*[local-name()="extent"]')[0]
    ext.set('cy', str(int(int(ext.get('cy')) * f)))
    gx = el.find('.//{%s}wgp/{%s}grpSpPr/{%s}xfrm' % (WPG, WPG, A))
    ge = gx.find('{%s}ext' % A)
    ge.set('cy', str(int(int(ge.get('cy')) * f)))
    for g in el.xpath('.//*[local-name()="Fallback"]//*[local-name()="group"]')[:1]:
        st = g.get('style')
        g.set('style', re.sub(r'height:([\d.]+)pt', lambda m: 'height:%.1fpt' % (float(m.group(1)) * f), st))
    print(k, 'nisbat %.2f -> balandlik x%.2f, yangi nisbat %.2f' % (r, f, ratio(el)))
d.save(F)
