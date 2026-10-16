"""4-bob tekshiruvi natijalari bo‘yicha tuzatishlar (ko‘k rangda).

Tahrirlar: tools/bob4_tahrirlar/edits4[123].json
Turlari: para, cell, fn, chart, chart_str, tx.
"""
import json, sys
sys.path.insert(0, '/home/user/tarjima/tools')
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from lxml import etree
import kamlib as K
from pardoz import pardoz

K.RED = '0070C0'            # o‘zgarishlar ko‘k rangda
F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
SRC = '/home/user/tarjima/tools/bob4_tahrirlar/'
E = []
for k in ('41', '42', '43'):
    E += json.load(open(SRC + f'edits{k}.json'))

D = K.Doc(F)
PS = D.d.paragraphs           # indekslar tahrirdan oldingi holat bo‘yicha
CNS = '{http://schemas.openxmlformats.org/drawingml/2006/chart}'
ANS = '{http://schemas.openxmlformats.org/drawingml/2006/main}'


def text_of(pe):
    return ''.join(t.text or '' for t in pe.iter(qn('w:t')))


def minimal(p, old, new):
    """Faqat o‘zgargan qismni almashtiradi (umumiy boshlanish va oxir saqlanadi)."""
    a = 0
    while a < min(len(old), len(new)) and old[a] == new[a]:
        a += 1
    b = 0
    while b < min(len(old), len(new)) - a and old[-1 - b] == new[-1 - b]:
        b += 1
    # so‘z chegarasigacha kengaytirish (o‘qilishi uchun)
    while a > 0 and old[a - 1] not in ' (':
        a -= 1
    while b > 0 and old[len(old) - b] not in ' ,.;)':
        b -= 1
    mo, mn = old[a:len(old) - b], new[a:len(new) - b]
    full = text_of(p._p)
    pos = full.find(old)
    assert pos >= 0 and full.count(old) == 1, (old[:60],)
    if not mo:                # sof qo‘shimcha – butun bo‘lak almashtiriladi
        D.replace(p, old, new, occ=full[:pos].count(old))
        return
    occ = full[:pos + a].count(mo)
    D.replace(p, mo, mn, occ=occ)


def table_by_caption(cap):
    return D.table_after(cap)


log = []
for e in E:
    t = e['type']
    if t == 'para':
        p = PS[e['p']]
        assert text_of(p._p).count(e['old']) == 1, ('para', e['p'], e['old'][:60])
        minimal(p, e['old'], e['new'])
    elif t == 'cell':
        tb = table_by_caption(e['table_caption'])
        hits = []
        seen = set()
        for row in tb.rows:
            for c in row.cells:
                if c._tc in seen:
                    continue
                seen.add(c._tc)
                for p in c.paragraphs:
                    if e['old'] in text_of(p._p):
                        hits.append(p)
        assert len(hits) == 1, ('cell', e['table_caption'][:40], e['old'][:40], len(hits))
        minimal(hits[0], e['old'], e['new'])
    elif t == 'fn':
        f = [f for f in D.fx.findall(qn('w:footnote')) if f.get(qn('w:id')) == str(e['id'])]
        assert len(f) == 1, ('fn', e['id'])
        pe = [pp for pp in f[0].findall(qn('w:p')) if e['old'] in text_of(pp)]
        assert len(pe) == 1, ('fn', e['id'], e['old'][:50])
        minimal(Paragraph(pe[0], None), e['old'], e['new'])
    elif t == 'chart':
        D.chart_set(e['part'], e['ser'], e['values'])
        K.chart_sync_xlsx(D, e['part'])
    elif t == 'chart_str':
        x, part = D.chart_xml(e['part'])
        k = 0
        for v in list(x.iter(CNS + 'v')) + list(x.iter(ANS + 't')):
            if v.text == e['old']:
                v.text = e['new']; k += 1
        assert k >= 1, ('chart_str', e['old'])
        part._blob = etree.tostring(x, xml_declaration=True, encoding='UTF-8', standalone=True)
        K.chart_sync_xlsx(D, e['part'])
    elif t == 'tx':
        k = 0
        for tb in D.d.element.body.iter(qn('w:txbxContent')):
            for pe in tb.iter(qn('w:p')):
                ts = list(pe.iter(qn('w:t')))
                s = ''.join(x.text or '' for x in ts)
                if e['old'] in s:
                    ts[0].text = s.replace(e['old'], e['new'])
                    ts[0].set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                    for x in ts[1:]:
                        x.text = ''
                    k += 1
        assert k >= 1, ('tx', e['old'][:50])
    else:
        raise ValueError(t)
    log.append(t)
D.save(F)
import collections
print(len(log), dict(collections.Counter(log)), pardoz(F))
