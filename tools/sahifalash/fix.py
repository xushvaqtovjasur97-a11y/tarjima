"""Sahifa pastidagi bo‘sh joylarni bartaraf etish: jadval/rasm keyingi sahifaga o‘tib ketgan bo‘lsa,
undan keyingi tahlil xatboshisini blokdan oldinga ko‘chirish (so‘zlarni moslab) yoki diagrammani kichraytirish.

Foydalanish: python3 fix.py <pdf> <docx> [--dry]
Har bir bo‘lim (§) uchun birinchi muammoli blok tuzatiladi (bo‘limlar yangi sahifadan boshlanadi).
"""
import re, sys, subprocess, collections, json
sys.path.insert(0, '/home/user/tarjima/tools')
import kamlib as K
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

K.RED = '0070C0'
PDF, DOCX = sys.argv[1], sys.argv[2]
DRY = '--dry' in sys.argv
MIN_GAP = 2.5 * 24            # pt

# ---------------- PDF qatorlari
html = subprocess.run(['pdftotext', '-bbox', PDF, '-'], capture_output=True, text=True).stdout
pages = re.split(r'<page width="[\d.]+" height="[\d.]+">', html)[1:]
SEC = re.compile(r'^(KIRISH|XULOSA|FOYDALANILGAN|ILOVALAR|[IV]+ BOB|[1-4]\.[1-3]-§|[IV]+ bob bo‘yicha xulosalar)')
PG = []
for k, pg in enumerate(pages):
    W = [(float(a), float(b), float(c), float(d), t) for a, b, c, d, t in
         re.findall(r'<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">([^<]*)</word>', pg)]
    lines = collections.OrderedDict()
    for w in W:
        lines.setdefault(round(w[3]), []).append(w)
    L = []
    for y, ws in sorted(lines.items()):
        h = max(w[3] - w[1] for w in ws)
        L.append(dict(y=y, h=h, top=min(w[1] for w in ws), t=' '.join(w[4] for w in sorted(ws))))
    L = [l for l in L if l['y'] < 795]
    i = len(L) - 1; fn_top = None
    while i >= 0 and L[i]['h'] <= 11.6:
        fn_top = L[i]['top']; i -= 1
    body = L[:i + 1]
    limit = (fn_top - 14) if fn_top is not None else 786
    bottom = body[-1]['y'] if body else 57
    PG.append(dict(n=k + 1, body=body, gap=limit - bottom, sec=bool(body and SEC.match(body[0]['t']))))
for a, b in zip(PG, PG[1:]):
    a['sec_end'] = b['sec']
PG[-1]['sec_end'] = True
secid = 0
for p in PG:
    if p['sec']:
        secid += 1
    p['secid'] = secid

# ---------------- docx bloklari
D = K.Doc(DOCX)
body_el = list(D.d.element.body.iterchildren())


def ptxt(el):
    return ''.join(t.text or '' for t in el.iter(qn('w:t')))


def has_draw(el):
    return el.find('.//' + qn('w:drawing')) is not None or el.find('.//' + qn('w:pict')) is not None


def is_text_para(el):
    if el.tag != qn('w:p') or has_draw(el):
        return False
    t = ptxt(el).strip()
    if not t or len(t) < 60:
        return False
    if re.match(r'^(\d\.\d\.\d+-rasm\.|\d\.\d\.\d+-jadval$|Izoh|Manba|[IV]+ BOB|\d\.\d-§|\d+-ilova)', t):
        return False
    if el.find('.//' + qn('w:pageBreakBefore')) is not None:
        return False
    return True


blocks = []
i = 0
while i < len(body_el):
    el = body_el[i]
    if el.tag == qn('w:p') and re.match(r'^\d\.\d\.\d+-jadval$', ptxt(el).strip()):
        num = ptxt(el).strip()[:-len('-jadval')]
        j = i + 1
        while j < len(body_el) and body_el[j].tag != qn('w:tbl'):
            j += 1
        j += 1
        while j < len(body_el) and body_el[j].tag == qn('w:p') and re.match(r'^(Izoh|Manba)', ptxt(body_el[j]).strip()):
            j += 1
        blocks.append(dict(kind='jadval', num=num, start=i, end=j))   # end – blokdan keyingi element indeksi
        i = j; continue
    if el.tag == qn('w:p') and has_draw(el) and el.getparent().tag == qn('w:body'):
        # rasm: keyingi elementlar ichida sarlavha
        j = i + 1
        while j < len(body_el) and j < i + 4 and not re.match(r'^\d\.\d\.\d+-rasm\.', ptxt(body_el[j]).strip()):
            j += 1
        if j < len(body_el) and re.match(r'^\d\.\d\.\d+-rasm\.', ptxt(body_el[j]).strip()):
            num = re.match(r'^(\d\.\d\.\d+)-rasm\.', ptxt(body_el[j]).strip()).group(1)
            k = j + 1
            while k < len(body_el) and body_el[k].tag == qn('w:p') and re.match(r'^(Izoh|Manba)', ptxt(body_el[k]).strip()):
                k += 1
            blocks.append(dict(kind='rasm', num=num, start=i, end=k, cap=j))
            i = k; continue
    i += 1


def find_page(regex):
    for p in PG:
        for li, l in enumerate(p['body']):
            if re.match(regex, l['t']):
                return p, li
    return None, None


cands = []
for b in blocks:
    if b['kind'] == 'jadval':
        p, li = find_page(r'^' + re.escape(b['num']) + r'-jadval$')
        if p is None or li != 0:
            continue
    else:
        p, li = find_page(r'^' + re.escape(b['num']) + r'-rasm\.')
        if p is None:
            continue
        above = p['body'][:li]
        if any(l['h'] >= 15 and len(l['t']) > 40 for l in above):   # rasm ustida oddiy matn bor
            continue
    prev = PG[p['n'] - 2]
    if prev['sec_end'] or prev['gap'] < MIN_GAP:
        continue
    b.update(page=p['n'], gap=prev['gap'], secid=prev['secid'],
             block_h=(p['body'][li]['y'] - 57) if b['kind'] == 'rasm' else None)
    cands.append(b)

def fixable(b):
    nxt = body_el[b['end']] if b['end'] < len(body_el) else None
    if nxt is not None and is_text_para(nxt):
        return True
    el = body_el[b['start']]
    return (b['kind'] == 'rasm' and b['block_h'] and b['gap'] >= 0.78 * b['block_h']
            and el.find('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}chart') is not None)


first = {}
for b in sorted(cands, key=lambda x: x['page']):
    if b['secid'] in first:
        continue
    if fixable(b) or not any(fixable(c) for c in cands if c['secid'] == b['secid'] and c is not b):
        first[b['secid']] = b

REF = [(r'^Rasmdan ', '{n}-rasmdan '), (r'^Rasmda ', '{n}-rasmda '), (r'^Rasm ma’lumotlari', '{n}-rasm ma’lumotlari'),
       (r'^Jadvaldan ', '{n}-jadvaldan '), (r'^Jadvalda ', '{n}-jadvalda '), (r'^Jadval ma’lumotlari', '{n}-jadval ma’lumotlari'),
       (r'^Ushbu rasm', '{n}-rasm'), (r'^Ushbu jadval', '{n}-jadval'), (r'^Mazkur rasm', '{n}-rasm'), (r'^Mazkur jadval', '{n}-jadval')]
log = []
for sid, b in sorted(first.items()):
    nxt = body_el[b['end']] if b['end'] < len(body_el) else None
    desc = f"§{sid} {b['num']}-{b['kind']} (sah. {b['page']}, oldingi sahifada {b['gap'] / 24:.1f} qator bo‘sh)"
    if nxt is not None and is_text_para(nxt):
        t = ptxt(nxt)
        action = 'ko‘chirish'
        if not DRY:
            first_el = body_el[b['start']]
            nxt.getparent().remove(nxt)
            first_el.addprevious(nxt)
            p = Paragraph(nxt, D.d)
            for rx, rep in REF:
                m = re.match(rx, t)
                if m:
                    D.replace(p, m.group(0), rep.format(n=b['num']))
                    break
        log.append(desc + f" → keyingi xatboshi blokdan oldinga ko‘chirildi: «{t[:70]}…»")
    elif b['kind'] == 'rasm' and b['block_h'] and b['gap'] >= 0.78 * b['block_h']:
        el = body_el[b['start']]
        ext = el.find('.//{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent')
        is_chart = el.find('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}chart') is not None
        if ext is not None and is_chart:
            need = b['block_h'] - b['gap'] + 8           # pt
            cy = int(ext.get('cy')); cx = int(ext.get('cx'))
            f = max(0.78, 1 - need * 12700 / cy)
            if not DRY:
                ext.set('cy', str(int(cy * f))); ext.set('cx', str(int(cx * f)))
            log.append(desc + f" → diagramma {f:.2f} marta kichraytirildi")
        else:
            log.append(desc + " → QO‘LDA: ko‘chiriladigan xatboshi yo‘q, sxema")
    else:
        log.append(desc + f" → QO‘LDA: keyingi element matn emas («{ptxt(nxt)[:50] if nxt is not None else ''}»)")
if not DRY:
    D.save(DOCX)
print('\n'.join(log) if log else 'muammo topilmadi')
