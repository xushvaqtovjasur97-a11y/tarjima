"""Jadval/rasm keyingi sahifaga o‘tib, oldingi sahifa pastida bo‘sh joy qolganda: blokdan keyingi matndan bo‘sh joyni
to‘ldiradigan miqdordagi gaplarni blokdan oldinga olib chiqish, qolgan gaplar blokdan keyin davom etadi.
Foydalanish: python3 fix2.py <docx> <pdf> <bloklar.json> [--dry]
Har bir bo‘limda birinchi muammoli blok tuzatiladi (keyingi sahifalar siljiydi – takrorlash kerak)."""
import sys, re, json, copy, math, subprocess
sys.path.insert(0, '/home/user/tarjima/tools')
import kamlib as K
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from PIL import ImageFont

K.RED = '000000'
DOCX, PDF, BL = sys.argv[1:4]
DRY = '--dry' in sys.argv
LINE = 24.15; WIDTH = 467.75; IND = 35.45
FONT = ImageFont.truetype('/usr/share/fonts/truetype/liberation/LiberationSerif-Regular.ttf', 140)
def wpt(s): return FONT.getlength(s) / 10
def nlines(text, indent=IND):
    n, cur, sp = 1, indent, wpt(' ')
    first = True
    for w in text.split():
        L = wpt(w)
        if first: cur += L; first = False
        elif cur + sp + L > WIDTH: n += 1; cur = L
        else: cur += sp + L
    return n

G = json.loads(subprocess.run(['python3', '/home/user/tarjima/tools/sahifalash/gaps.py', PDF], capture_output=True, text=True).stdout)
GP = {g['p']: g for g in G}
sec = 0
for g in G:
    if g['sec']: sec += 1
    g['secid'] = sec
B = json.load(open(BL))
NPAGES = [x for x in B if x['kind'] == 'pages'][0]['n']
JH = {x['num']: x['p'] for x in B if x['kind'] == 'jhead'}
CAP = {x['num']: x['p'] for x in B if x['kind'] == 'cap'}
TB = {x['name']: x for x in B if x['kind'] == 'tbl'}

D = K.Doc(DOCX)
body = D.d.element.body
els = list(body)
def tx(e): return ''.join(t.text or '' for t in e.iter(qn('w:t')))
def has_draw(e): return e.find('.//' + qn('w:drawing')) is not None or e.find('.//' + qn('w:pict')) is not None
IZOH = re.compile(r'^(Izoh|Manba|\*)')
def is_text(e):
    if e is None or e.tag != qn('w:p') or has_draw(e): return False
    t = tx(e).strip()
    if len(t) < 60 or re.match(r'^(\d\.\d\.\d+-(rasm\.|jadval)|Izoh|Manba|[IV]+ BOB|\d\.\d-§|\d+-ilova)', t): return False
    if e.find('.//' + qn('w:pageBreakBefore')) is not None or e.find('.//' + qn('w:br')) is not None and 'page' in str(e.xpath('.//w:br/@w:type')): return False
    st = e.find('.//' + qn('w:pStyle'))
    if st is not None and re.match(r'(?i)heading|toc|[1-9]$', st.get(qn('w:val'))): return False
    return True

FN = {f.get(qn('w:id')): tx(f) for f in D.fx.findall(qn('w:footnote'))}

# bloklar: (turi, raqami, boshlanish elementi, oxiridan keyingi element, balandlik pt)
blocks = []
tno = 0
for i, e in enumerate(els):
    if e.tag == qn('w:tbl'):
        tno += 1; e.set('_tname', f'Table{tno}')
for i, e in enumerate(els):
    t = tx(e).strip() if e.tag == qn('w:p') else ''
    m = re.match(r'^(\d\.\d\.\d+)-jadval$', t)
    if m:
        j = i + 1
        while j < len(els) and els[j].tag != qn('w:tbl'): j += 1
        tb = TB.get(els[j].get('_tname'))
        h = (tb['y1'] - tb['y0']) / 100 / 25.4 * 72 + 55 if tb and tb['p0'] == tb['p1'] else None
        j += 1
        while j < len(els) and els[j].tag == qn('w:p') and (IZOH.match(tx(els[j]).strip()) or tx(els[j]).strip() == '') and not has_draw(els[j]): j += 1
        blocks.append(dict(kind='jadval', num=m.group(1), s=i, e=j, page=JH.get(m.group(1)), h=h))
    m = re.match(r'^(\d\.\d\.\d+)-rasm\.', t)
    if m:
        k = i - 1
        while k > i - 6 and not (els[k].tag == qn('w:p') and has_draw(els[k])): k -= 1
        if not (els[k].tag == qn('w:p') and has_draw(els[k])): continue
        j = i + 1
        while j < len(els) and els[j].tag == qn('w:p') and (IZOH.match(tx(els[j]).strip()) or tx(els[j]).strip() == '') and not has_draw(els[j]): j += 1
        blocks.append(dict(kind='rasm', num=m.group(1), s=k, e=j, page=CAP.get(m.group(1)), h=None))

def letters(s): return re.sub(r'[^A-Za-z‘’]', '', s)[-18:]
cands = []
for b in blocks:
    P = b['page']
    if not P or P < 2: continue
    g = GP.get(P - 1)
    if g is None or g.get('sec_end') or g['lines'] < 1.4: continue
    prev = els[b['s'] - 1]
    if not is_text(prev) or letters(tx(prev)) not in letters(g['last']) and letters(g['last'])[-10:] not in letters(tx(prev)):
        continue
    if b['kind'] == 'rasm':
        # rasm balandligi: sahifadagi sarlavha qatorigacha (pdf)
        b['h'] = None
    b.update(gap=g['lines'], secid=g['secid'])
    cands.append(b)

SENT = re.compile(r'(?<=[.!?…])(?<![A-Z]\.)(?<!\b[A-Z][a-z]\.)\s+(?=[A-ZO‘“"(0-9А-Я])')
def sentences(p_el):
    t = tx(p_el)
    cuts = [0] + [m.end() for m in SENT.finditer(t) if not re.search(r'\b(?:[A-Z]|Ye|Sh|Ch|Yu|Ya)\.$', t[:m.start()])] + [len(t)]
    return t, [(a, b) for a, b in zip(cuts, cuts[1:]) if t[a:b].strip()]

def fn_ids_between(p_el, upto_run):
    ids = []
    for c in p_el:
        if c is upto_run: break
        for r in c.iter(qn('w:footnoteReference')): ids.append(r.get(qn('w:id')))
    return ids

REF = [(r'^Rasmdan ', '{n}-rasmdan '), (r'^Rasmda ', '{n}-rasmda '), (r'^Rasm ma’lumotlari', '{n}-rasm ma’lumotlari'),
       (r'^Jadvaldan ', '{n}-jadvaldan '), (r'^Jadvalda ', '{n}-jadvalda '), (r'^Jadval ma’lumotlari', '{n}-jadval ma’lumotlari'),
       (r'^Ushbu rasm', '{n}-rasm'), (r'^Ushbu jadval', '{n}-jadval'), (r'^Mazkur rasm', '{n}-rasm'), (r'^Mazkur jadval', '{n}-jadval')]

done, log = set(), []
for b in sorted(cands, key=lambda x: x['page']):
    if b['secid'] in done: continue
    done.add(b['secid'])
    # blokdan keyingi matn xatboshilari
    paras, j = [], b['e']
    while j < len(els) and is_text(els[j]) and len(paras) < 4:
        paras.append(els[j]); j += 1
    desc = f"{b['num']}-{b['kind']} (sah. {b['page']}, oldingida {b['gap']:.1f} qator bo‘sh)"
    if not paras:
        log.append(desc + ' → QO‘LDA: blokdan keyin matn yo‘q'); continue
    # gaplar bo‘yicha tanlash
    plan, acc_lines, fn_lines, best = [], 0.0, 0.0, None
    had_fn = GP[b['page'] - 1]['lines'] is not None
    for pi, pe in enumerate(paras):
        t, ss = sentences(pe)
        for si in range(len(ss)):
            moved_txt = t[:ss[si][1]].strip()
            L = nlines(moved_txt)
            ids = []
            for r in pe.iter(qn('w:footnoteReference')):
                ids.append(r.get(qn('w:id')))
            # izoh havolalari soni moved qismida (taxminiy: matn ulushi bo‘yicha)
            frac = ss[si][1] / max(len(t), 1)
            n_ids = int(round(len(ids) * frac + 0.01)) if ids else 0
            fl = sum(math.ceil(len(FN.get(x, '')) / 118) * 0.48 for x in ids[:n_ids])
            tot_body = acc_lines + L
            tot_fn = fn_lines + fl
            need = b['gap'] - tot_fn
            cand = dict(pi=pi, si=si, body=tot_body, need=need, last=(si == len(ss) - 1))
            if tot_body >= need - 0.3:
                best = cand; break
            best = cand
        else:
            acc_lines += nlines(t); fn_lines += sum(math.ceil(len(FN.get(x, '')) / 118) * 0.48 for x in ids)
            continue
        break
    over = best['body'] - best['need']
    if b['h'] and b['h'] + max(over, 0) * LINE > 700 and not (best['pi'] == 0 and best['si'] == 0):
        log.append(desc + f' → blok baland ({b["h"]:.0f} pt), ortiqcha {over:.1f} qator – bitta kam gap')
        if best['si'] > 0: best['si'] -= 1; best['last'] = False
        elif best['pi'] > 0: best['pi'] -= 1; best['si'] = len(sentences(paras[best['pi']])[1]) - 1; best['last'] = True
    if b['h'] and b['h'] + max(over, 0) * LINE > 720:
        log.append(desc + ' → QO‘LDA: blok + matn sahifaga sig‘maydi'); continue
    first = els[b['s']]
    moved_desc = []
    if not DRY:
        for pi in range(best['pi'] + 1):
            pe = paras[pi]
            if pi < best['pi'] or best['last']:
                first.addprevious(pe); moved = pe
            else:
                P_ = Paragraph(pe, None)
                t, ss = sentences(pe)
                cut = ss[best['si']][1]
                while cut < len(t) and t[cut] == ' ': cut += 1
                D._normalize(P_)
                run = D._split_at(P_, cut)
                rest = copy.deepcopy(pe)
                kids = [c for c in pe if c.tag != qn('w:pPr')]
                k = kids.index(run)
                for c in kids[k:]: pe.remove(c)
                kidsr = [c for c in rest if c.tag != qn('w:pPr')]
                for c in kidsr[:k]: rest.remove(c)
                ts = list(pe.iter(qn('w:t')))
                if ts: ts[-1].text = ts[-1].text.rstrip()
                pe.addnext(rest)
                first.addprevious(pe); moved = pe
            if pi == 0:
                t0 = tx(moved)
                for rx, rep in REF:
                    mm = re.match(rx, t0)
                    if mm:
                        D.replace(Paragraph(moved, None), mm.group(0), rep.format(n=b['num'])); break
            moved_desc.append(tx(moved)[:50])
    log.append(desc + f" → {best['pi'] + 1} xatboshidan gaplar oldinga (taxm. {best['body']:.1f} qator, kerak {best['need']:.1f}): " + ' | '.join(moved_desc))
if not DRY:
    D.save(DOCX)
print('\n'.join(log) if log else 'muammo topilmadi')
