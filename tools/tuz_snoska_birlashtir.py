"""Ketma-ket qo‘yilgan izoh belgilari (masalan, ⁴⁴⁴⁴⁴⁵⁴⁴⁶) bitta izohga birlashtiriladi: birinchi izoh qoladi,
qolgan manbalar unda nuqtali vergul bilan ketma-ket yoziladi (N.H. Sharipova avtoreferatidagi kabi).
Foydalanish: python3 tuz_snoska_birlashtir.py [docx]"""
import sys, copy
sys.path.insert(0, '/home/user/tarjima/tools')
from docx.oxml.ns import qn
import kamlib as K

F = sys.argv[1] if len(sys.argv) > 1 else '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
def tx(e): return ''.join(t.text or '' for t in e.iter(qn('w:t')))
FNS = {f.get(qn('w:id')): f for f in D.fx.findall(qn('w:footnote'))}

groups = []
for p in D.d.element.body.iter(qn('w:p')):
    cur = []
    for c in p.iter():
        if c.tag == qn('w:footnoteReference'):
            cur.append(c)
        elif c.tag == qn('w:t') and (c.text or '').strip():
            if len(cur) > 1: groups.append(cur)
            cur = []
    if len(cur) > 1: groups.append(cur)

def body_runs(fn):
    """izoh matni run'lari (footnoteRef va undan keyingi bo‘sh joysiz)"""
    ps = fn.findall(qn('w:p'))
    out = []
    for k, p in enumerate(ps):
        rs = [r for r in p if r.tag in (qn('w:r'), qn('w:hyperlink'))]
        rs = [r for r in rs if r.find(qn('w:footnoteRef')) is None]
        if k == 0:
            while rs and rs[0].tag == qn('w:r') and not tx(rs[0]).strip():
                rs = rs[1:]
            if rs:
                t0 = rs[0].find(qn('w:t'))
                if t0 is not None: t0.text = t0.text.lstrip()
        out += rs
    return out

def strip_end_dot(fn):
    ts = [t for t in fn.iter(qn('w:t')) if (t.text or '').strip()]
    t = ts[-1]
    s = t.text.rstrip()
    last = s.split()[-1][:-1] if s.endswith('.') else ''
    abbr = last.isalpha() and len(last) <= 3          # p., с., bb., gg. – qisqartma nuqtasi qoladi
    if s.endswith('.') and not s.endswith('..') and not abbr:
        s = s[:-1]
    t.text = s

n_groups = n_removed = 0
for g in groups:
    first = FNS[g[0].get(qn('w:id'))]
    lastp = first.findall(qn('w:p'))[-1]
    tmpl = [r for r in lastp.iter(qn('w:r')) if r.find(qn('w:t')) is not None][-1]
    for ref in g[1:]:
        fid = ref.get(qn('w:id'))
        fn = FNS[fid]
        strip_end_dot(first)
        sep = copy.deepcopy(tmpl)
        for e in list(sep):
            if e.tag != qn('w:rPr'): sep.remove(e)
        t = sep.makeelement(qn('w:t'), {}); t.text = '; '
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve'); sep.append(t)
        lastp.append(sep)
        for r in body_runs(fn):
            lastp.append(r)
        fn.getparent().remove(fn)
        run = ref.getparent()
        run.getparent().remove(run)
        n_removed += 1
    n_groups += 1
D.save(F)
print('guruhlar:', n_groups, 'olib tashlangan izohlar:', n_removed)
