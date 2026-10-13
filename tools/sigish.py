"""Chizmalardagi (wps) qutichalarga matn sig‘ishini taxminiy tekshirish (Word o‘lchamlari bo‘yicha, zaxira bilan)."""
import re, sys
import docx
from docx.oxml.ns import qn

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
WPG = 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup'
EMU_PT = 12700
# Times New Roman o‘rtacha belgi kengligi (em ulushida), Word uchun zaxira bilan
CW = 0.50
CWB = 0.54


from PIL import ImageFont
_F = {}


def tw(text, sz, bold=False, italic=False):
    k = (bold, italic)
    if k not in _F:
        nm = {(False, False): 'Regular', (True, False): 'Bold', (False, True): 'Italic', (True, True): 'BoldItalic'}[k]
        _F[k] = ImageFont.truetype('/usr/share/fonts/truetype/liberation/LiberationSerif-%s.ttf' % nm, 1000)
    return _F[k].getlength(text) / 1000 * sz


def group_scale(sp):
    """Shakl ota-guruhlari bo‘yicha (sx, sy) masshtab."""
    sx = sy = 1.0
    g = sp.getparent()
    while g is not None:
        if g.tag in ('{%s}wgp' % WPG, '{%s}grpSp' % WPG):
            x = g.find('{%s}grpSpPr/{%s}xfrm' % (WPG, A))
            if x is not None:
                ext = x.find('{%s}ext' % A); ch = x.find('{%s}chExt' % A)
                if ext is not None and ch is not None and int(ch.get('cx')) and int(ch.get('cy')):
                    sx *= int(ext.get('cx')) / int(ch.get('cx')); sy *= int(ext.get('cy')) / int(ch.get('cy'))
        g = g.getparent()
    return sx, sy


def need(sp):
    xf = sp.find('{%s}spPr/{%s}xfrm' % (WPS, A))
    if xf is None:
        return None
    ext = xf.find('{%s}ext' % A)
    sx, sy = group_scale(sp)
    w = int(ext.get('cx')) * sx / EMU_PT; h = int(ext.get('cy')) * sy / EMU_PT
    bp = sp.find('{%s}bodyPr' % WPS)
    li = int(bp.get('lIns', 91440)) / EMU_PT; ri = int(bp.get('rIns', 91440)) / EMU_PT
    ti = int(bp.get('tIns', 45720)) / EMU_PT; bi = int(bp.get('bIns', 45720)) / EMU_PT
    tb = sp.find('{%s}txbx' % WPS)
    if tb is None:
        return None
    avail = w - li - ri
    tot = 0.0; text = ''
    for p in tb.iter(qn('w:p')):
        runs = list(p.iter(qn('w:r')))
        if not runs:
            # bo‘sh paragraf ham joy egallaydi
            sz = 11
            tot += sz * 1.15; continue
        line_w = 0.0; lines = 1; maxsz = 0
        words = []
        for r in runs:
            s = r.find(qn('w:rPr'))
            szv = s.find(qn('w:sz')) if s is not None else None
            sz = int(szv.get(qn('w:val'))) / 2 if szv is not None else 11
            b = s is not None and s.find(qn('w:b')) is not None
            maxsz = max(maxsz, sz)
            t = ''.join(x.text or '' for x in r.iter(qn('w:t')))
            text += t
            for wd in re.split(r'(\s+)', t):
                if wd:
                    words.append((wd, sz, b))
        for wd, sz, b in words:
            ww = tw(wd, sz, b) * 1.03
            if line_w + ww > avail and not wd.isspace() and line_w > 0:
                lines += 1; line_w = ww
            else:
                line_w += ww
        sp_after = 0
        ppr = p.find(qn('w:pPr'))
        spc = ppr.find(qn('w:spacing')) if ppr is not None else None
        if spc is not None:
            sp_after = (int(spc.get(qn('w:after'), 0)) + int(spc.get(qn('w:before'), 0))) / 20
            lr = spc.get(qn('w:line'))
            mult = int(lr) / 240 if lr and spc.get(qn('w:lineRule'), 'auto') == 'auto' else 1.0
        else:
            mult = 1.0
        tot += lines * maxsz * 1.15 * mult + sp_after
        text += ' / '
    return dict(w=w, h=h, need=tot + ti + bi, text=text.strip(' /'))


def check(path, verbose=False):
    d = docx.Document(path)
    out = {}
    for p in d.paragraphs:
        m = re.match(r'^(\d\.\d\.\d+)-rasm', p.text)
        if not m:
            continue
        prev = p._p.getprevious()
        if prev is None:
            continue
        seen = set(); bad = []
        for sp in prev.iter('{%s}wsp' % WPS):
            r = need(sp)
            if not r or not r['text']:
                continue
            key = (r['text'], round(r['w']), round(r['h']))
            if key in seen:   # mc:Fallback dagi nusxa
                continue
            seen.add(key)
            if r['need'] > r['h'] * 1.0:
                bad.append(r)
        if bad:
            out[m.group(1)] = bad
    return out


if __name__ == '__main__':
    res = check(sys.argv[1] if len(sys.argv) > 1 else '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx')
    for k, v in res.items():
        print('==', k, len(v))
        for r in v:
            print('   h=%.0f need=%.0f w=%.0f | %s' % (r['h'], r['need'], r['w'], r['text'][:90]))
