"""Adabiyotlar ro‘yxatiga 11 ta MHXS hisoboti (qizil), qayta raqamlash; 3.2.2-jadval izohi umumiy shaklda; Kirishda manbalar soni."""
import re, sys
sys.path.insert(0, '/home/user/tarjima/tools')
from docx.oxml.ns import qn
import kamlib as K
import namuna_adab11 as N
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
# 1) 3.2.2-jadval izohi
tcap = [p for p in D.d.paragraphs if p.text.strip() == 'Integral indeks bloklarining tayanch va empirik vaznlari'][0]
fid = [r.get(qn('w:id')) for r in tcap._p.iter(qn('w:footnoteReference'))][0]
fp = D.fn_para(fid=fid)
old = ''.join(t.text or '' for t in fp._p.iter(qn('w:t')))
pre = 'Tadqiqot natijasida muallif tomonidan '
i = old.index(pre) + len(pre)
D.replace(fp, old[i:], N.FN_NEW[len(pre):], tag='adab')


def bib_paras():
    on = False
    out = []
    for p in D.d.paragraphs:
        if p.text.strip() == 'FOYDALANILGAN ADABIYOTLAR RO‘YXATI':
            on = True; out = []; continue
        if on and p.text.strip() == 'ILOVALAR':
            break
        if on:
            out.append(p)
    return out


# 2) yangi yozuvlar
last = None
for anc, txt in N.NEW:
    if anc == 'NEXT':
        a = last
    else:
        a = [p for p in bib_paras() if anc in p.text][0]
    last = D.new_para_after(a, '0. ' + txt, tag='adab')
# qayta raqamlash
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
# 3) Kirish: manbalar soni
D.replace(D.para('foydalanilgan adabiyotlar ro‘yxati 197 ta manbani'), '197', str(n), tag='adab')
D.save(F)
print('manbalar:', n, pardoz(F))
