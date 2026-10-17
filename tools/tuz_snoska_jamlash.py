"""Ko‘p manbali izohlar → “Tadqiqotlar jarayonida muallif tomonidan tuzilgan.” (manbalar adabiyotlar ro‘yxatida);
bir manbaning ikki tildagi nusxasidan o‘zbekchasi qoldiriladi; adabiyotlar qayta raqamlanadi (ko‘k rangda)."""
import json, re, sys
sys.path.insert(0, '/home/user/tarjima/tools')
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
import kamlib as K
from pardoz import pardoz

K.RED = '0070C0'
F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
P = json.load(open('/home/user/tarjima/tools/snoska_reja/fn_plan.json'))
D = K.Doc(F)
FNS = {f.get(qn('w:id')): f for f in D.fx.findall(qn('w:footnote'))}
NEW = 'Tadqiqotlar jarayonida muallif tomonidan tuzilgan.'


def ftext(f):
    return ''.join(t.text or '' for t in f.iter(qn('w:t')))


n_rep = 0
for e in P['fn_replace']:
    f = FNS[e['id']]
    ps = f.findall(qn('w:p'))
    assert len(ps) == 1, e['id']
    old = ftext(f).strip()
    assert old == e['old_full'].strip(), e['id']
    new = NEW
    m = re.search(r'\s(Izoh: .*)$', old)          # hisob usuli izohi saqlanadi
    if m:
        new = NEW + ' ' + m.group(1)
    D.replace(Paragraph(ps[0], None), old, new)
    n_rep += 1
n_lang = 0
for e in P['fn_lang']:
    f = FNS[e['id']]
    pe = [p for p in f.findall(qn('w:p')) if e['old'] in ''.join(t.text or '' for t in p.iter(qn('w:t')))]
    assert len(pe) == 1, e['id']
    D.replace(Paragraph(pe[0], None), e['old'], e['new'])
    n_lang += 1


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


def body_of(p):
    return re.sub(r'^\d+\.\s*', '', p.text).strip()


for a in P['bib_add']:
    anc = [p for p in bib_paras() if body_of(p) == a['after'].strip()]
    assert len(anc) == 1, a['after'][:50]
    D.new_para_after(anc[0], '0. ' + a['text'])
for txt in P['bib_delete']:
    tg = [p for p in bib_paras() if body_of(p) == txt.strip()]
    assert len(tg) == 1, txt[:50]
    tg[0]._p.getparent().remove(tg[0]._p)
# 2023-yilgi sharhning inglizcha nusxasi → o‘zbekcha
fsr23 = [p for p in bib_paras() if body_of(p) == 'The Central Bank of the Republic of Uzbekistan. Financial Stability Report for 2023. – Tashkent, 2024.']
assert len(fsr23) == 1
D.replace(fsr23[0], 'The Central Bank of the Republic of Uzbekistan. Financial Stability Report for 2023. – Tashkent, 2024.',
          'O‘zbekiston Respublikasi Markaziy banki. Moliyaviy barqarorlik sharhi: 2023-yil. – Toshkent, 2024.')
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
kp = D.para('foydalanilgan adabiyotlar ro‘yxati 209 ta manbani')
D.replace(kp, '209', str(n))
D.save(F)
print('izohlar:', n_rep, 'til:', n_lang, 'manbalar:', n, pardoz(F))
