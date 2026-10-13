# Dissertatsiya matnida bog'lanmagan gaplar, yarim qolgan jumlalar va tinish belgilarini tuzatish (qizil rangda)
import json, glob, sys
sys.path.insert(0, 'tools')
from kamlib import Doc, LOG
from pardoz import pardoz
from docx.text.paragraph import Paragraph

F = 'Dissertatsiya_Maxmudov_O.T_asosiy.docx'
S = sys.argv[1]
E = []
for f in sorted(glob.glob(S + 'edits_*.json')):
    E += json.load(open(f))

# muallif tomonidan qayta ko'rilgan tahrirlar
FIX = {
    'Biroq bu baho faqat sakkiz chorakni qamraydi': 'Shuni ta’kidlash joizki, YaIM o‘sishi va asosiy stavka bo‘yicha olingan baholar faqat sakkiz chorakni qamraydi',
    'ballni tashkil etib, uning asosiy sababi 2.3-paragrafda aniqlangan holat bo‘lib,': 'ballni tashkil etgan. Uning asosiy sababi 2.3-paragrafda aniqlangan holat bilan izohlanadi: ',
    'Beshinchi gipoteza esa muammoli': 'To‘rtinchi gipoteza asosiy faoliyatdan olinadigan barqaror daromad bank kapitalini ichki manbalar hisobiga to‘ldirish imkonini berishi bilan izohlanadi. Beshinchi gipoteza esa muammoli',
}
E = [e for e in E if not (e['p'] == 773 and e['new'].strip().endswith('– bo‘yicha'))]
E += [{'p': 773, 'old': 'besh blok – kapital', 'new': 'besh blok (kapital'},
      {'p': 773, 'old': 'valyuta riskiga sezgirlik bo‘yicha', 'new': 'valyuta riskiga sezgirlik) bo‘yicha'}]
for e in E:
    if e['old'] in FIX:
        e['new'] = FIX[e['old']]

D = Doc(F)
els = list(D.d.element.body.iterchildren())
ok = 0
for e in E:
    p = Paragraph(els[e['p']], D.d)
    txt = p.text
    old, new = e['old'], e['new']
    pos = txt.find(old)
    assert pos >= 0 and txt.count(old) == 1, (e["p"], old, txt[200:300])
    # umumiy boshi va oxirini ajratib, faqat o'zgargan qismni qizil qilamiz
    a = 0
    while a < min(len(old), len(new)) and old[a] == new[a]:
        a += 1
    b = 0
    while b < min(len(old), len(new)) - a and old[-1 - b] == new[-1 - b]:
        b += 1
    mo, mn = old[a:len(old) - b], new[a:len(new) - b]
    start = pos + a
    if mo == '':
        if a == 0:
            mo, mn = old[:1], mn + old[:1]
        else:
            after = txt[:start]
            occ = 0; i = after.find(txt[start - min(start, 30):start])
            anchor = txt[max(0, start - 30):start]
            occ = txt[:start].count(anchor) - 1
            D.insert(p, anchor, mn, occ=occ, tag=f"P{e['p']}")
            ok += 1
            continue
    occ = 0; i = txt.find(mo)
    while i != start:
        occ += 1; i = txt.find(mo, i + 1)
        assert i >= 0
    D.replace(p, mo, mn, occ=occ, tag=f"P{e['p']}")
    ok += 1
D.save(F)
pardoz(F)
print('kiritildi:', ok, 'jami:', len(E))
