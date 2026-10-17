"""“… tashkil etadi: …” kabi ikki nuqta bilan uzilgan mustaqil gaplarni mantiqan bog‘lash (137 joy)."""
import json, sys
sys.path.insert(0, '/home/user/tarjima/tools')
from docx.oxml.ns import qn
import kamlib as K
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
E = []
for g in ('g1', 'g2', 'g3'):
    E += json.load(open(f'/home/user/tarjima/tools/boglash_tahrirlar/colon_{g}_edits.json'))
D = K.Doc(F)
ps = D.d.paragraphs        # indekslar tahrirdan oldin olinadi (paragraflar soni o‘zgarmaydi)
n = 0
for e in E:
    p = ps[e['p']]
    t = ''.join(x.text or '' for x in p._p.iter(qn('w:t')))
    assert t.count(e['old']) == 1, (e['p'], e['old'])
    D.replace(p, e['old'], e['new'])
    n += 1
D.save(F)
print(n, pardoz(F))
