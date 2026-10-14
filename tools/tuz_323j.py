"""3.2.3-jadval: sarlavha, umumiy izoh, AUC oralig‘i, 'deyarli', logit izohi (qizil)."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
from docx.oxml.ns import qn
import kamlib as K
import namuna_323j as N
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
CAP = 'Beshta bankning integral indeksi turli usullar vaznlari bo‘yicha, 2025-yil'
t = D.table_after(CAP)
D.replace(t.rows[0].cells[0].paragraphs[0], 'Ko‘rsatkich', 'Bank / ko‘rsatkich', tag='323j')
cap = [p for p in D.d.paragraphs if p.text.strip() == CAP][0]
fid = [r.get(qn('w:id')) for r in cap._p.iter(qn('w:footnoteReference'))][0]
fp = D.fn_para(fid=fid)
old = ''.join(x.text or '' for x in fp._p.iter(qn('w:t')))
pre = 'Tadqiqot natijasida muallif tomonidan '
i = old.index(pre) + len(pre)
D.replace(fp, old[i:], N.FN_NEW[len(pre):], tag='323j')
p = D.para('Muammo hodisalarini bashorat qilish kuchi bo‘yicha empirik vaznlar')
D.replace(p, 'AUC 0,864–0,907)', 'AUC 0,865–0,907, tayanch vaznlarda – 0,864)', tag='323j')
D.insert(p, 'bosh komponentlar usuli bilan ', 'deyarli ', tag='323j')
D.insert(p, N.ANC, N.INS, tag='323j')
D.save(F)
print(pardoz(F))
