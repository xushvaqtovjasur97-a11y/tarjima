"""3.2.1-rasm (Monte-Karlo): vergulli yorliqlar, qisqartmalar izohi, umumiy izoh; Dirixle parametri va sezgir banklar (qizil)."""
import re, sys
sys.path.insert(0, '/home/user/tarjima/tools')
from docx.oxml.ns import qn
import kamlib as K
import wordart as W
import namuna_321mc as N
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
cap = [p for p in D.d.paragraphs if p.text.startswith('3.2.1-rasm. Monte-Karlo')][0]
ch = cap._p.getprevious()
rid = [e.get(qn('r:id')) for e in ch.iter('{http://schemas.openxmlformats.org/drawingml/2006/chart}chart')][0]
part = D.d.part.related_parts[rid]
xrel = [r for r in part.rels.values() if r.reltype.endswith('/package')][0]
xml, rows = N.chart()
part._blob = re.sub(r'(<c:externalData r:id=")[^"]+(")', r'\g<1>' + xrel.rId + r'\g<2>', xml).encode('utf-8')
xrel.target_part._blob = W._xlsx(rows)
# qisqartmalar izohi (rasm va nomi orasida)
izsrc = D.para('Izoh: “Aloqabank” ATB empirik vaznlarni baholashda qatnashmagan')
iz = D.new_para_after(izsrc, N.IZOH, tag='mk')
ch.addnext(iz._p)
# umumiy izoh
fid = [r.get(qn('w:id')) for r in cap._p.iter(qn('w:footnoteReference'))][0]
fp = D.fn_para(fid=fid)
old = ''.join(x.text or '' for x in fp._p.iter(qn('w:t')))
pre = 'Tadqiqot natijasida muallif tomonidan '
i = old.index(pre) + len(pre)
D.replace(fp, old[i:], N.FN_NEW[len(pre):], tag='mk')
# matn
D.insert(D.para(N.ANC1), N.ANC1, N.INS1, tag='mk')
D.insert(D.para(N.ANC2), N.ANC2, N.INS2, tag='mk')
D.save(F)
print(pardoz(F))
