"""1.2.3-rasm: o‘q nomi, legenda, pog‘ona foizlari, vergulli x qiymatlari; matnga kontrsiklik bufer izohi (qizil)."""
import re, sys
sys.path.insert(0, '/home/user/tarjima/tools')
from docx.oxml.ns import qn
import kamlib as K
import wordart as W
import namuna_123r as N
from pardoz import pardoz

F = N.F
D = K.Doc(F)
cap = [p for p in D.d.paragraphs if p.text.startswith(N.CAP)][0]
ch = cap._p.getprevious()
rid = [e.get(qn('r:id')) for e in ch.iter('{http://schemas.openxmlformats.org/drawingml/2006/chart}chart')][0]
part = D.d.part.related_parts[rid]
xrel = [r for r in part.rels.values() if r.reltype.endswith('/package')][0]
xml, rows = N.chart(part.blob.decode('utf-8'))
part._blob = re.sub(r'(<c:externalData r:id=")[^"]+(")', r'\g<1>' + xrel.rId + r'\g<2>', xml).encode('utf-8')
xrel.target_part._blob = W._xlsx(rows)
tp = [p for p in D.d.paragraphs if N.INS_AFTER in p.text and p.text.startswith('Rasmdan ko‘rinadiki, asosiy kapital 4,5')][0]
D.insert(tp, N.INS_AFTER, N.INS, tag='123r')
D.save(F)
print(part.partname, pardoz(F))
