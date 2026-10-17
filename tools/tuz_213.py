"""2.1.3-rasm snoskasi standart ifodaga; 2.1.4-rasm “Regulyativ minimum” – 2026-yildan buferlar bilan (8,5 va 12)."""
import re, sys
sys.path.insert(0, '/home/user/tarjima/tools')
from lxml import etree
from docx.oxml.ns import qn
import kamlib as K
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
cap3 = D.para('2.1.3-rasm. O‘zbekiston bank tizimida kapital monandligi ko‘rsatkichlari dinamikasi')
cap4 = D.para('2.1.4-rasm. Salbiy stress-ssenariyda bank tizimi kapital ko‘rsatkichlarining o‘zgarishi')
fid = [r.get(qn('w:id')) for r in cap3._p.iter(qn('w:footnoteReference'))][0]
D.replace(D.fn_para(fid), 'Muallif tomonidan quyidagi manbalar asosida tuzilgan:', 'Tadqiqot natijasida muallif tomonidan quyidagi manbalar asosida ishlab chiqildi:', tag='213')
rid = [c.get(qn('r:id')) for c in cap4._p.getprevious().iter('{http://schemas.openxmlformats.org/drawingml/2006/chart}chart')][0]
name = str(D.d.part.related_parts[rid].partname)
x, part = D.chart_xml(name)
ns = {'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'}
sers = x.findall('.//c:ser', ns)
idx = [i for i, s in enumerate(sers) if s.find('./c:tx//c:v', ns).text == 'Regulyativ minimum'][0]
sers[idx].find('./c:tx//c:v', ns).text = 'Regulyativ minimum (2026-yildan, buferlar bilan)'
part._blob = etree.tostring(x, xml_declaration=True, encoding='UTF-8', standalone=True)
D.chart_set(name, idx, [8.5, 12], tag='214')
D.chart_sync_xlsx(name)
D.save(F)
print(name, pardoz(F))
