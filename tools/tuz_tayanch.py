"""O‘z vaznlarimiz: “ekspert vaznlari” → “tayanch vaznlar”, “ekspert yo‘li bilan” → CAMEL nisbati asosida moslashtirish (qizil).
Adabiyotdagi “ekspert usuli/yo‘li” (Shatalova, Karminskiy, Rossiya metodikalari, SAABA va h.k.) o‘zgarmaydi."""
import io, re, sys
sys.path.insert(0, '/home/user/tarjima/tools')
import openpyxl
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn
import kamlib as K
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
MAXSUS = [
    ('chegaraviy qiymatlar va vaznlar qisman ekspert yo‘li bilan belgilangan',
     'chegaraviy qiymatlarning bir qismi teng qadam tamoyili bo‘yicha belgilangan, vaznlar esa CAMEL nisbati asosida moslashtirilgan'),
    ('Metodikaning ekspert yo‘li bilan belgilangan elementlari', 'Metodikaning empirik tekshiruv talab qiladigan elementlari'),
    ('ikkinchi bob tahlili asosida ekspert yo‘li bilan belgilandi', 'ikkinchi bob tahlili asosida CAMEL nisbatini moslashtirish yo‘li bilan belgilandi'),
    ('ekspert va empirik usullar bo‘yicha vaznlari', 'tayanch va empirik vaznlari'),
]
UMUMIY = [('ekspert vazn', 'tayanch vazn'), ('Ekspert vazn', 'Tayanch vazn')]

D = K.Doc(F)
n = 0
for pe in list(D.d.element.body.iter(qn('w:p'))):
    p = Paragraph(pe, D.d)
    for old, new in MAXSUS + UMUMIY:
        while old in p.text:
            D.replace(p, old, new, tag='tayanch'); n += 1
# diagrammalar (3.2.1, 3.2.2-rasmlar): kesh va ichki xlsx
ch = 0
for rel in D.d.part.rels.values():
    if not rel.reltype.endswith('/chart'):
        continue
    part = rel.target_part
    x = part.blob.decode('utf-8')
    if 'kspert' not in x:
        continue
    x2 = x.replace('Vaznlar ekspert vaznlari atrofida', 'Vaznlar tayanch vaznlar atrofida').replace('>Ekspert<', '>Tayanch<')
    x2 = x2.replace('ekspert vaznlari', 'tayanch vaznlar').replace('Ekspert vaznlari', 'Tayanch vaznlar')
    part._blob = x2.encode('utf-8'); ch += 1
    for r in part.rels.values():
        if r.reltype.endswith('/package'):
            wb = openpyxl.load_workbook(io.BytesIO(r.target_part.blob))
            for ws in wb.worksheets:
                for row in ws.iter_rows():
                    for c in row:
                        if isinstance(c.value, str) and 'kspert' in c.value:
                            c.value = (c.value.replace('Vaznlar ekspert vaznlari atrofida', 'Vaznlar tayanch vaznlar atrofida')
                                       .replace('Ekspert vaznlari', 'Tayanch vaznlar').replace('ekspert vaznlari', 'tayanch vaznlar')
                                       .replace('Ekspert', 'Tayanch'))
            b = io.BytesIO(); wb.save(b); r.target_part._blob = b.getvalue()
D.save(F)
print('matn:', n, 'diagramma:', ch, pardoz(F))
