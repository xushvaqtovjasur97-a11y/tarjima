# 2.2.7-rasm: «Aloqabank» ATB qo'shilgan variant – alohida faylda ko'rish uchun
import sys, json
sys.path.insert(0, '/home/user/tarjima/tools')
import docx
from docx.shared import Pt, Cm
import wordart as W
A = json.load(open('/home/user/tarjima/manbalar/hisob_kitob/aloqa23.json'))
Y = ['2021', '2022', '2023', '2024', '2025']
LBLUE, DBLUE = '9DC3E6', '2E75B6'
d = docx.Document()
s = d.sections[0]; s.page_width, s.page_height = Cm(21), Cm(29.7)
s.left_margin = Cm(3); s.right_margin = Cm(1.5)
p = d.add_paragraph('{{FIG}}')
SER = [
    ('SQB: rasmiy muammoli', [3.8, 2.8, 2.2, 2.8, 2.5], '5B9BD5', 'dash'), ('SQB: 3-bosqich (MHXS)', [6.2, 7.4, 4.8, 5.9, 7.5], '1F4E79'),
    ('MKB: rasmiy muammoli', [5.9, 4.8, 6.0, 4.6, 4.4], '70AD47', 'dash'), ('MKB: 3-bosqich (MHXS)', [17.4, 13.5, 16.6, 15.2, 12.3], '375623'),
    ('Aloqabank: rasmiy muammoli', [round(A[y]['NPL_n'], 1) for y in Y], 'ED7D31', 'dash'), ('Aloqabank: 3-bosqich (MHXS)', [round(A[y]['S3_i'], 1) for y in Y], 'C00000')]
rid = W.add_chart(d, W.line_table_chart(Y, SER, ymin=0, ymax=20, major=4, layout=(0.30, 0.03, 0.68, 0.50)))
W.replace_placeholder(d, '{{FIG}}', W.chart_run_xml(rid, 16.5, 11.0, '2.2.7-rasm'))
c = d.add_paragraph('2.2.7-rasm. «O‘zsanoatqurilishbank» ATB, «Mikrokreditbank» ATB va «Aloqabank» ATB da MHXS 9 bo‘yicha 3-bosqich kreditlari va rasmiy muammoli kreditlar ulushi, %')
for r in c.runs:
    r.font.name = 'Times New Roman'; r.font.size = Pt(12); r.bold = True
c.alignment = 1
n = d.add_paragraph('Namuna (ko‘rish uchun): chiziqli diagramma va ostida ma’lumotlar jadvali; uzuq chiziq – rasmiy muammoli kreditlar, to‘liq chiziq – MHXS 9 bo‘yicha 3-bosqich kreditlari. Aloqabank ma’lumotlari: muallif hisob-kitobi (prudensial shakllar va MHXS hisobotlari).')
for r in n.runs:
    r.font.name = 'Times New Roman'; r.font.size = Pt(11); r.italic = True
d.save('/home/user/tarjima/Namuna_2.2.7-rasm_Aloqabank.docx'); print('saqlandi')
