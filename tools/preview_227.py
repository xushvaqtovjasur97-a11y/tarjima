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
rid = W.add_chart(d, W.bar_chart(Y, [
    ('SQB: rasmiy muammoli kreditlar', [3.8, 2.8, 2.2, 2.8, 2.5], LBLUE), ('SQB: MHXS 9 bo‘yicha 3-bosqich', [6.2, 7.4, 4.8, 5.9, 7.5], DBLUE),
    ('MKB: rasmiy muammoli kreditlar', [5.9, 4.8, 6.0, 4.6, 4.4], 'A9D18E'), ('MKB: MHXS 9 bo‘yicha 3-bosqich', [17.4, 13.5, 16.6, 15.2, 12.3], '38761D'),
    ('Aloqabank: rasmiy muammoli kreditlar', [round(A[y]['NPL_n'], 1) for y in Y], 'F4B183'), ('Aloqabank: MHXS 9 bo‘yicha 3-bosqich', [round(A[y]['S3_i'], 1) for y in Y], 'C55A11')],
    ytitle='Jami kreditlarga nisbatan, %', ymax=20, major=4, label_fmt='0.0', gap=40))
W.replace_placeholder(d, '{{FIG}}', W.chart_run_xml(rid, 16.0, 8.0, '2.2.7-rasm'))
c = d.add_paragraph('2.2.7-rasm. «O‘zsanoatqurilishbank» ATB, «Mikrokreditbank» ATB va «Aloqabank» ATB da MHXS 9 bo‘yicha 3-bosqich kreditlari va rasmiy muammoli kreditlar ulushi, %')
for r in c.runs:
    r.font.name = 'Times New Roman'; r.font.size = Pt(12); r.bold = True
c.alignment = 1
n = d.add_paragraph('Namuna (ko‘rish uchun): 16×8 sm o‘lchamda 6 ta ustun – har bir yil uchun uch bankning ikki ko‘rsatkichi. Aloqabank ma’lumotlari: muallif hisob-kitobi (prudensial shakllar va MHXS hisobotlari).')
for r in n.runs:
    r.font.name = 'Times New Roman'; r.font.size = Pt(11); r.italic = True
d.save('/home/user/tarjima/Namuna_2.2.7-rasm_Aloqabank.docx'); print('saqlandi')
