"""Namuna davomi: (1) “Rossiya …” gapi mazmunan to‘ldiriladi (sahifa oxiri bo‘sh qolmasligi uchun);
(2) barcha jadval kataklarida oraliq (oldin/keyin) 0 pt."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
import kamlib as K
K.RED = '0070C0'
F = sys.argv[1]
D = K.Doc(F)
p = D.para('Rossiya iqtisodiy maktabining boshqa vakillari ham bu tushunchani turlicha talqin qiladi.')
D.replace(p, 'turlicha talqin qiladi.', 'turlicha talqin qiladi, biroq ularning ta’riflarida ham muvozanat, ishonchlilik '
                                         'va rivojlanish belgilari asosiy o‘rinni egallaydi, mualliflarning har biri esa ulardan birini ustuvor deb hisoblaydi.')
n = 0
for tbl in D.d.element.body.iter(qn('w:tbl')):
    for pe in tbl.iter(qn('w:p')):
        pf = Paragraph(pe, None).paragraph_format
        if pf.space_before != 0 or pf.space_after != 0:
            n += 1
        pf.space_before = 0; pf.space_after = 0
D.save(F)
print('kataklar:', n)
