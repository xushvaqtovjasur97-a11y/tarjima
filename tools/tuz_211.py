"""2.1.1-rasm: chizma (oddiy rang), snoska va matn (qizil)."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
from docx.oxml.ns import qn
import kamlib as K
import namuna_211 as N
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
cap = D.para('2.1.1-rasm. O‘zbekistonda banklar moliyaviy barqarorligini baholashning me’yoriy-huquqiy asoslari')
pt = D.para('Ikkinchi daraja – prudensial talablar: kapital monandligi, likvidlik, aktivlarni tasniflash va zaxiralar to‘g‘risidagi nizomlar.')
fid = [r.get(qn('w:id')) for r in cap._p.iter(qn('w:footnoteReference'))][0]
print('chizma:', N.fix(cap._p.getprevious()))
fp = D.fn_para(fid)
old = ''.join(x.text or '' for x in fp._p.iter(qn('w:t'))).strip()
D.replace(fp, old, N.NEW_FN, tag='211')
D.replace(pt, 'kapital monandligi, likvidlik, aktivlarni tasniflash va zaxiralar to‘g‘risidagi nizomlar.',
          'kapital monandligi (2025-yil 28-oktabrdagi 3697-son), likvidlikni boshqarish (2709-son; sof barqaror moliyalashtirish me’yori bo‘yicha talab 2025-yil 23-maydan kuchini yo‘qotgan va “yo‘l xaritasi”ga ko‘ra 2026-yildan monitoring qilinadi) hamda aktivlarni tasniflash va zaxiralar (2015-yil 14-iyuldagi 2696-son) to‘g‘risidagi nizomlar.', tag='211')
D.save(F)
print(pardoz(F))
print(K.Doc(F).para('Ikkinchi daraja – prudensial talablar').text[:600])
