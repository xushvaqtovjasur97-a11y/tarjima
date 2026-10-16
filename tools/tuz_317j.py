"""3.1.7-jadval aniqlashtirishlari, izoh va keyingi xatboshi (qizil)."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
import kamlib as K
import namuna_317j as N
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
CAP = 'Moliyaviy barqarorlikning integral indeksi va mavjud baholash metodikalarining qiyosiy tavsifi'
D = K.Doc(F)
t = D.table_after(CAP)
D.cell_replace(t, 0, 0, 'Belgi', 'Taqqoslash mezoni', tag='317j')
D.cell_replace(t, 3, 3, 'XVJ standarti', 'Bazel III ga yaqin ta’riflar (XVJ qo‘llanmasi, 2019)', tag='317j')
D.cell_replace(t, 4, 3, 'Milliy ta’rif', '90 kunlik mezon (amalda milliy ta’rif)', tag='317j')
D.cell_replace(t, 8, 4, 'Sinflar bo‘yicha', 'Sinflar va erta ogohlantirish zonalari bo‘yicha', tag='317j')
D.add_row(t, 7, ['Boshqaruv sifati', 'M komponenti', 'Boshqaruv sifati va shaffoflik guruhlari', 'Yo‘q',
                 'Yo‘q (ochiq ma’lumotlarda baholanmaydi)'], tag='317j')
D.add_row(t, 6, ['Vaznlar', 'Qat’iy belgilanmagan, nazoratchi mulohazasi', 'Har bir ko‘rsatkichga 1–3 vazn',
                 'Qo‘llanmaydi', 'Bloklar vaznlari 0,25/0,30/0,20/0,15/0,10, empirik tekshirilgan (3.2-§)'], tag='317j')
fn = D.fn_para(contains='Лаврушина, И.Д. Мамоновой. – М.: КНОРУС, 2011. – С. 91–92, 135–154; Cargill')
D.replace(fn, N.FN_OLD, N.FN_NEW, tag='317j')
D.replace(D.para(N.TXT_OLD), N.TXT_OLD, N.TXT_NEW, tag='317j')
D.save(F)
print(pardoz(F))
