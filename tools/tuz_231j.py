"""2.3.1-jadval: aniqlashtirishlar va restrukturizatsiya qatori (qizil)."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
from docx.oxml.ns import qn
import kamlib as K
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
tit = [p for p in D.d.paragraphs if p.text.startswith('Milliy prudensial hisobot va MHXS bo‘yicha hisobot o‘rtasidagi asosiy farqlar')][0]
fid = [r.get(qn('w:id')) for r in tit._p.iter(qn('w:footnoteReference'))][0]
t = [tb for tb in D.d.tables if tb.rows[0].cells[0].text.strip() == 'Belgi' and 'MHXS bo‘yicha hisobot' in tb.rows[0].cells[2].text][0]
R = {r.cells[0].text.strip(): i for i, r in enumerate(t.rows)}
fn = D.fn_para(fid=fid)
del tit

D.cell_replace(t, R['Hisobot doirasi'], 2, 'Bank guruhi (sho‘ba korxonalar bilan birga)',
               'Bank guruhi (sho‘ba korxonalari bo‘lsa – konsolidatsiyalangan), aks holda alohida bank', tag='231J')
D.cell_replace(t, R['Kreditlarni tasniflash asosi'], 2, 'Uch bosqich: 1-bosqich,', 'Uch bosqich: kredit riski sezilarli oshmagan 1-bosqich,', tag='231J')
D.cell_replace(t, R['Regulyativ kapital'], 2, 'Hisoblanmaydi, faqat oshkor qilinadi',
               'Hisoblanmaydi; izohlarda milliy me’yorlar bo‘yicha hisoblangan qiymat oshkor qilinadi', tag='231J')
D.cell_replace(t, R['Hisobot muddati va oshkoralik'], 1, 'Har oylik, choraklik; asosiy shakllar e’lon qilinadi',
               'Har oylik, choraklik; prudensial shakllar ochiq e’lon qilinmaydi, ko‘rsatkichlar Markaziy bank statistikasida jamlangan holda beriladi', tag='231J')
D.cell_replace(t, R['Hisobot muddati va oshkoralik'], 2, 'Yillik, auditorlik xulosasi bilan;', 'Yillik, auditorlik xulosasi bilan, olti oylik kechikish bilan;', tag='231J')
D.add_row(t, R['Zaxira shakllantirish'], [
    'Restrukturizatsiya qilingan kreditlar',
    '6 oygacha imtiyozli davr tasnifga ta’sir qilmaydi; yashovchanlik bahosi va muammoli toifadan chiqarish uchun sinov muddati yo‘q',
    'Kredit riski sezilarli oshganining belgisi sifatida 2 yoki 3-bosqichga o‘tkaziladi; yuqori bosqichga qaytarish uchun to‘lovlarning barqarorligi talab qilinadi',
    'Restrukturizatsiya qilingan kreditlarning asosiy qismi milliy hisobotda “standart” toifada qoladi'], tag='231J')
D.insert(fn, '', '; International Monetary Fund. Republic of Uzbekistan: Financial Sector Assessment Program – Financial System Stability Assessment. IMF Country Report No. 25/145. – Washington, DC: IMF, 2025. – P. 21, 30.', tag='231J')
D.save(F)
print(pardoz(F))
