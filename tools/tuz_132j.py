"""1.3.2-jadval (Rossiya banki tizimi evolyutsiyasi): 7302-U qatori va manbalar (qizil)."""
import re, sys
sys.path.insert(0, '/home/user/tarjima/tools')
from docx.oxml.ns import qn
import kamlib as K
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
U4336 = 'Указание Банка России от 03.04.2017 № 4336-У “Об оценке экономического положения банков”. – https://www.cbr.ru'
U7302 = 'Указание Банка России от 12.03.2026 № 7302-У “О внесении изменений в Указание Банка России от 3 апреля 2017 года № 4336-У “Об оценке экономического положения банков”” (зарегистрировано Минюстом России 16.04.2026 № 86081; вступает в силу с 01.10.2026)'
D = K.Doc(F)
cap = D.para('Rossiya bankining banklar moliyaviy barqarorligini baholash tizimi evolyutsiyasi')
fid = [r.get(qn('w:id')) for r in cap._p.iter(qn('w:footnoteReference'))][0]
t = D.table_after('Rossiya bankining banklar moliyaviy barqarorligini baholash tizimi evolyutsiyasi')
ri = [i for i, r in enumerate(t.rows) if r.cells[1].text.strip() == '7302-U ko‘rsatma'][0]
D.set_cell(t, ri, 2, '4336-U ko‘rsatmaga o‘zgartirishlar (2026-yil 1-oktyabrdan): 2- va 3-guruhlarga kiritish asoslari, kapital yetarliligi va sifati ko‘rsatkichlarini hisoblash tartibi aniqlashtirildi', tag='RU')
D.insert(D.fn_para(fid), 'С. 104–106', '; ' + U4336 + '; ' + U7302 + '; Национальный норматив краткосрочной ликвидности для СЗКО: порядок расчета: информация Банка России от 12 августа 2025 г.', tag='RU')
D.save(F)

D = K.Doc(F)
a = [p for p in D.paras() if re.match(r'^\d+\. ', p.text) and 'Национальный норматив краткосрочной ликвидности для СЗКО' in p.text][-1]
D.new_para_after(a, '0. ' + U7302 + '.', tag='RU')
D.new_para_after(a, '0. ' + U4336 + '.', tag='RU')
start = False; n = 0
for p in D.paras():
    if p.text.strip() == 'FOYDALANILGAN ADABIYOTLAR RO‘YXATI':
        start = True; continue
    if start and p.text.strip() == 'ILOVALAR':
        break
    if start:
        m = re.match(r'^(\d+)\. ', p.text)
        if m:
            n += 1
            r0 = D._runs(p)[0].find(qn('w:t'))
            if r0.text.startswith(m.group(1) + '.'):
                r0.text = str(n) + r0.text[len(m.group(1)):]
            elif r0.text == m.group(1):
                r0.text = str(n)
hp = D.para('Dissertatsiyaning tuzilishi va hajmi.')
D.replace(hp, re.search(r'ro‘yxati \d+ ta manbani', hp.text).group(0), f'ro‘yxati {n} ta manbani', tag='RU')
D.save(F)
print('adabiyotlar:', n, pardoz(F))
