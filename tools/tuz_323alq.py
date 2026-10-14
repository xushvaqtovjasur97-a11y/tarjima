"""3.2.3-jadvalga “Aloqabank” (tanlanmadan tashqari) qatori, izoh, tahlil; 3,248 talqini (qizil)."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
import kamlib as K
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
CAP = 'Beshta bankning integral indeksi turli usullar vaznlari bo‘yicha, 2025-yil'
t = D.table_after(CAP)
ri = [i for i, r in enumerate(t.rows) if r.cells[0].text.startswith('“Ipoteka-bank”')][0]
D.add_row(t, ri, ['“Aloqabank” ATB (tanlanmadan tashqari)', '2,78', '2,23', '2,41', '2,49', '3,01', '2,38'], tag='alq')
izsrc = D.para('Izoh: * – logit usulida baholanmagan, tayanch vazni saqlangan.')
iz = D.new_para_after(izsrc, 'Izoh: “Aloqabank” ATB empirik vaznlarni baholashda qatnashmagan va tanlanmadan tashqari tekshiruv '
                      'uchun keltirilgan; jadvalning quyi uch qatoridagi ko‘rsatkichlar beshta bank bo‘yicha hisoblangan.', tag='alq')
t._tbl.addnext(iz._p)
p = D.para('Jadvaldan ko‘rinib turibdiki, vaznlar tanlovi indeksning mutlaq qiymatiga')
D.insert(p, 'ammo uning ichida quyiroqqa tushadi.',
         ' Tanlanmadan tashqari “Aloqabank” ATB da esa manzara boshqacha: tayanch vaznlarda bank II sinfda (2,78 ball) bo‘lsa, '
         'tarqoqlikka asoslangan uch usul vaznlarida III sinfga (2,23–2,49 ball) tushadi. Buning sababi shundaki, bu usullar '
         'daromadlilik blokiga yuqori vazn beradi, bank esa 2025-yilda MHXS bo‘yicha zarar ko‘rgan. Bu holat empirik vaznlar '
         'ular baholangan tanlanmaning xususiyatlariga bog‘liqligini amalda tasdiqlaydi.', tag='alq')
q = D.para('indeks 3,248 ballni, ya’ni I sinf chegarasini tashkil etgan')
D.replace(q, 'indeks 3,248 ballni, ya’ni I sinf chegarasini tashkil etgan', 'indeks 3,248 ballni tashkil etib, I sinf chegarasiga yaqinlashgan', tag='alq')
D.save(F)
print(pardoz(F))
# --- jadval nomi: oltinchi bank qo‘shilgani sababli (alohida ishga tushirildi)
