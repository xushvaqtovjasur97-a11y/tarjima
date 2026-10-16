"""3.2.1-jadval: Shannon, bloklarga jamlash, PCA va logit vaznlari tavsifi (qizil)."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
import kamlib as K
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
CAP = 'Integral indeks vaznlarini aniqlashning empirik usullari'
FN = ('Tadqiqot natijasida muallif tomonidan quyidagi manbalar asosida ishlab chiqildi: OECD, JRC. Handbook on Constructing '
      'Composite Indicators: Methodology and User Guide. – Paris: OECD Publishing, 2008; Shannon C.E. A Mathematical Theory '
      'of Communication // Bell System Technical Journal. – 1948. – Vol. 27, No. 3. – P. 379–423; Saisana M., Saltelli A., '
      'Tarantola S. Uncertainty and sensitivity analysis techniques as tools for the quality assessment of composite '
      'indicators // Journal of the Royal Statistical Society: Series A. – 2005. – Vol. 168, No. 2. – P. 307–323.')
D = K.Doc(F)
cap = [p for p in D.d.paragraphs if p.text.strip() == CAP][0]
t = D.table_after(CAP)
D.cell_replace(t, 1, 1, 'K. Shannon', 'C. Shannon', tag='321a')
D.cell_replace(t, 1, 1, 'o‘lchanadi', 'o‘lchanadi; ko‘rsatkichlar vaznlari bloklar bo‘yicha jamlanadi', tag='321a')
D.cell_replace(t, 2, 1, 'o‘rtacha qiymatga nisbati', 'o‘rtacha qiymatga nisbati; ko‘rsatkichlar vaznlari bloklar bo‘yicha jamlanadi', tag='321a')
D.cell_replace(t, 3, 1, 'umumiy omillar ajratiladi',
               'umumiy omillar ajratiladi; blok vazni o‘z qiymati birdan katta komponentlar (Kayzer mezoni) yuklamalari '
               'kvadratlarining izohlangan dispersiya ulushi bilan tortilgan yig‘indisi sifatida aniqlanadi', tag='321a')
D.cell_replace(t, 4, 2, 'blok vazni – 0)',
               'blok vazni – 0; kapital va valyuta bloklari uchun tayanch vaznlar saqlanadi)', tag='321a')
D.save(F)
print(pardoz(F))
for r in t.rows: print(' | '.join(c.text for c in r.cells))
