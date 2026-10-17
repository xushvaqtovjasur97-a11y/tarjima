"""(2.2) formula: LMHXS ta’rifi aniqlashtirildi (bosqichlarga taqsimlangan kreditlar), qizil."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
import kamlib as K
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
p = [q for q in D.d.paragraphs if q.text.startswith('bu yerda S3')][0]
assert p.text.index(' va ') > p.text.index('kreditlar, L')
D.replace(p, ' va ', ' – MHXS 9 bo‘yicha bosqichlarga taqsimlangan, ya’ni amortizatsiyalangan qiymatda baholanadigan kreditlar (brutto), ', tag='lmhxs')
D.replace(p, ' – ikki hisobot bo‘yicha jami kreditlar (brutto).',
          ' – milliy hisobot bo‘yicha jami kreditlar (brutto). Foyda yoki zarar orqali adolatli qiymatda baholanadigan kreditlar '
          'bosqichlarga ajratilmagani sababli maxraj tarkibiga kiritilmaydi: masalan, “Mikrokreditbank” ATB da bunday kreditlar '
          '2021-yilda 2 729,6 mlrd so‘mni, 2025-yilda 1 362,7 mlrd so‘mni tashkil etgan.', tag='lmhxs')
D.save(F)
print(pardoz(F))
