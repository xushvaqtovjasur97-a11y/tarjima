"""3.1.1-rasm: och fon + to‘q matn, 6-qutichada blok harflari (chizma odatiy ranglarda)."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
import docx
import namuna_311r as N
import namuna_311r_rang as R
from pardoz import pardoz

F = N.F
d = docx.Document(F)
cap = [p for p in d.paragraphs if p.text.startswith('3.1.1-rasm.')][0]
fig = cap._p.getprevious()
print('6-quticha:', N.tuzat(fig), 'shakllar:', R.rangla(fig))
d.save(F)
print(pardoz(F))
