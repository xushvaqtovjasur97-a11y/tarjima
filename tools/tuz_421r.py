"""4.2.1-rasm: haqiqiy qiymatlar va 2027–2028-yil prognoz qiymatlari ko‘rinadigan yorliqlar (tools/namuna_421r.py dagi fix())."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
import namuna_421r as N
from pardoz import pardoz

D, cap, chp, part = N.source()
xml = part.blob.decode('utf-8')
assert '<c:max val="2028"/>' in xml, 'allaqachon tuzatilgan'
part._blob = N.fix(xml).encode('utf-8')
D.save(N.F)
print(pardoz(N.F))
