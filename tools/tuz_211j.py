"""2.1.1-jadval: I darajali kapital monandligi (K2) qatori qo‘shildi (qizil) – 3697-son nizom, 105-band."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
import kamlib as K
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
D = K.Doc(F)
t = [tb for tb in D.d.tables if any(r.cells[0].text.strip() == 'Leverej koeffitsiyenti' for r in tb.rows)][0]
ri = [i for i, r in enumerate(t.rows) if r.cells[0].text.strip().startswith('I darajali asosiy kapital (CET1) monandligi')][0]
D.add_row(t, ri, ['I darajali kapital monandligi', '≥ 10% (2025-yilgacha); ≥ 6% + buferlar (2026-yildan)', '≥ 6%',
                  '2026-yildan buferlar bilan 10%, tizimli ahamiyatga molik banklar uchun 11%'], tag='211J')
D.save(F)
print(pardoz(F))
