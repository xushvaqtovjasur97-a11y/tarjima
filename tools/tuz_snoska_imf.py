"""Roadmap snoskalarida tushib qolgan nuqta va qisqa “IMF Country Report No. 25/…, 2025” havolalarini to‘liq manbaga almashtirish (qizil)."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
import kamlib as K
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
FSSA = 'International Monetary Fund. Republic of Uzbekistan: Financial Sector Assessment Program – Financial System Stability Assessment. IMF Country Report No. 25/145. – Washington, DC: IMF, 2025'
BCP = 'International Monetary Fund. Republic of Uzbekistan: Financial Sector Assessment Program – Detailed Assessment of Observance – Basel Core Principles for Effective Banking Supervision. IMF Country Report No. 25/227. – Washington, DC: IMF, 2025'
D = K.Doc(F)
JOB = []
for fid in D.fn_ids('17.11.2025 – Item_'):
    JOB.append((fid, '17.11.2025 – Item_', '17.11.2025. – Item'))
for fid in D.fn_ids('IMF Country Report No.\xa025/145, 2025'):
    JOB.append((fid, 'IMF Country Report No.\xa025/145, 2025', FSSA))
for fid in D.fn_ids('IMF Country Report No.\xa025/227, 2025'):
    JOB.append((fid, 'IMF Country Report No.\xa025/227, 2025', BCP))
P = {fid: D.fn_para(fid=fid) for fid, _, _ in JOB}
for fid, old, new in JOB:
    D.replace(P[fid], old, new, tag='IMF')
print(JOB and [j[0] for j in JOB])
D.save(F)
print(pardoz(F))
