"""“O‘sha joyda”, “Ko‘rsatilgan asar” va qisqartirilgan “Roadmap”… snoskalarini to‘liq manbaga almashtirish (qizil)."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
import kamlib as K
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
RM = 'The Central Bank of the Republic of Uzbekistan. “Roadmap” aimed at further improving the bank regulation and supervision system in 2025–2028. – Tashkent, 17.11.2025.'
BER = 'Ўзбекистон Республикаси тижорат банкларининг ликвидлилиги ва тўлов қобилиятини таъминлаш масалалари: монография. – Т.: BMA, 2020.'
KAR = 'Tijorat banklari faoliyatini tartibga solish va nazorati: darslik. – T.: Iqtisod-Moliya, 2019.'
D = K.Doc(F)
JOB = {252: ('Ko‘rsatilgan asar.', BER), 253: ('Ko‘rsatilgan asar.', KAR), 255: ('Ko‘rsatilgan asar.', KAR),
       272: ('O‘sha joyda.', RM), 277: ('O‘sha joyda.', RM)}
for fid in D.fn_ids('“Roadmap”…'):
    JOB[int(fid)] = ('The Central Bank of the Republic of Uzbekistan. “Roadmap”…', RM[:-1])
P = {fid: D.fn_para(fid=fid) for fid in JOB}
for fid, (old, new) in sorted(JOB.items()):
    D.replace(P[fid], old, new, tag='IBID')
print(sorted(JOB))
D.save(F)
print(pardoz(F))
