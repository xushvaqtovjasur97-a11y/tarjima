# 4.2-§: tizim darajasidagi choraklik qator 2020_I–2025_IV (2020-yil faqat modelni baholash uchun)
import openpyxl, json, re, numpy as np, pandas as pd
B='/home/user/tarjima/manbalar/markaziy_bank/'
def sheet(f,s): return [list(r) for r in openpyxl.load_workbook(B+f,read_only=True,data_only=True)[s].iter_rows(values_only=True)]
out={}
# NPL ulushi (oy oxiri): 2020 – 5.1.8а, 2021 – 5.1.8б (bul2021); 2022–2025 – makro.json (bul2603)
for f,s,y in (('bul2021.xlsx','5.1.8а',2020),('bul2021.xlsx','5.1.8б',2021)):
  R=sheet(f,s); r=next(x for x in R if x[0] and str(x[0]).startswith('Muammoli kreditlar va mikroqarzlarning jami'))
  for k,q in ((3,'I'),(6,'II'),(9,'III'),(12,'IV')): out[f'{y}_{q}']={'NPL':r[k]}
MK=json.load(open('makro.json'))
for k,v in MK.items(): out.setdefault(k,{})['NPL']=v['NPL_tizim']
# YaIM real o'sishi (yil boshidan, %) – bul2603 1.1
R=sheet('bul2603.xlsx','1.1'); hdr=R[4]; real=next(R[i+1] for i,r in enumerate(R) if r[1] and str(r[1]).startswith('YaIM (nominal'))
for c,x in enumerate(hdr):
  if isinstance(x,str) and re.match(r'20\d\d-yil',x) and int(x[:4])<=2025:
    for k,q in enumerate(('I','II','III','IV')):
      out.setdefault(f'{x[:4]}_{q}',{})['YaIM']=real[c+k]-100
D=pd.DataFrame(out).T; D=D.loc[[f'{y}_{q}' for y in range(2020,2026) for q in ('I','II','III','IV')]].astype(float)
print(D.round(2).T.to_string())
D.to_csv('tizim42.csv')
