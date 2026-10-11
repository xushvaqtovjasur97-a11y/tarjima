# Yagona metodika: milliy hisobot varianti (SQB, MKB 2021-2025)
import json,csv,collections,sys
sys.path.insert(0,'/home/user/tarjima/manbalar/hisob_kitob')
import mbii as M
R='/home/user/tarjima/manbalar/'
P=json.load(open(R+'hisob_kitob/panel31.json'))
S=collections.defaultdict(dict)
for r in csv.DictReader(open(R+'markaziy_bank/banklar/banklar_2021IV_2026II.csv')):
  try: S[(r['bank'],r['davr'])][r['korsatkich']]=float(r['mlrd_som'])
  except: pass
SN={'SQ':"O'ZSANOATQURILISHBANK",'MK':'MIKROKREDITBANK'}
def ind(b,y):
  p=P[f'{b}{y}']; s=S[(SN[b],f'{y}_IV')]
  x={k:p[k] for k in ('CET1','CAR','S3','S2','NPLnet','LTD','ROA','ROE','CI','FXL','FXLI')}
  x['LEV']=p['_EQ']/p['_TA']*100
  x['HLA']=s['Высоколиквидные активы']/s['Всего активы']*100
  return x
if __name__=='__main__':
  for b in ('SQ','MK'):
    for y in range(2021,2026):
      x=ind(b,y); m,bs=M.mbii(x,M.W['taklif'])
      print(b,y,round(m,2),M.cls(m),{k:round(v,2) for k,v in bs.items()},{k:round(v,1) for k,v in x.items()})
