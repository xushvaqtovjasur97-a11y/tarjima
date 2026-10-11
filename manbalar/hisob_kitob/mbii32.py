# 3.2-§: MBII ochiq ma'lumotlar varianti (MHXS hisobotlari + SNS bazasidagi yuqori likvid aktivlar)
import json,csv,collections,sys,os
sys.path.insert(0,'/home/user/tarjima/manbalar/hisob_kitob')
import mbii as M
R='/home/user/tarjima/manbalar/'
SNS={'SQB':"O'ZSANOATQURILISHBANK",'MKB':'MIKROKREDITBANK','BRB':'BIZNESNI RIVOJLANTIRISH BANKI','HAMKOR':'HAMKORBANK','ALOQA':'ALOQABANK','TURON':'TURONBANK','XALQ':'XALQ BANKI','KAPITAL':'KAPITALBANK','NBU':"O'ZMILLIYBANK",'ASAKA':'ASAKABANK','IPOTEKA':'IPOTEKA-BANK'}
S=collections.defaultdict(dict)
for r in csv.DictReader(open(R+'markaziy_bank/banklar/banklar_2021IV_2026II.csv')):
  try: S[(r['bank'],r['davr'])][r['korsatkich']]=float(r['mlrd_som'])
  except: pass
IND=M.IND
W=M.W['taklif']
def ind(code):
  d=json.load(open(R+f'mhxs_banklar/{code}.json'))['yillar']; out={}
  for y in range(2021,2026):
    v=d[str(y)]; p=d[str(y-1)]; g=lambda k: v.get(k)
    rwa=g('RWA') or (g('REG_CAP')/g('CAR')*100 if g('REG_CAP') and g('CAR') else None)
    np_=g('NP'); eq=g('EQ'); nci=g('NCI') or 0
    x={}
    x['CET1']=(eq-nci-max(np_,0))/rwa*100 if rwa else None
    x['CAR']=g('CAR'); x['LEV']=eq/g('TA')*100
    st=g('S_TOTAL') or g('LOANS_GROSS')
    x['S3']=g('S3')/st*100 if g('S3') is not None else None
    x['S2']=g('S2')/st*100 if g('S2') is not None else None
    x['NPLnet']=(g('S3')-g('ECL'))/g('REG_CAP')*100 if (g('S3') is not None and g('ECL') and g('REG_CAP')) else None
    s=S.get((SNS[code],f'{y}_IV'),{})
    x['HLA']=s['Высоколиквидные активы']/s['Всего активы']*100 if s.get('Всего активы') else None
    x['LTD']=g('LOANS_GROSS')/g('DEP')*100 if g('DEP') else None
    x['ROA']=np_/((g('TA')+p['TA'])/2)*100; x['ROE']=np_/((eq+p['EQ'])/2)*100
    inc=sum(g(k) or 0 for k in ('NII','FEE_NET','FX_NET','OTHER_INC'))
    x['CI']=g('OPEX')/inc*100 if inc>0 and g('OPEX') else None
    x['FXL']=g('FX_LOANS_SHARE'); x['FXLI']=g('FX_LIAB_SHARE')
    out[y]=x
  return out
def mbii(x): return M.mbii(x,W)
if __name__=='__main__':
  res={}
  for code in sys.argv[1:]:
    I=ind(code); res[code]={}
    for y,x in I.items():
      m,bs=mbii(x); res[code][y]=dict(ind=x,bloklar=bs,mbii=m)
      print(code,y,round(m,2),M.cls(m),{b:round(s,2) for b,s in bs.items()},{k:(round(v,1) if v is not None else None) for k,v in x.items()})
  json.dump(res,open(R+'hisob_kitob/mbii32.json','w'),indent=1,ensure_ascii=False,default=float)
