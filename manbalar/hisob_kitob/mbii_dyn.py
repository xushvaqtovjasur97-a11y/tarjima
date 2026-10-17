# MBII dinamikasi: milliy (MK 2019-2025, SQ 2020-2025) va MHXS (2021-2025)
import json,sys
sys.path.insert(0,'/home/user/tarjima/manbalar/hisob_kitob')
import mbii as M
P=json.load(open('/home/user/tarjima/manbalar/hisob_kitob/panel31.json'))
K=['CET1','CAR','LEV','S3','S2','NPLnet','LCR','NSFR','LTD','ROA','ROE','CI','FXL','FXLI']
NAT={k:{i:P[k][i] for i in K} for k in P}
Y=[2021,2022,2023,2024,2025]
IF={
'SQ':dict(CET1=[11.89,11.62,10.72,10.04,8.87],CAR=[15.14,14.96,15.73,14.77,15.32],T1=[11.91,11.63,10.74,10.05,12.66],
  S3=[6.24,7.38,4.81,5.86,7.48],S2=[20.4,14.7,15.8,27.1,35.66],NPLnet=[10.1,10.5,4.6,6.6,7.3],
  ROA=[1.6,1.1,1.3,1.4,1.70],ROE=[13.2,8.7,10.7,12.5,15.7],LTD=[328.4,334.7,421.6,330.6,221.0]),
'MK':dict(CET1=[13.59,20.17,16.20,15.02,18.24],CAR=[13.86,20.80,17.01,17.44,19.96],T1=[13.59,20.17,16.20,15.02,18.24],
  S3=[17.44,13.49,16.58,15.16,12.34],S2=[26.7,29.7,26.2,25.9,25.3],NPLnet=[35.1,23.5,36.7,34.3,26.4],
  ROA=[0.04,0.07,-2.71,-4.57,0.15],ROE=[0.26,0.35,-13.0,-25.0,0.81],LTD=[407.3,268.9,251.0,255.4,324.4]),
}
OUT={}
for b in ('SQ','MK'):
  for i,y in enumerate(Y):
    n=NAT[f'{b}{y}']; v=dict(n)
    for k in ('CET1','CAR','S3','S2','NPLnet','ROA','ROE','LTD'): v[k]=IF[b][k][i]
    t1n=n['CET1'] if b=='MK' else {2021:11.88,2022:12.10,2023:11.04,2024:10.63,2025:14.01}[y]
    v['LEV']=n['LEV']*IF[b]['T1'][i]/t1n
    if b=='SQ' and y==2025: v.update(LCR=154.0,NSFR=124.8,CI=41.9)
    if b=='MK' and y==2025: v.update(CI=88.0)
    OUT[f'{b}{y}']=dict(milliy=n,mhxs=v)
res={}
for key in sorted(k for k in NAT if int(k[2:])>=2021):
  m,bs=M.mbii(NAT[key],M.W['taklif']); r=dict(milliy=round(m,2),milliy_bloklar={k:round(x,2) for k,x in bs.items()})
  if key in OUT:
    m2,bs2=M.mbii(OUT[key]['mhxs'],M.W['taklif']); r.update(mhxs=round(m2,2),mhxs_bloklar={k:round(x,2) for k,x in bs2.items()},delta=round(m2-m,2))
  res[key]=r
  print(key,r.get('milliy'),M.cls(r['milliy']),'|',r.get('mhxs'),M.cls(r['mhxs']) if 'mhxs' in r else '', r.get('delta',''), r['milliy_bloklar'], r.get('mhxs_bloklar',''))
json.dump(res,open('/home/user/tarjima/manbalar/hisob_kitob/mbii_dyn.json','w'),indent=1,ensure_ascii=False)
