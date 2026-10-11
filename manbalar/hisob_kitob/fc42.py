# 4.2-§: 2026–2028-yillar ssenariy prognozi (faqat 2021–2025 ma'lumotlari asosida)
import pandas as pd, numpy as np, json, warnings; warnings.filterwarnings('ignore')
from linearmodels.panel import PanelOLS
import mbii as M
p=pd.read_csv('panel33.csv'); d=p[p.tanlov].copy().sort_values(['bank','t'])
def w(x,a=.01): lo,hi=x.quantile([a,1-a]); return x.clip(lo,hi)
for c in ['S3','S3_l1','ROAq']: d[c]=w(d[c])
OUT={}
def est(dd):
  q=dd.set_index(['bank','t'])[['S3','S3_l1','YaIM_osish']].dropna()
  r1=PanelOLS(q.S3,q[['S3_l1','YaIM_osish']].assign(c=1),entity_effects=True).fit(cov_type='clustered',cluster_entity=True)
  q=dd.set_index(['bank','t'])[['ROAq','S3']].dropna()
  r2=PanelOLS(q.ROAq,q[['S3']].assign(c=1),entity_effects=True).fit(cov_type='clustered',cluster_entity=True)
  return r1,r2
r1,r2=est(d)
for nm,r in (('NPL',r1),('ROA',r2)):
  OUT['model_'+nm]={'b':r.params.round(4).to_dict(),'se':r.std_errors.round(4).to_dict(),'p':r.pvalues.round(4).to_dict(),'N':int(r.nobs),'R2w':round(float(r.rsquared_within),3)}
  print(nm,OUT['model_'+nm])
RHO=r1.params.S3_l1; BG=r1.params.YaIM_osish; BR=r2.params.S3; BRG=0.0
print('uzoq muddatli: 1 p.p. YaIM -> NPL',round(BG/(1-RHO),2))
# ssenariylar (yillik, har chorakka bir xil)
MK=json.load(open('makro.json')); g25=MK['2025_IV']['YaIM_osish']
SC={'Asosiy':{2026:6.9,2027:6.9,2028:6.9},'Salbiy':{2026:4.2,2027:5.5,2028:6.9},'Ijobiy':{2026:8.0,2027:8.0,2028:8.0}}
OUT['ssenariy']={k:v for k,v in SC.items()}; OUT['g25']=g25
P8={'LEV':'LEV','S3':'S3','NPLnet':'NPLnet','HLA':'HLA_r','LTD':'LTD','ROA':'ROA','ROE':'ROE','CI':'CI'}
W={b:v for b,v in M.W['taklif'].items() if b!='V'}
def mb(v):
  if sum(x is not None and x==x for x in v.values())<6: return np.nan,{}
  return M.mbii({k:(None if (x is None or x!=x) else x) for k,x in v.items()},W)
def path(start,gpath,rho=RHO,bg=BG,br=BR,brg=BRG,g0=g25):
  # start: dict 2025_IV holat; gpath: chorak bo'yicha YaIM o'sishi ro'yxati
  dev=0; res=[]
  for gq in gpath:
    dev=rho*dev+bg*(gq-g0)
    res.append(dev)
  return res
b25=d[d.davr=='2025_IV'].set_index('bank')
rows=[]
for sc,yy in SC.items():
  gp=[yy[y] for y in (2026,2027,2028) for _ in range(4)]
  for b,r in b25.iterrows():
    devs=path(None,gp)
    cov=min(r.RES/r.NPL,1.0) if (r.RES==r.RES and r.NPL>0) else 0.5
    for k,(y,dv) in enumerate(zip([2026]*4+[2027]*4+[2028]*4,devs)):
      if k%4!=3: continue
      s3=max(r.S3+dv,0.0)
      droa=BR*(s3-r.S3)
      roa=r.ROA+droa; roe=roa*r.Aktiv/r.Kapital
      npln=r.NPLnet+(s3-r.S3)*r.Kredit/r.Kapital*(1-cov) if r.NPLnet==r.NPLnet else np.nan
      v={'LEV':r.LEV,'S3':s3,'NPLnet':npln,'HLA':r.HLA_r,'LTD':r.LTD,'ROA':roa,'ROE':roe,'CI':r.CI}
      m,bs=mb(v)
      rows.append(dict(ssenariy=sc,bank=b,yil=y,DAVLAT=r.DAVLAT,S3=s3,ROA=roa,MBII=m,bA=bs.get('A'),bD=bs.get('D'),Aktiv=r.Aktiv))
F=pd.DataFrame(rows)
base=b25.reset_index()[['bank','DAVLAT','S3','ROA','MBII','Aktiv']].assign(ssenariy='2025 (haqiqiy)',yil=2025)
def zona(m): return 'yashil' if m>=3.25 else 'sariq' if m>=2.75 else 'qizil'
F['zona']=F.MBII.apply(zona); base['zona']=base.MBII.apply(zona)
def agg(x):
  wavg=lambda c: np.average(x[c],weights=x.Aktiv)
  return pd.Series({'S3_w':wavg('S3'),'ROA_w':wavg('ROA'),'MBII_med':x.MBII.median(),'MBII_w':wavg('MBII'),
    'yashil':(x.zona=='yashil').sum(),'sariq':(x.zona=='sariq').sum(),'qizil':(x.zona=='qizil').sum(),
    'qizil_aktiv_ulushi':x.loc[x.zona=='qizil','Aktiv'].sum()/x.Aktiv.sum()*100})
T=pd.concat([base,F]).groupby(['ssenariy','yil']).apply(agg).round(2); print(T.to_string()); OUT['tizim']={f'{a}|{b}':v for (a,b),v in T.to_dict('index').items()}
G=pd.concat([base,F]).groupby(['ssenariy','yil','DAVLAT']).apply(agg)[['S3_w','ROA_w','MBII_med','qizil']].round(2); print(G.to_string()); OUT['guruh']={f'{a}|{b}|{c}':v for (a,b,c),v in G.to_dict('index').items()}
# SQB va MKB: milliy va MHXS variantlari
R32=json.load(open('mbii32.json')); W13=M.W['taklif']; NAME={'SQB':'O‘zsanoatqurilishbank','MKB':'Mikrokreditbank'}
bk={}
for code,nm in NAME.items():
  x=R32[code]['2025']['ind']; r=b25.loc[nm]; koef=x['S3']/r.S3
  bk[code]={'koef':round(koef,2),'2025':{'milliy':round(r.MBII,2),'MHXS':round(R32[code]['2025']['mbii'],2)}}
  for sc in SC:
    for y in (2026,2027,2028):
      f=F[(F.bank==nm)&(F.ssenariy==sc)&(F.yil==y)].iloc[0]
      ds3=f.S3-r.S3; xi=dict(x); xi['S3']=x['S3']+koef*ds3
      if x.get('NPLnet') is not None: xi['NPLnet']=x['NPLnet']+koef*ds3*r.Kredit/r.Kapital*(1-min(r.RES/r.NPL,1))
      dro=f.ROA-r.ROA; xi['ROA']=x['ROA']+dro
      if x.get('ROE') is not None and x['ROA']: xi['ROE']=x['ROE']*(xi['ROA']/x['ROA']) if x['ROA']>0 else x['ROE']+dro*r.Aktiv/r.Kapital
      m,_=M.mbii(xi,W13)
      bk[code][f'{sc}|{y}']={'milliy':round(f.MBII,2),'MHXS':round(m,2),'S3_milliy':round(f.S3,2),'S3_MHXS':round(xi['S3'],2)}
print(json.dumps(bk,ensure_ascii=False,indent=0)); OUT['banklar']=bk
# orqaga tekshiruv: 2022_I–2024_IV da baholab, 2025 ni prognoz qilish (haqiqiy 2025 YaIM bilan)
tr=d[d.t<=2024*4+3]; a1,a2=est(tr); rho,bg,br,brg=a1.params.S3_l1,a1.params.YaIM_osish,a2.params.S3,0.0
b24=d[d.davr=='2024_IV'].set_index('bank'); g24=MK['2024_IV']['YaIM_osish']; gq=[MK[f'2025_{q}']['YaIM_osish'] for q in ('I','II','III','IV')]
bt=[]
for b,r in b24.iterrows():
  if b not in b25.index: continue
  dev=0
  for gg in gq: dev=rho*dev+bg*(gg-g24)
  s3=max(r.S3+dev,0); act=b25.loc[b]
  roa=r.ROA+br*(s3-r.S3)+brg*(np.mean(gq)-g24)
  v={'LEV':r.LEV,'S3':s3,'NPLnet':r.NPLnet,'HLA':r.HLA_r,'LTD':r.LTD,'ROA':roa,'ROE':roa*r.Aktiv/r.Kapital,'CI':r.CI}
  m,_=mb(v)
  bt.append(dict(bank=b,S3_prog=s3,S3_hq=act.S3,S3_naiv=r.S3,MBII_prog=m,MBII_hq=act.MBII,MBII_naiv=r.MBII,zona_prog=zona(m),zona_hq=zona(act.MBII)))
B=pd.DataFrame(bt).dropna()
OUT['backtest']={'N':len(B),'MAE_S3_model':round((B.S3_prog-B.S3_hq).abs().mean(),2),'MAE_S3_naiv':round((B.S3_naiv-B.S3_hq).abs().mean(),2),
 'MAE_MBII_model':round((B.MBII_prog-B.MBII_hq).abs().mean(),3),'MAE_MBII_naiv':round((B.MBII_naiv-B.MBII_hq).abs().mean(),3),
 'zona_mos':int((B.zona_prog==B.zona_hq).sum()),'tizim_S3_prog':round(B.S3_prog.mean(),2),'tizim_S3_hq':round(B.S3_hq.mean(),2),
 'koef':{'rho':round(rho,3),'bg':round(bg,3),'br':round(br,3)}}
print(OUT['backtest'])
json.dump(OUT,open('fc42.json','w'),ensure_ascii=False,indent=1,default=float)
