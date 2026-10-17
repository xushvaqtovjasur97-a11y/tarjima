# 4.2-§: 2026–2028-yillar shartli ssenariy prognozi
# Sezgirlik: tizim darajasidagi NPL modeli (2020_II–2025_IV; 2020-yil faqat baholash uchun), YaIM 4 chorak lag bilan;
# bank darajasiga proporsional taqsimlanadi; ROA – bank satelliti; boshqa ko'rsatkichlar 2025_IV darajasida.
import pandas as pd, numpy as np, json, warnings; warnings.filterwarnings('ignore')
import statsmodels.api as sm
from linearmodels.panel import PanelOLS
import mbii as M
OUT={}
D=pd.read_csv('tizim42.csv',index_col=0); K=list(D.index)
def fit_sys(end='2025_IV',k=4):
  q=pd.concat([D.NPL,D.NPL.shift(1).rename('NPL_l1'),D.YaIM.shift(k).rename('G')],axis=1).loc[:end].dropna()
  return sm.OLS(q.NPL,sm.add_constant(q[['NPL_l1','G']])).fit(cov_type='HAC',cov_kwds={'maxlags':2}),q
rs,qs=fit_sys(); RHO,BG=rs.params.NPL_l1,rs.params.G
OUT['tizim_model']={'b':rs.params.round(4).to_dict(),'se':rs.bse.round(4).to_dict(),'p':rs.pvalues.round(4).to_dict(),'N':int(rs.nobs),'R2':round(rs.rsquared,3),'davr':f'{qs.index[0]}–{qs.index[-1]}','LR':round(BG/(1-RHO),3)}
# muqobil lag tanlovi (AIC)
OUT['lag_tanlov']={k:dict(b=round(fit_sys(k=k)[0].params.G,3),p=round(fit_sys(k=k)[0].pvalues.G,3),AIC=round(fit_sys(k=k)[0].aic,1)) for k in range(0,6)}
# bank paneli satellitlari
p=pd.read_csv('panel33.csv'); d=p[p.tanlov].copy().sort_values(['bank','t'])
def w(x,a=.01): lo,hi=x.quantile([a,1-a]); return x.clip(lo,hi)
for c in ['S3','S3_l1','ROAq']: d[c]=w(d[c])
q=d.set_index(['bank','t'])[['S3','S3_l1','YaIM_osish']].dropna()
rp=PanelOLS(q.S3,q[['S3_l1','YaIM_osish']].assign(c=1),entity_effects=True).fit(cov_type='clustered',cluster_entity=True)
q=d.set_index(['bank','t'])[['ROAq','S3']].dropna()
rr=PanelOLS(q.ROAq,q[['S3']].assign(c=1),entity_effects=True).fit(cov_type='clustered',cluster_entity=True); BR=rr.params.S3
for nm,r,R2 in (('panel_NPL',rp,rp.rsquared_within),('panel_ROA',rr,rr.rsquared_within)):
  OUT[nm]={'b':r.params.round(4).to_dict(),'se':r.std_errors.round(4).to_dict(),'p':r.pvalues.round(4).to_dict(),'N':int(r.nobs),'R2w':round(float(R2),3)}
OUT['panel_NPL']['LR']=round(rp.params.YaIM_osish/(1-rp.params.S3_l1),3)
print(json.dumps({k:OUT[k] for k in ('tizim_model','panel_NPL','panel_ROA','lag_tanlov')},ensure_ascii=False))
# ssenariylar: yillik YaIM o'sishi (2021–2025 va baholash davridagi haqiqiy holatlar asosida)
SC={'Asosiy':{2026:6.9,2027:6.9,2028:6.9},'Ijobiy':{2026:8.0,2027:8.0,2028:8.0},
    'Salbiy':{2026:4.5,2027:5.5,2028:6.9},'Keskin salbiy':{2026:1.6,2027:5.5,2028:6.9}}
OUT['ssenariy']=SC
Gh=D.YaIM.to_dict()
def sys_dev(gpath,start_idx,hist=Gh,keys=K,rho=RHO,bg=BG,k=4):
  # 2025-yil darajasi saqlanadi deb olingan bazaga nisbatan og'ish: dev_t = rho*dev_{t-1} + bg*(G_{t-k} - G_ref)
  # G_ref – start nuqtasidagi oxirgi 4 chorak YaIM (o'zgarishsiz holat)
  allG=[hist[x] for x in keys[:start_idx+1]]+list(gpath); ref=hist[keys[start_idx]]
  dev=0; out=[]
  for j in range(1,len(gpath)+1):
    gl=allG[start_idx+j-k]; dev=rho*dev+bg*(gl-ref); out.append(dev)
  return out
i25=K.index('2025_IV'); S25=D.NPL.iloc[i25]
SYS={}
for sc,yy in SC.items():
  gp=[yy[y] for y in (2026,2027,2028) for _ in range(4)]
  SYS[sc]=[round(S25+x,2) for x in sys_dev(gp,i25)]
OUT['tizim_NPL_yo‘li']=SYS; print('Tizim NPL:',{k:v[3::4] for k,v in SYS.items()})
# banklar
P8={'LEV':'LEV','S3':'S3','NPLnet':'NPLnet','HLA':'HLA_r','LTD':'LTD','ROA':'ROA','ROE':'ROE','CI':'CI'}
W={b:v for b,v in M.W['taklif'].items() if b!='V'}
def mb(v):
  v={k:(None if (x is None or x!=x) else x) for k,x in v.items()}
  if sum(x is not None for x in v.values())<6: return np.nan,{}
  return M.mbii(v,W)
def zona(m): return 'yashil' if m>=3.25 else 'sariq' if m>=2.75 else 'qizil'
def bank_proj(r,s3sys_dev,s3sys0,br=BR):
  s3=max(r.S3*(1+s3sys_dev/s3sys0),0.0)          # proporsional taqsimlash
  roa=r.ROA+br*(s3-r.S3); roe=roa*r.Aktiv/r.Kapital
  cov=min(r.RES/r.NPL,1.0) if (r.RES==r.RES and r.NPL>0) else 0.5
  npln=r.NPLnet+(s3-r.S3)*r.Kredit/r.Kapital*(1-cov) if r.NPLnet==r.NPLnet else np.nan
  v={'LEV':r.LEV,'S3':s3,'NPLnet':npln,'HLA':r.HLA_r,'LTD':r.LTD,'ROA':roa,'ROE':roe,'CI':r.CI}
  m,bs=mb(v); return s3,roa,m,bs
b25=d[d.davr=='2025_IV'].set_index('bank')
rows=[]
for sc in SC:
  for j,y in ((3,2026),(7,2027),(11,2028)):
    dv=SYS[sc][j]-S25
    for b,r in b25.iterrows():
      s3,roa,m,bs=bank_proj(r,dv,S25)
      rows.append(dict(ssenariy=sc,yil=y,bank=b,DAVLAT=r.DAVLAT,Aktiv=r.Aktiv,S3=s3,ROA=roa,MBII=m,bA=bs.get('A'),bD=bs.get('D')))
F=pd.DataFrame(rows); base=b25.reset_index()[['bank','DAVLAT','Aktiv','S3','ROA','MBII']].assign(ssenariy='2025',yil=2025)
A=pd.concat([base,F]); A['zona']=A.MBII.apply(zona)
def agg(x):
  wv=lambda c: np.average(x[c],weights=x.Aktiv)
  return pd.Series({'S3_w':wv('S3'),'ROA_w':wv('ROA'),'MBII_med':x.MBII.median(),'MBII_w':wv('MBII'),'yashil':(x.zona=='yashil').sum(),'sariq':(x.zona=='sariq').sum(),'qizil':(x.zona=='qizil').sum(),'qizil_aktiv':x.loc[x.zona=='qizil','Aktiv'].sum()/x.Aktiv.sum()*100,'sariq_qizil_aktiv':x.loc[x.zona!='yashil','Aktiv'].sum()/x.Aktiv.sum()*100})
T=A.groupby(['ssenariy','yil']).apply(agg).round(2); print(T.to_string()); OUT['tizim']={f'{a}|{b}':v for (a,b),v in T.to_dict('index').items()}
G=A.groupby(['ssenariy','yil','DAVLAT']).apply(agg)[['S3_w','ROA_w','MBII_med','sariq','qizil']].round(2); OUT['guruh']={f'{a}|{b}|{c}':v for (a,b,c),v in G.to_dict('index').items()}
print(G.loc[['2025','Salbiy','Keskin salbiy']].to_string())
OUT['qizil_banklar']={f'{a}|{b}':sorted(x.bank.tolist()) for (a,b),x in A[A.zona=='qizil'].groupby(['ssenariy','yil'])}
print(OUT['qizil_banklar'])
# SQB, MKB: milliy va MHXS
R32=json.load(open('mbii32.json')); W13=M.W['taklif']; NAME={'SQB':'O‘zsanoatqurilishbank','MKB':'Mikrokreditbank'}
bk={}
for code,nm in NAME.items():
  x=R32[code]['2025']['ind']; r=b25.loc[nm]; koef=x['S3']/r.S3
  bk[code]={'koef':round(koef,2),'2025':{'milliy':round(r.MBII,2),'MHXS':round(R32[code]['2025']['mbii'],2),'S3_milliy':round(r.S3,2),'S3_MHXS':round(x['S3'],2)}}
  for sc in SC:
    for y in (2026,2027,2028):
      f=F[(F.bank==nm)&(F.ssenariy==sc)&(F.yil==y)].iloc[0]; ds3=f.S3-r.S3
      xi=dict(x); xi['S3']=x['S3']+koef*ds3
      if x.get('NPLnet') is not None: xi['NPLnet']=x['NPLnet']+koef*ds3*r.Kredit/r.Kapital*(1-min(r.RES/r.NPL,1))
      dro=BR*koef*ds3; xi['ROA']=x['ROA']+dro
      if x.get('ROE') is not None: xi['ROE']=x['ROE']+dro*r.Aktiv/r.Kapital
      m,_=M.mbii(xi,W13)
      bk[code][f'{sc}|{y}']={'milliy':round(f.MBII,2),'MHXS':round(m,2),'S3_milliy':round(f.S3,2),'S3_MHXS':round(xi['S3'],2)}
OUT['banklar']=bk
for c in bk: print(c,{k:v for k,v in bk[c].items() if k.endswith('2027') or k=='2025'})
# ko'p martalik orqaga tekshiruv: 2022_IV, 2023_IV, 2024_IV dan keyingi 4 chorak
bt=[]
for end in ('2022_IV','2023_IV','2024_IV'):
  r,_=fit_sys(end); ie=K.index(end); act=D.NPL.iloc[ie+1:ie+5].values
  gp=D.YaIM.iloc[ie+1:ie+5].tolist()
  dev=sys_dev(gp,ie,rho=r.params.NPL_l1,bg=r.params.G)
  ours=D.NPL.iloc[ie]+np.array(dev)
  n=D.NPL.iloc[ie]; mr=[]
  for j in range(1,5): n=r.params.const+r.params.NPL_l1*n+r.params.G*D.YaIM.iloc[ie+j-4]; mr.append(n)
  bt.append(dict(davr=end[:4],haqiqiy=act.round(2).tolist(),bizning=ours.round(2).tolist(),
     MAE_bizning=round(np.abs(ours-act).mean(),2),MAE_naiv=round(np.abs(D.NPL.iloc[ie]-act).mean(),2),MAE_qaytuvchi=round(np.abs(np.array(mr)-act).mean(),2)))
OUT['backtest_tizim']=bt; print(pd.DataFrame(bt).to_string())
# bank darajasida 2025 bo'yicha: zonalar
ie=K.index('2024_IV'); r,_=fit_sys('2024_IV'); gp=D.YaIM.iloc[ie+1:ie+5].tolist()
dev=sys_dev(gp,ie,rho=r.params.NPL_l1,bg=r.params.G)[-1]
b24=d[d.davr=='2024_IV'].set_index('bank'); z=[]
for b,rr_ in b24.iterrows():
  if b not in b25.index: continue
  s3,roa,m,_=bank_proj(rr_,dev,D.NPL.iloc[ie])
  if m==m and b25.loc[b].MBII==b25.loc[b].MBII: z.append((zona(m),zona(b25.loc[b].MBII),abs(m-b25.loc[b].MBII),abs(rr_.MBII-b25.loc[b].MBII)))
OUT['backtest_bank']={'N':len(z),'zona_mos':sum(a==b for a,b,_,_ in z),'MAE_MBII':round(np.mean([c for *_,c,_ in z]),3),'MAE_MBII_naiv':round(np.mean([e for *_,e in z]),3)}
print(OUT['backtest_bank'])
json.dump(OUT,open('fc42.json','w'),ensure_ascii=False,indent=1,default=float)
