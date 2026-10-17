# 3.3-§: panel regressiyalar (asosiy: RE + chorak effektlari; dinamik: Anderson–Hsiao; robastlik: TWFE; satellit: NPL ~ makro)
import pandas as pd, numpy as np, json, warnings; warnings.filterwarnings('ignore')
from linearmodels.panel import PooledOLS, RandomEffects, PanelOLS
from linearmodels.iv import IV2SLS
from scipy import stats
from statsmodels.stats.outliers_influence import variance_inflation_factor as vif
p=pd.read_csv('panel33.csv'); d=p[p.tanlov].copy()
def w(x,a=.01): lo,hi=x.quantile([a,1-a]); return x.clip(lo,hi)
for c in ['KR_osish','KR_osish_l4','KR_osish_l8','NIM','DEP_A','S3','S3_l1','lnZ']: d[c]=w(d[c])
TD=pd.get_dummies(d.davr,prefix='ch',drop_first=True).astype(float); d=pd.concat([d,TD],axis=1)
d=d.set_index(['bank','t'],drop=False)
XB=['lnTA','KR_osish_l4','DEP_A','NIM','DAVLAT']
OUT={}
def keep(r,name,R2):
  OUT[name]={'b':r.params.round(4).to_dict(),'se':r.std_errors.round(4).to_dict(),'p':r.pvalues.round(4).to_dict(),'N':int(r.nobs),'R2':round(float(R2),3)}
  print('\n==',name,'N',r.nobs,'R2',round(float(R2),3))
  print(pd.DataFrame({'b':r.params,'se':r.std_errors,'p':r.pvalues}).filter(regex='^(?!ch_)',axis=0).round(4).to_string())
for y in ['MBII','lnZ']:
  dd=d[[y]+XB+list(TD)].dropna(); T=[c for c in TD if dd[c].sum()>0]; T=T[1:] if dd[T].sum(axis=1).min()>0 else T; Y=dd[y]
  po=PooledOLS(Y,dd[XB+T].assign(const=1)).fit(cov_type='clustered',cluster_entity=True); keep(po,y+' Pooled',po.rsquared)
  re=RandomEffects(Y,dd[XB+T].assign(const=1)).fit(cov_type='clustered',cluster_entity=True); keep(re,y+' RE',re.rsquared_overall)
  Xf=[x for x in XB if x!='DAVLAT']
  fe=PanelOLS(Y,dd[Xf],entity_effects=True,time_effects=True).fit(cov_type='clustered',cluster_entity=True); keep(fe,y+' TWFE',fe.rsquared_within)
  fe0=PanelOLS(Y,dd[Xf+T].assign(const=1),entity_effects=True).fit(); re0=RandomEffects(Y,dd[XB+T].assign(const=1)).fit()
  b=fe0.params[Xf]-re0.params[Xf]; V=fe0.cov.loc[Xf,Xf]-re0.cov.loc[Xf,Xf]
  H=float(b@np.linalg.pinv(V)@b); pv=1-stats.chi2.cdf(H,len(Xf)); print('Hausman',round(H,2),'p',round(pv,4))
  OUT[y+' Hausman']={'H':round(H,2),'df':len(Xf),'p':round(pv,4)}
  # Breusch–Pagan LM (RE vs Pooled)
  e=po.resids; n=dd.groupby(level=0).size(); se=e.groupby(level=0).sum()
  LM=len(e)/(2*(n.mean()-1))*((se**2).sum()/(e**2).sum()-1)**2; OUT[y+' BP_LM']={'LM':round(LM,2),'p':round(1-stats.chi2.cdf(LM,1),4)}; print('BP LM',round(LM,2))
# davlat ulushining umumiy ta'siri (mediatsiya)
for X,tag in ((['DAVLAT','lnTA'],'davlat_umumiy'),(['DAVLAT','lnTA','DEP_A','NIM'],'davlat_funding')):
  dd=d[['MBII']+X+list(TD)].dropna(); T=[c for c in TD if dd[c].sum()>0]; T=T[1:] if dd[T].sum(axis=1).min()>0 else T
  r=PooledOLS(dd.MBII,dd[X+T].assign(const=1)).fit(cov_type='clustered',cluster_entity=True); keep(r,'MBII '+tag,r.rsquared)
# dinamik: Anderson–Hsiao (birinchi farqlar, MBII_{t-2} instrument, chorak effektlari)
r=d.reset_index(drop=True).sort_values(['bank','t']); g=r.groupby('bank')
Xd=['lnTA','KR_osish_l4','DEP_A','NIM']
for c in ['MBII','MBII_l1']+Xd: r['D_'+c]=g[c].diff()
r['MBII_l2']=g['MBII'].shift(2)
q=r.dropna(subset=['D_MBII','D_MBII_l1','MBII_l2']+['D_'+x for x in Xd]); T=[c for c in TD if q[c].sum()>0]; T=T[1:] if q[T].sum(axis=1).min()>0 else T
iv=IV2SLS(q.D_MBII,q[['D_'+x for x in Xd]+T].assign(const=1),q[['D_MBII_l1']],q[['MBII_l2']]).fit(cov_type='clustered',clusters=q.bank.astype('category').cat.codes)
keep(iv,'MBII Anderson-Hsiao',0); F=float(iv.first_stage.diagnostics['f.stat'].iloc[0]); print('1-bosqich F',round(F,1)); OUT['MBII Anderson-Hsiao']['F1']=round(F,1)
# satellit: NPL ulushi ~ makroomillar
Xs=['S3_l1','KR_osish_l4','KR_osish_l8','lnTA','YaIM_osish','Asosiy_stavka','USD_osish','Inflyatsiya']
dd=d[['S3']+Xs+['DAVLAT']].dropna()
fe=PanelOLS(dd.S3,dd[Xs].assign(const=1),entity_effects=True).fit(cov_type='clustered',cluster_entity=True); keep(fe,'S3 FE',fe.rsquared_within)
po=PooledOLS(dd.S3,dd[Xs+['DAVLAT']].assign(const=1)).fit(cov_type='clustered',cluster_entity=True); keep(po,'S3 Pooled',po.rsquared)
X=d[XB].dropna(); OUT['VIF']={x:round(vif(X.assign(c=1).values,i),2) for i,x in enumerate(XB)}; print('VIF',OUT['VIF'])
# deskriptiv
D=d[['MBII','bK','bA','bL','bD','lnZ','LEV','S3','NPLnet','HLA_r','LTD','ROA','ROE','CI','lnTA','KR_osish','DEP_A','NIM']].describe().T[['count','mean','std','min','50%','max']].round(2)
OUT['deskriptiv']=D.to_dict('index'); print(D.to_string())
E=d[d.q==4].groupby(['y','DAVLAT'])[['MBII','bK','bA','bL','bD','lnZ']].median().round(2); print(E.to_string())
OUT['guruh']={f'{a}_{b}':v for (a,b),v in E.to_dict('index').items()}
json.dump(OUT,open('model33.json','w'),ensure_ascii=False,indent=1)
