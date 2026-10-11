# 3.3-§: panel regressiyalar
import pandas as pd, numpy as np, json, warnings; warnings.filterwarnings('ignore')
from linearmodels.panel import PooledOLS, RandomEffects, PanelOLS
from linearmodels.iv import IV2SLS
from scipy import stats
p=pd.read_csv('panel33.csv'); d=p[p.tanlov].copy()
def w(x,a=.01): lo,hi=x.quantile([a,1-a]); return x.clip(lo,hi)
for c in ['KR_osish','KR_osish_l4','KR_osish_l8','NIM','DEP_A','ROA','S3','lnZ','S3_l1','lnZ_l1']: d[c]=w(d[c])
d=d.set_index(['bank','t'])
X=['lnTA','KR_osish_l4','DEP_A','NIM','DAVLAT','YaIM_osish','Inflyatsiya','Asosiy_stavka','USD_osish']
OUT={}
def tab(r,name):
  OUT[name]={'coef':r.params.round(4).to_dict(),'p':r.pvalues.round(4).to_dict(),'N':int(r.nobs),
    'R2':round(float(getattr(r,'rsquared_within',r.rsquared) if 'FE' in name else r.rsquared),3)}
  print('\n==',name,'N',r.nobs,'R2',round(r.rsquared,3))
  print(pd.DataFrame({'b':r.params,'se':r.std_errors,'p':r.pvalues}).round(4).to_string())
def run(y,X,tag):
  dd=d[[y]+X].dropna(); Y=dd[y]; XX=dd[X].assign(const=1)
  po=PooledOLS(Y,XX).fit(cov_type='clustered',cluster_entity=True); tab(po,tag+' Pooled')
  re=RandomEffects(Y,XX).fit(cov_type='clustered',cluster_entity=True); tab(re,tag+' RE')
  Xf=[x for x in X if x!='DAVLAT']
  fe=PanelOLS(Y,dd[Xf].assign(const=1),entity_effects=True).fit(cov_type='clustered',cluster_entity=True); tab(fe,tag+' FE')
  # Hausman (klassik kovariatsiyalar bilan)
  fe0=PanelOLS(Y,dd[Xf].assign(const=1),entity_effects=True).fit(); re0=RandomEffects(Y,XX).fit()
  b=fe0.params[Xf]-re0.params[Xf]; V=fe0.cov.loc[Xf,Xf]-re0.cov.loc[Xf,Xf]
  H=float(b@np.linalg.pinv(V)@b); pv=1-stats.chi2.cdf(H,len(Xf)); print('Hausman',round(H,2),'p',round(pv,4))
  OUT[tag+' Hausman']={'H':round(H,2),'p':round(pv,4)}
  print('RE theta',re0.theta.mean().round(3) if hasattr(re0,'theta') else '')
run('MBII',X,'MBII')
run('lnZ',X,'lnZ')
# dinamik: FE + lag (LSDV) va Anderson–Hsiao (birinchi farqlar, y_{t-2} instrument)
dd=d[['MBII','MBII_l1']+X].dropna()
fe=PanelOLS(dd.MBII,dd[['MBII_l1']+[x for x in X if x!='DAVLAT']].assign(const=1),entity_effects=True).fit(cov_type='clustered',cluster_entity=True); tab(fe,'MBII dinamik FE')
r=d.reset_index().sort_values(['bank','t']); g=r.groupby('bank')
Xd=['lnTA','KR_osish_l4','DEP_A','NIM','YaIM_osish','Inflyatsiya','Asosiy_stavka','USD_osish']
for c in ['MBII','MBII_l1']+Xd: r['D_'+c]=g[c].diff()
r['MBII_l2']=g['MBII'].shift(2)
q=r.dropna(subset=['D_MBII','D_MBII_l1','MBII_l2']+['D_'+x for x in Xd])
iv=IV2SLS(q.D_MBII,q[['D_'+x for x in Xd]].assign(const=1),q[['D_MBII_l1']],q[['MBII_l2']]).fit(cov_type='clustered',clusters=q.bank.astype('category').cat.codes)
print('\n== Anderson-Hsiao N',iv.nobs); print(pd.DataFrame({'b':iv.params,'se':iv.std_errors,'p':iv.pvalues}).round(4).to_string())
print('1-bosqich F',round(iv.first_stage.diagnostics['f.stat'].iloc[0],1))
OUT['MBII AH']={'coef':iv.params.round(4).to_dict(),'p':iv.pvalues.round(4).to_dict(),'N':int(iv.nobs)}
# satellit: NPL ulushi kredit o'sishi lagi bilan
Xs=['S3_l1','KR_osish_l4','KR_osish_l8','lnTA','YaIM_osish','Asosiy_stavka','USD_osish']
dd=d[['S3']+Xs+['DAVLAT']].dropna()
fe=PanelOLS(dd.S3,dd[Xs].assign(const=1),entity_effects=True).fit(cov_type='clustered',cluster_entity=True); tab(fe,'S3 satellit FE')
po=PooledOLS(dd.S3,dd[Xs+['DAVLAT']].assign(const=1)).fit(cov_type='clustered',cluster_entity=True); tab(po,'S3 satellit Pooled')
json.dump(OUT,open('model33.json','w'),ensure_ascii=False,indent=1)
