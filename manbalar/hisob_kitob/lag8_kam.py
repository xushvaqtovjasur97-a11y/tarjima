# Kamchiliklar jadvali (29-band): H2 – kreditlar o'sishining 8 chorak lagli ta'siri (MBII va ln Z)
import pandas as pd, numpy as np, json, warnings; warnings.filterwarnings('ignore')
from linearmodels.panel import PooledOLS, RandomEffects, PanelOLS
p=pd.read_csv('panel33.csv'); d=p[p.tanlov].copy()
def w(x,a=.01): lo,hi=x.quantile([a,1-a]); return x.clip(lo,hi)
for c in ['KR_osish_l4','KR_osish_l8','NIM','DEP_A','lnZ']: d[c]=w(d[c])
TD=pd.get_dummies(d.davr,prefix='ch',drop_first=True).astype(float); d=pd.concat([d,TD],axis=1)
d=d.set_index(['bank','t'],drop=False)
OUT={}
for y in ['MBII','lnZ']:
  for tag,XB in (('l8',['lnTA','KR_osish_l8','DEP_A','NIM','DAVLAT']),('l4+l8',['lnTA','KR_osish_l4','KR_osish_l8','DEP_A','NIM','DAVLAT'])):
    dd=d[[y]+XB+list(TD)].dropna(); T=[c for c in TD if dd[c].sum()>0]; T=T[1:] if dd[T].sum(axis=1).min()>0 else T; Y=dd[y]
    po=PooledOLS(Y,dd[XB+T].assign(const=1)).fit(cov_type='clustered',cluster_entity=True)
    Xf=[x for x in XB if x!='DAVLAT']
    fe=PanelOLS(Y,dd[Xf],entity_effects=True,time_effects=True).fit(cov_type='clustered',cluster_entity=True)
    re=RandomEffects(Y,dd[XB+T].assign(const=1)).fit(cov_type='clustered',cluster_entity=True)
    for nm,r in (('Pooled',po),('RE',re),('FE',fe)):
      ks=[k for k in ('KR_osish_l4','KR_osish_l8') if k in r.params]
      OUT[f'{y}|{tag}|{nm}']={k:dict(b=round(float(r.params[k]),4),se=round(float(r.std_errors[k]),4),p=round(float(r.pvalues[k]),3)) for k in ks}; OUT[f'{y}|{tag}|{nm}']['N']=int(r.nobs)
      print(y,tag,nm,OUT[f'{y}|{tag}|{nm}'])
json.dump(OUT,open('lag8_kam.json','w'),ensure_ascii=False,indent=1)
