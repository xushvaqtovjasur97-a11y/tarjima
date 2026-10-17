# 3.2-§: integral indeks vaznlari va chegaralarini empirik asoslash
import pandas as pd, numpy as np, json, warnings; warnings.filterwarnings('ignore')
import statsmodels.api as sm
import mbii as M
OUT={}
EXP=M.W['taklif']; BL=['K','A','L','D','V']
# --- A) MHXS asosidagi 5 bank × 5 yil, 13 ko'rsatkich
R=json.load(open('mbii32.json')); rows=[]
for b in R:
  for y in R[b]:
    x=R[b][y]['ind']; sc={}
    for k,bl,n,dr,a in M.IND:
      if x.get(n) is not None: sc[n]=M.score(x[n],dr,a)
    rows.append(dict(bank=b,yil=int(y),**sc))
S=pd.DataFrame(rows); IND=[n for _,_,n,_,_ in M.IND]; BOF={n:bl for _,bl,n,_,_ in M.IND}
# entropiya usuli (ballar 1–4 bo'yicha)
def entropy_w(X):
  P=X.div(X.sum()); n=len(X); E=-(P*np.log(P)).sum()/np.log(n); d=1-E; return d/d.sum()
X=S[IND].apply(lambda c:c.fillna(c.mean()))
we=entropy_w(X); wb_e=we.groupby(lambda n:BOF[n]).sum().reindex(BL)
# bosh komponentlar: bloklar ballari bo'yicha
B=pd.DataFrame([R[b][y]['bloklar'] for b in R for y in R[b]])[BL]
Z=(B-B.mean())/B.std(ddof=0); Z=Z.fillna(0)
ev,evec=np.linalg.eigh(np.cov(Z.T)); o=np.argsort(ev)[::-1]; ev,evec=ev[o],evec[:,o]
k=int((ev>1).sum()) or 1
load=np.abs(evec[:,:k])*np.sqrt(ev[:k]); wpca=(load**2*ev[:k]/ev[:k].sum()).sum(axis=1)
wb_p=pd.Series(wpca/wpca.sum(),index=BL)
OUT['pca']={'ozqiymatlar':ev.round(3).tolist(),'izohlangan_ulush':(ev/ev.sum()).round(3).tolist(),'komponentlar':k}
# variatsiya koeffitsiyenti usuli (qo'shimcha)
cv=X.std()/X.mean(); wb_cv=(cv/cv.sum()).groupby(lambda n:BOF[n]).sum().reindex(BL)
# --- B) panel (25 bank, 8 ko'rsatkich) – logit asosidagi vaznlar (4.1 hodisalari)
p=pd.read_csv('panel33.csv').sort_values(['bank','t']); g=p.groupby('bank')
p['lossq']=-p.NPq/g.Kapital.shift(1)*100; p['EV']=(p.lossq>=5).astype(int)
for j in range(1,5): p[f'e{j}']=g.EV.shift(-j).where(g.t.shift(-j)==p.t+j)
p['Y']=p[[f'e{j}' for j in range(1,5)]].max(axis=1)
d=p[p.tanlov&(p.t<=2024*4+3)&(p.EV==0)].dropna(subset=['Y','bA','bL','bD'])
lg=sm.Logit(d.Y,sm.add_constant(d[['bA','bL','bD']])).fit(disp=0,cov_type='cluster',cov_kwds={'groups':d.bank.astype('category').cat.codes})
coef=-lg.params[['bA','bL','bD']].rename({'bA':'A','bL':'L','bD':'D'}); OUT['logit']={'b':lg.params.round(3).to_dict(),'p':lg.pvalues.round(4).to_dict(),'N':int(lg.nobs)}
# panelda entropiya (8 ko'rsatkich ballari)
P8={'LEV':'LEV','S3':'S3','NPLnet':'NPLnet','HLA':'HLA_r','LTD':'LTD','ROA':'ROA','ROE':'ROE','CI':'CI'}
pp=p[p.tanlov].copy(); SP=pd.DataFrame({n:pp[c].apply(lambda v: M.score(v,[x for x in M.IND if x[2]==n][0][3],[x for x in M.IND if x[2]==n][0][4]) if v==v else np.nan) for n,c in P8.items()})
SP=SP.apply(lambda c:c.fillna(c.mean())); wpe=entropy_w(SP)
OUT['panel_entropiya_korsatkich']=wpe.round(3).to_dict()
OUT['ind_entropiya']=we.round(3).to_dict()
W=pd.DataFrame({'Ekspert':pd.Series(EXP)[BL],'Entropiya':wb_e,'Variatsiya':wb_cv,'Bosh komponentlar':wb_p})
cl=coef.clip(lower=0); W['Logit (K, V – ekspert)']=pd.Series({'K':EXP['K'],'V':EXP['V'],**(cl/cl.sum()*(1-EXP['K']-EXP['V'])).to_dict()})[BL]
# o'rtacha faqat barcha 5 blokni empirik baholagan usullardan olinadi (logit K va V ni baholamagan, unda ekspert vaznlari bor)
W['O‘rtacha (empirik)']=W[['Entropiya','Variatsiya','Bosh komponentlar']].mean(axis=1)
print(W.round(3).to_string()); OUT['vaznlar']=W.round(3).to_dict()
# --- C) 5 bankka qo'llash: turli vaznlarda MBII va sinf
def mbii_w(code,y,w): return M.mbii(R[code][y]['ind'],w)[0]
res={}
for col in W.columns:
  w=W[col].to_dict(); res[col]={f'{c}|{y}':round(mbii_w(c,y,w),2) for c in R for y in R[c]}
OUT['bank_mbii']=res
T=pd.DataFrame({col:{k:v for k,v in res[col].items() if k.endswith('2025')} for col in W.columns}); print(T.to_string())
# sinflar barqarorligi
base=res['Ekspert']; same={col:sum(M.cls(base[k])==M.cls(res[col][k]) for k in base) for col in W.columns}; OUT['sinf_mos']=same; print('sinf mos (25 dan):',same)
from scipy.stats import spearmanr
OUT['spearman']={col:round(spearmanr([base[k] for k in base],[res[col][k] for k in base])[0],3) for col in W.columns}; print(OUT['spearman'])
# --- D) Monte-Karlo: 10000 tasodifiy vazn (Dirichlet, ekspert atrofida va bir tekis)
rng=np.random.default_rng(42); keys=list(base)
def mc(alpha,N=10000):
  sm_=np.zeros(len(keys)); rk=[]; mx=np.zeros(len(keys))
  for _ in range(N):
    wv=rng.dirichlet(alpha); w=dict(zip(BL,wv))
    v=np.array([mbii_w(k.split('|')[0],k.split('|')[1],w) for k in keys])
    sm_+=np.array([M.cls(a)==M.cls(base[k]) for a,k in zip(v,keys)])
    rk.append(spearmanr(v,[base[k] for k in keys])[0])
  return (sm_/N), np.array(rk)
for nm,alpha in (('ekspert_atrofida',np.array([EXP[b] for b in BL])*20),('bir_tekis',np.ones(5))):
  share,rk=mc(alpha,4000)
  OUT['mc_'+nm]={'sinf_saqlanishi_ortacha':round(share.mean()*100,1),'sinf_saqlanishi_min':round(share.min()*100,1),'eng_sezgir':keys[int(share.argmin())],
    'spearman_median':round(float(np.median(rk)),3),'spearman_p5':round(float(np.percentile(rk,5)),3),'bank_yil':{k:round(s*100,1) for k,s in zip(keys,share)}}
  print(nm,{k:v for k,v in OUT['mc_'+nm].items() if k!='bank_yil'})
# 2025 tartibi barqarorligi (5 bank)
# --- E) sinf chegaralari: panelda sinflar bo'yicha hodisa chastotasi
d['sinf']=d.MBII.apply(M.cls)
fr=d.groupby('sinf').Y.agg(['count','sum','mean']).reindex(['barqaror','qoniqarli','zaif','nobarqaror']); print(fr); OUT['sinf_hodisa']=fr.fillna(0).round(3).to_dict('index')
# muqobil chegaralar: Youden bo'yicha optimal (4.1 bilan bir xil yondashuv)
OUT['davlat_komagisiz']={'MKB_2025':1.72,'MKB_amaldagi':2.47}
json.dump(OUT,open('vazn32.json','w'),ensure_ascii=False,indent=1,default=float)
