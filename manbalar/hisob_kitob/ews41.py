# 4.1-§: erta ogohlantirish – signallar usuli, logit, ROC/AUC, vaqtdan tashqari tekshiruv
import pandas as pd, numpy as np, json, warnings; warnings.filterwarnings('ignore')
import statsmodels.api as sm
from scipy import stats
p=pd.read_csv('panel33.csv').sort_values(['bank','t']); g=p.groupby('bank')
p['K_l1']=g.Kapital.shift(1)
p['lossq']=-p.NPq/p.K_l1*100
p['EV']=(p.lossq>=5).astype(int)
# kelgusi 1–4 chorakda hodisa
for k in range(1,5): p[f'ev{k}']=g.EV.shift(-k).where(g.t.shift(-k)==p.t+k)
p['Y']=p[[f'ev{k}' for k in range(1,5)]].max(axis=1)
d=p[p.tanlov&(p.t<=2024*4+3)&(p.EV==0)].copy()     # oyna to'liq kuzatilgan, hodisa choragining o'zi chiqarilgan
d=d.dropna(subset=['Y','MBII'])
def auc(y,s):  # s – risk bali (katta = xavfli)
  y=np.asarray(y); s=np.asarray(s); pos=s[y==1]; neg=s[y==0]
  return (np.greater.outer(pos,neg).sum()+0.5*np.equal.outer(pos,neg).sum())/(len(pos)*len(neg))
def delong(y,s1,s2):
  y=np.asarray(y)
  def comp(s):
    pos=s[y==1]; neg=s[y==0]
    V10=np.array([((x>neg).sum()+0.5*(x==neg).sum())/len(neg) for x in pos]); V01=np.array([((pos>x).sum()+0.5*(pos==x).sum())/len(pos) for x in neg])
    return V10,V01
  a=[comp(np.asarray(s)) for s in (s1,s2)]
  A=[v[0].mean() for v in a]
  S10=np.cov(np.vstack([a[0][0],a[1][0]])); S01=np.cov(np.vstack([a[0][1],a[1][1]]))
  S=S10/len(a[0][0])+S01/len(a[0][1]); var=S[0,0]+S[1,1]-2*S[0,1]
  z=(A[0]-A[1])/np.sqrt(var); return A[0],A[1],2*(1-stats.norm.cdf(abs(z)))
def boot(y,s,B=2000,seed=1):
  rng=np.random.default_rng(seed); y=np.asarray(y); s=np.asarray(s); r=[]
  for _ in range(B):
    i=rng.integers(0,len(y),len(y))
    if y[i].sum()>0 and y[i].sum()<len(y): r.append(auc(y[i],s[i]))
  return np.percentile(r,[2.5,97.5])
OUT={'N':int(len(d)),'pos':int(d.Y.sum()),'hodisalar':p[p.tanlov&(p.EV==1)][['bank','davr','lossq']].round(1).values.tolist()}
# 1) bir o'lchovli AUC: risk bali sifatida -MBII, -lnZ, S3, -LEV, -bloklar
cand={'MBII':-d.MBII,'ln(Z-score)':-d.lnZ,'Muammoli kreditlar ulushi':d.S3,'Kapital / aktivlar':-d.LEV,'ROA':-d.ROA,
 'Aktivlar sifati bloki':-d.bA,'Likvidlik bloki':-d.bL,'Daromadlilik bloki':-d.bD}
res=[]
for k,s in cand.items():
  m=s.notna(); a=auc(d.Y[m],s[m]); lo,hi=boot(d.Y[m],s[m]); res.append((k,int(m.sum()),int(d.Y[m].sum()),a,lo,hi))
OUT['AUC']=res
for r in res: print(f'{r[0]:28} N={r[1]} pos={r[2]} AUC={r[3]:.3f} [{r[4]:.3f};{r[5]:.3f}]')
m=d.lnZ.notna(); a1,a2,pv=delong(d.Y[m],-d.MBII[m].values,-d.lnZ[m].values); print('DeLong MBII vs lnZ',round(a1,3),round(a2,3),round(pv,3)); OUT['DeLong_Z']=[a1,a2,pv]
a1,a2,pv=delong(d.Y,-d.MBII.values,d.S3.values); print('DeLong MBII vs S3',round(a1,3),round(a2,3),round(pv,3)); OUT['DeLong_S3']=[a1,a2,pv]
# 2) signallar usuli: chegaralar
def sig(th,s=d.MBII,y=d.Y):
  S=s<th; TP=int((S&(y==1)).sum()); FN=int((~S&(y==1)).sum()); FP=int((S&(y==0)).sum()); TN=int((~S&(y==0)).sum())
  t1=FN/(TP+FN); t2=FP/(FP+TN); nsr=t2/(1-t1) if t1<1 else np.nan
  return dict(th=th,TP=TP,FN=FN,FP=FP,TN=TN,I_tur=t1*100,II_tur=t2*100,NSR=nsr,signal_ulushi=S.mean()*100)
ths=np.round(np.arange(2.5,3.81,0.05),2); tab=[sig(t) for t in ths]
best=max(tab,key=lambda r:(100-r['I_tur'])-r['II_tur']); print('Youden eng yaxshi',best)
for t in (2.75,3.0,3.1,3.25,3.5): print(sig(t))
OUT['signal']=[sig(t) for t in (2.75,3.0,best['th'],3.25,3.5)]; OUT['youden']=best
# 3) logit modellar
def logit(cols,dd=d):
  q=dd[['Y']+cols].dropna(); X=sm.add_constant(q[cols]); r=sm.Logit(q.Y,X).fit(disp=0,cov_type='cluster',cov_kwds={'groups':dd.loc[q.index,'bank'].astype('category').cat.codes})
  return r,q
L={}
for nm,cols in {'1: MBII':['MBII'],'2: ln(Z)':['lnZ'],'3: MBII + ln(Z)':['MBII','lnZ'],'4: Bloklar':['bA','bL','bD'],'5: MBII + davlat':['MBII','DAVLAT']}.items():
  r,q=logit(cols); pr=r.predict(sm.add_constant(q[cols])); a=auc(q.Y,pr)
  L[nm]=dict(b=r.params.round(3).to_dict(),se=r.bse.round(3).to_dict(),p=r.pvalues.round(4).to_dict(),N=int(r.nobs),pos=int(q.Y.sum()),pseudoR2=round(r.prsquared,3),AUC=round(a,3),AIC=round(r.aic,1))
  print(nm,L[nm])
OUT['logit']=L
# marginal ta'sir: MBII 0,5 ballga pasaysa ehtimol
r,q=logit(['MBII']); b0,b1=r.params; 
for x in (3.75,3.25,2.75,2.5): print('MBII',x,'ehtimol',round(100/(1+np.exp(-(b0+b1*x))),1))
OUT['ehtimol']={x:round(100/(1+np.exp(-(b0+b1*x))),1) for x in (4.0,3.75,3.5,3.25,3.0,2.75,2.5,2.25)}
# 4) vaqtdan tashqari: o'qitish 2021_IV–2022_IV, sinov 2023_I–2024_IV
tr=d[d.t<=2022*4+3]; te=d[d.t>=2023*4]
r,q=logit(['MBII'],tr); pr=r.predict(sm.add_constant(te[['MBII']]))
OUT['oos']=dict(train_N=int(len(tr)),train_pos=int(tr.Y.sum()),test_N=int(len(te)),test_pos=int(te.Y.sum()),AUC_MBII=round(auc(te.Y,pr),3),
  AUC_Z=round(auc(te.Y[te.lnZ.notna()],-te.lnZ.dropna()),3),AUC_S3=round(auc(te.Y,te.S3),3),
  th_train=max([sig(t,tr.MBII,tr.Y) for t in ths],key=lambda r:(100-r['I_tur'])-r['II_tur'])['th'])
OUT['oos']['test_at_train_th']=sig(OUT['oos']['th_train'],te.MBII,te.Y)
print('OOS',OUT['oos'])
# 5) misollar: hodisa yuz bergan banklar indeks dinamikasi
cs={}
for b in ['Mikrokreditbank','Biznesni rivojlantirish banki','Garant bank','Ipoteka-bank','Tibisi bank','Xalq banki']:
  x=p[(p.bank==b)&p.tanlov][['davr','MBII','lnZ','S3','lossq']].round(2); cs[b]=x.values.tolist()
OUT['misollar']=cs
for b,v in cs.items(): print(b,[(r[0],r[1],r[4]) for r in v])
json.dump(OUT,open('ews41.json','w'),ensure_ascii=False,indent=1,default=float)
