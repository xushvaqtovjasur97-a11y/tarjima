# Kamchiliklar jadvali (23, 25, 27-bandlar): robastlik tekshiruvlari
# 1) bank bo'yicha klasterlangan bootstrap AUC; 2) muqobil hodisa ta'riflari; 3) vaznlar bo'yicha AUC; 4) Wilson oraliqlari
import pandas as pd, numpy as np, json, warnings; warnings.filterwarnings('ignore')
from scipy import stats
import mbii as M
OUT={}
p=pd.read_csv('panel33.csv').sort_values(['bank','t']); g=p.groupby('bank')
p['lossq']=-p.NPq/g.Kapital.shift(1)*100
def target(p,th):
  ev=(p.lossq>=th).astype(int); gg=ev.groupby(p.bank)
  Y=pd.concat([gg.shift(-k).where(p.groupby('bank').t.shift(-k)==p.t+k) for k in range(1,5)],axis=1).max(axis=1)
  return ev,Y
def auc(y,s):
  y=np.asarray(y); s=np.asarray(s); pos=s[y==1]; neg=s[y==0]
  return (np.greater.outer(pos,neg).sum()+0.5*np.equal.outer(pos,neg).sum())/(len(pos)*len(neg))
def cboot(d,cols,B=2000,seed=7):
  # bank bo'yicha klasterlangan bootstrap: banklar qaytariladigan tanlov bilan olinadi
  rng=np.random.default_rng(seed); banks=d.bank.unique(); G={b:d[d.bank==b] for b in banks}; r={c:[] for c in cols}; dif=[]
  for _ in range(B):
    s=pd.concat([G[b] for b in rng.choice(banks,len(banks))])
    if 0<s.Y.sum()<len(s):
      a={c:auc(s.Y,s[c]) for c in cols}
      for c in cols: r[c].append(a[c])
      dif.append(a[cols[0]]-a[cols[1]])
  return {c:np.percentile(r[c],[2.5,97.5]).round(3).tolist() for c in cols}, np.percentile(dif,[2.5,97.5]).round(3).tolist(), round(float(np.mean(np.array(dif)<=0)),3)
def sample(th):
  q=p.copy(); q['EV'],q['Y']=target(q,th)
  d=q[q.tanlov&(q.t<=2024*4+3)&(q.EV==0)].dropna(subset=['Y','MBII']).copy()
  d['rMBII']=-d.MBII; d['rS3']=d.S3; d['rZ']=-d.lnZ
  return q,d
# asosiy ta'rif (5%) va muqobil ta'riflar
for th in (5,3,10):
  q,d=sample(th)
  ci,dif,pdif=cboot(d,['rMBII','rS3'])
  dz=d[d.lnZ.notna()]; ciz,_,_=cboot(dz,['rMBII','rZ'])
  OUT[f'hodisa_{th}']=dict(hodisalar=int(q[q.tanlov].EV.sum()),banklar=int(q[q.tanlov&(q.EV==1)].bank.nunique()),N=int(len(d)),pos=int(d.Y.sum()),
     AUC_MBII=round(auc(d.Y,d.rMBII),3),AUC_S3=round(auc(d.Y,d.rS3),3),AUC_Z=round(auc(dz.Y,dz.rZ),3),
     CI_MBII=ci['rMBII'],CI_S3=ci['rS3'],CI_Z=ciz['rZ'],farq_MBII_S3_CI=dif,farq_manfiy_ulush=pdif)
  print(th,OUT[f'hodisa_{th}'])
# bir-birini qoplamaydigan oynalar: har bir bank uchun faqat IV chorak kuzatuvlari
q,d=sample(5); d4=d[d.q==4]
ci,dif,pdif=cboot(d4,['rMBII','rS3'])
OUT['IV_chorak']=dict(N=int(len(d4)),pos=int(d4.Y.sum()),AUC_MBII=round(auc(d4.Y,d4.rMBII),3),AUC_S3=round(auc(d4.Y,d4.rS3),3),CI_MBII=ci['rMBII'],CI_S3=ci['rS3'])
print('IV',OUT['IV_chorak'])
# 3) vaznlar bo'yicha AUC (panel indeksi: K, A, L, D bloklari)
V=json.load(open('vazn32.json'))['vaznlar']
def mb(w): 
  s=w['K']*d.bK+w['A']*d.bA+w['L']*d.bL+w['D']*d.bD; return s/(w['K']+w['A']+w['L']+w['D'])
chk=np.corrcoef(mb({'K':.25,'A':.30,'L':.20,'D':.15}),d.MBII)[0,1]; print('tekshiruv: panel MBII bilan korrelyatsiya',round(chk,4))
OUT['AUC_vaznlar']={col:round(auc(d.Y,-mb({b:V[col][b] for b in 'KALD'})),3) for col in V}
print(OUT['AUC_vaznlar'])
# 4) sinflar bo'yicha hodisa chastotasi va Wilson 95% oraliqlari
def wilson(k,n,z=1.96):
  if n==0: return None
  ph=k/n; den=1+z*z/n; c=(ph+z*z/(2*n))/den; h=z*np.sqrt(ph*(1-ph)/n+z*z/(4*n*n))/den; return [round((c-h)*100,1),round((c+h)*100,1)]
d['sinf']=d.MBII.apply(M.cls); W={}
for s in ['barqaror','qoniqarli','zaif','nobarqaror']:
  x=d[d.sinf==s]; W[s]=dict(n=int(len(x)),k=int(x.Y.sum()),ulush=round(x.Y.mean()*100,1) if len(x) else None,wilson=wilson(int(x.Y.sum()),len(x)),banklar=int(x.bank.nunique()))
OUT['wilson']=W; print(W)
# Fisher aniq testi: I va II, II va III sinflar farqi
f12=stats.fisher_exact([[W['barqaror']['k'],W['barqaror']['n']-W['barqaror']['k']],[W['qoniqarli']['k'],W['qoniqarli']['n']-W['qoniqarli']['k']]])[1]
f23=stats.fisher_exact([[W['qoniqarli']['k'],W['qoniqarli']['n']-W['qoniqarli']['k']],[W['zaif']['k'],W['zaif']['n']-W['zaif']['k']]])[1]
OUT['fisher']={'I_II':round(f12,4),'II_III':round(f23,5)}; print(OUT['fisher'])
json.dump(OUT,open('robust_kam.json','w'),ensure_ascii=False,indent=1,default=float)
# MBII va Z-score AUC farqi (klasterlangan bootstrap)
q,d=sample(5); dz=d[d.lnZ.notna()]
ci,dif,pdif=cboot(dz,['rMBII','rZ'])
OUT['farq_MBII_Z']=dict(N=int(len(dz)),pos=int(dz.Y.sum()),AUC_MBII=round(auc(dz.Y,dz.rMBII),3),AUC_Z=round(auc(dz.Y,dz.rZ),3),CI=dif,manfiy_ulush=pdif)
print('MBII-Z',OUT['farq_MBII_Z'])
json.dump(OUT,open('robust_kam.json','w'),ensure_ascii=False,indent=1,default=float)
