# Kamchiliklar jadvali (36-band): ssenariylarga kapital kanalini qo'shish – zaxira ajratmalari → foyda → kapital
# Soddalashtirilgan yuqori baho: ROA pasayishi (2025-yilga nisbatan) yillar bo'yicha yig'ilib kapitalni kamaytiradi
import numpy as np, pandas as pd, json, io, contextlib
src=open('fc42.py').read()
cut=src.split("A=pd.concat([base,F]); A['zona']=A.MBII.apply(zona)")[0]
ns={}
with contextlib.redirect_stdout(io.StringIO()): exec(cut,ns)
b25=ns['b25']; SC=ns['SC']; SYS=ns['SYS']; S25=ns['S25']; mb=ns['mb']; BR=ns['BR']; zona=ns['zona']
res=[]
for sc in SC:
  for j,y,k in ((3,2026,1),(7,2027,2),(11,2028,3)):
    dv=SYS[sc][j]-S25
    for b,r in b25.iterrows():
      s3=max(r.S3*(1+dv/S25),0.0); roa=r.ROA+BR*(s3-r.S3); roe=roa*r.Aktiv/r.Kapital
      cov=min(r.RES/r.NPL,1.0) if (r.RES==r.RES and r.NPL>0) else 0.5
      npln=r.NPLnet+(s3-r.S3)*r.Kredit/r.Kapital*(1-cov) if r.NPLnet==r.NPLnet else np.nan
      v={'LEV':r.LEV,'S3':s3,'NPLnet':npln,'HLA':r.HLA_r,'LTD':r.LTD,'ROA':roa,'ROE':roe,'CI':r.CI}
      m0,_=mb(v)
      K2=r.Kapital+min(0,roa-r.ROA)/100*r.Aktiv*k
      v2=dict(v,LEV=K2/r.Aktiv*100,ROE=roa*r.Aktiv/K2,NPLnet=npln*r.Kapital/K2 if npln==npln else npln)
      m1,_=mb(v2)
      res.append(dict(ssenariy=sc,yil=y,bank=b,MBII=m0,MBII_kapital=m1,zona=zona(m0),zona_kapital=zona(m1)))
R=pd.DataFrame(res); OUT={}
for (sc,y),g in R.groupby(['ssenariy','yil']):
  OUT[f'{sc}|{y}']=dict(max_farq=round(float((g.MBII-g.MBII_kapital).max()),3),qizil=int((g.zona=='qizil').sum()),qizil_kapital=int((g.zona_kapital=='qizil').sum()),
     zona_ozgargan=g[g.zona!=g.zona_kapital][['bank','MBII','MBII_kapital']].round(3).values.tolist())
print(json.dumps(OUT,ensure_ascii=False,indent=0))
json.dump(OUT,open('kapkanal_kam.json','w'),ensure_ascii=False,indent=1)
