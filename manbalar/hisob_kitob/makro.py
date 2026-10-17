# Makroiqtisodiy omillar, choraklik 2021_IV–2025_IV (Markaziy bank statistik byulleteni, 2026-yil mart)
import openpyxl,re,json,datetime
wb=openpyxl.load_workbook('/home/user/tarjima/manbalar/markaziy_bank/bul2603.xlsx',read_only=True,data_only=True)
def rows(s): return [list(r) for r in wb[s].iter_rows(values_only=True)]
Q=['2021_IV']+[f'{y}_{q}' for y in range(2022,2026) for q in ('I','II','III','IV')]
out={q:{} for q in Q}
# 1.1 YaIM: yil sarlavhalari va chorak ustunlari
R=rows('1.1'); hdr=R[4]; years=[]
for c,x in enumerate(hdr):
  if isinstance(x,str) and re.match(r'20\d\d-yil',x) and x[:4]!='2026': years.append((int(x[:4]),c))
nom=next(r for r in R if r[1] and str(r[1]).startswith('YaIM (nominal'))
real=R[R.index(nom)+1]
for y,c in years:
  for k,q in enumerate(('I','II','III','IV')):
    key=f'{y}_{q}'
    if key in out:
      out[key]['YaIM_nom_kum']=nom[c+k]; out[key]['YaIM_osish']=real[c+k]-100 if isinstance(real[c+k],(int,float)) else None
# choraklik nominal YaIM va oxirgi 4 chorak yig'indisi
cum={}
for y,c in years:
  for k,q in enumerate(('I','II','III','IV')): cum[(y,k)]=nom[c+k]
def qn(y,k): return cum[(y,k)]-(cum[(y,k-1)] if k>0 else 0)
for key in Q:
  y,q=int(key[:4]),('I','II','III','IV').index(key[5:])
  s=0; yy,kk=y,q
  for _ in range(4):
    s+=qn(yy,kk); kk-=1
    if kk<0: yy-=1; kk=3
  out[key]['YaIM_4ch']=s
# 1.2 INI yillik (dekabr/mart/iyun/sentabr)
R=rows('1.2'); yr=None; M={'Mart':'I','Iyun':'II','Sentabr':'III','Dekabr':'IV'}
for r in R:
  if r[0] and re.match(r'20\d\d-yil',str(r[0])): yr=int(str(r[0])[:4]); continue
  if r[0] in M and yr and f'{yr}_{M[r[0]]}' in out: out[f'{yr}_{M[r[0]]}']['Inflyatsiya']=r[5]
# 3.7 USD kursi (chorak oxirgi oyi o'rtachasi)
R=rows('3.7'); yr=None; usd={}
for r in R:
  if r[0] and re.match(r'20\d\d-yil',str(r[0])): yr=int(str(r[0])[:4]); continue
  if r[0] in M and yr: usd[f'{yr}_{M[r[0]]}']=r[1]
for k in Q:
  out[k]['USD']=usd.get(k); y,q=int(k[:4]),k[5:]
  p=usd.get(f'{y-1}_{q}'); out[k]['USD_osish']=(usd[k]/p-1)*100 if p and usd.get(k) else None
# 3.1 asosiy stavka (chorak oxiridagi amaldagi)
R=rows('3.1'); st=[]
for r in R:
  m=re.match(r'\s*(\d\d)\.(\d\d)\.(\d{4})',str(r[1] or ''))
  if m and isinstance(r[2],(int,float)): st.append((datetime.date(int(m[3]),int(m[2]),int(m[1])),r[2]))
st.sort()
QE={'I':(3,31),'II':(6,30),'III':(9,30),'IV':(12,31)}
for k in Q:
  d=datetime.date(int(k[:4]),*QE[k[5:]]); v=[s for dd,s in st if dd<=d]; out[k]['Asosiy_stavka']=v[-1] if v else None
# 5.3.8 iqtisodiyotga kreditlar qoldig'i (keyingi oy 1-sanasi) va 5.1.8 NPL
def by_date(s,col):
  d={}
  for r in rows(s):
    m=re.match(r'01\.(\d\d)\.(\d{4})',str(r[0] or ''))
    if m and isinstance(r[col],(int,float)): d[(int(m[2]),int(m[1]))]=r[col]
  return d
cr=by_date('5.3.8',1); npl=by_date('5.1.8',7)
NX={'I':(0,4),'II':(0,7),'III':(0,10),'IV':(1,1)}
for k in Q:
  a,mo=NX[k[5:]]; key=(int(k[:4])+a,mo)
  out[k]['Kredit_qoldiq']=cr.get(key); out[k]['NPL_tizim']=npl.get(key)
  out[k]['Kredit_YaIM']=cr[key]/out[k]['YaIM_4ch']*100 if cr.get(key) else None
for k in Q:
  a,mo=NX[k[5:]]; prev=cr.get((int(k[:4])-1+a,mo))
  out[k]['Kredit_osish']=(out[k]['Kredit_qoldiq']/prev-1)*100 if prev and out[k]['Kredit_qoldiq'] else None
json.dump(out,open('/home/user/tarjima/manbalar/hisob_kitob/makro.json','w'),indent=1)
for k in Q: print(k,{a:(round(b,1) if isinstance(b,(int,float)) else b) for a,b in out[k].items() if a not in ('YaIM_nom_kum',)})
