# Moliyaviy barqarorlikning integral indeksi (MBII) – 3.1-§ hisob-kitobi
IND=[ # kod, blok, nomi, yo'nalish(+ yuqori yaxshi), langarlar (1,2,3,4 ball)
('K1','K','CET1',+1,[4.5,7.0,8.0,10.5]),
('K2','K','CAR',+1,[8.0,10.5,13.0,15.5]),
('K3','K','LEV',+1,[3.0,4.0,6.0,8.0]),
('A1','A','S3',-1,[15.0,10.0,5.0,3.0]),
('A2','A','S2',-1,[35.0,25.0,15.0,10.0]),
('A3','A','NPLnet',-1,[30.0,20.0,10.0,0.0]),
('L1','L','LCR',+1,[90.0,100.0,120.0,150.0]),
('L2','L','NSFR',+1,[90.0,100.0,110.0,120.0]),
('L3','L','LTD',-1,[200.0,150.0,100.0,80.0]),
('D1','D','ROA',+1,[0.0,0.7,1.4,2.0]),
('D2','D','ROE',+1,[0.0,5.0,10.0,15.0]),
('D3','D','CI',-1,[100.0,80.0,60.0,45.0]),
('V1','V','FXL',-1,[65.0,50.0,35.0,20.0]),
('V2','V','FXLI',-1,[70.0,55.0,40.0,25.0]),
]
W={'taklif':dict(K=.25,A=.30,L=.20,D=.15,V=.10),'teng':dict(K=.2,A=.2,L=.2,D=.2,V=.2),'CAMEL':dict(K=.23,A=.37,L=.15,D=.25,V=0)}
def score(x,d,a):
  if d<0: x,a=-x,[-v for v in a]
  if x<=a[0]: return 1.0
  if x>=a[3]: return 4.0
  for i in range(3):
    if x<=a[i+1]: return 1+i+(x-a[i])/(a[i+1]-a[i])
def mbii(v,w):
  bl={}
  for k,b,n,dr,a in IND: bl.setdefault(b,[]).append(score(v[n],dr,a))
  bs={b:sum(s)/len(s) for b,s in bl.items()}
  return sum(w[b]*bs[b] for b in bs), bs
def cls(m): return 'barqaror' if m>=3.25 else 'qoniqarli' if m>=2.5 else 'zaif' if m>=1.75 else 'nobarqaror'
DATA={
'SQB milliy':dict(CET1=10.21,CAR=17.02,LEV=12.25,S3=2.53,S2=18.44,NPLnet=-8.1,LCR=222.9,NSFR=122.2,LTD=206.5,ROA=1.98,ROE=17.2,CI=36.5,FXL=60.1,FXLI=60.9),
'SQB MHXS':dict(CET1=8.87,CAR=15.32,LEV=11.07,S3=7.48,S2=35.66,NPLnet=7.3,LCR=154.0,NSFR=124.8,LTD=221.0,ROA=1.70,ROE=15.7,CI=41.9,FXL=60.1,FXLI=60.9),
'MKB milliy':dict(CET1=16.39,CAR=18.46,LEV=16.82,S3=4.41,S2=9.0,NPLnet=10.3,LCR=231.9,NSFR=113.0,LTD=208.5,ROA=0.59,ROE=3.5,CI=65.5,FXL=18.4,FXLI=31.4),
'MKB MHXS':dict(CET1=18.24,CAR=19.96,LEV=18.72,S3=12.34,S2=25.3,NPLnet=26.4,LCR=231.9,NSFR=113.0,LTD=324.4,ROA=0.15,ROE=0.81,CI=88.0,FXL=18.4,FXLI=31.4),
}
if __name__=='__main__':
  for k,v in DATA.items():
    m,bs=mbii(v,W['taklif'])
    print(k,round(m,2),cls(m),{b:round(s,2) for b,s in bs.items()},'| teng',round(mbii(v,W['teng'])[0],2),'CAMEL',round(mbii(v,W['CAMEL'])[0],2))
    print('   ',{n:round(score(v[n],dr,a),2) for c,b,n,dr,a in IND})
