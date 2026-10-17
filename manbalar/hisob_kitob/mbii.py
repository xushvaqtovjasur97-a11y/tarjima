# Moliyaviy barqarorlikning integral indeksi (MBII) – 3.1-§ hisob-kitobi
IND=[ # yagona metodika (3.1 va 3.2): kod, blok, nomi, yo'nalish, langarlar (1,2,3,4 ball)
('K1','K','CET1',+1,[4.5,7.0,8.0,10.5]),
('K2','K','CAR',+1,[8.0,10.5,13.0,15.5]),
('K3','K','LEV',+1,[3.0,4.0,6.0,8.0]),
('A1','A','S3',-1,[15.0,10.0,5.0,3.0]),
('A2','A','S2',-1,[35.0,25.0,15.0,10.0]),
('A3','A','NPLnet',-1,[30.0,20.0,10.0,0.0]),
('L1','L','HLA',+1,[5.0,10.0,15.0,20.0]),
('L2','L','LTD',-1,[200.0,150.0,100.0,80.0]),
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
  for k,b,n,dr,a in IND:
    if v.get(n) is not None: bl.setdefault(b,[]).append(score(v[n],dr,a))
  bs={b:sum(s)/len(s) for b,s in bl.items()}
  return sum(w[b]*bs[b] for b in bs)/sum(w[b] for b in bs), bs
def cls(m): return 'barqaror' if m>=3.25 else 'qoniqarli' if m>=2.5 else 'zaif' if m>=1.75 else 'nobarqaror'
DATA={}
if __name__=='__main__':
  for k,v in DATA.items():
    m,bs=mbii(v,W['taklif'])
    print(k,round(m,2),cls(m),{b:round(s,2) for b,s in bs.items()},'| teng',round(mbii(v,W['teng'])[0],2),'CAMEL',round(mbii(v,W['CAMEL'])[0],2))
    print('   ',{n:round(score(v[n],dr,a),2) for c,b,n,dr,a in IND})
