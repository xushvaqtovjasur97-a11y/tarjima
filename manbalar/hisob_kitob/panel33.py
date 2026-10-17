# 3.3-§: bank-chorak paneli (2021_IV–2025_IV), panel MBII (8 ko'rsatkich) va Z-score
import pandas as pd, numpy as np, json, os
import mbii as M
B='../markaziy_bank/banklar/'
c=pd.read_csv(B+'cbu_banklar_choraklik.csv').pivot_table(index=['bank','davr'],columns='korsatkich',values='qiymat').reset_index()
s=pd.read_csv(B+'banklar_2021IV_2026II.csv')
MAP={'AGROBANK':'Agrobank','ALOQABANK':'Aloqa bank','ANOR BANK':'Anor bank','APEX BANK':'Apeks bank','ASAKABANK':'Asaka bank',
'ASIA ALLIANCE BANK':'Asia Alliance bank','AVO BANK':'AVO bank',"BANK IPAK YO'LI":'Ipak yo‘li bank','BIZNESNI RIVOJLANTIRISH BANKI':'Biznesni rivojlantirish banki',
'DAVR-BANK':'Davr bank','EUROOSIYO BANKI':'Yevroosiyo banki','GARANT BANK':'Garant bank','HAMKORBANK':'Hamkorbank','HAYOT BANK':'Hayot bank',
'INFINBANK':'Invest Finance bank','IPOTEKA-BANK':'Ipoteka-bank','KAPITALBANK':'Kapital bank',"KDB BANK O'ZBEKISTON":'KDB Bank O‘zbekiston',
'MADAD INVEST BANK':'Madad Invest bank','MIKROKREDITBANK':'Mikrokreditbank',"O'ZMILLIYBANK":'O‘zmilliybank',"O'ZSANOATQURILISHBANK":'O‘zsanoatqurilishbank',
'OCTOBANK':'Oktobank','OPEN BANK':'Open bank','ORIENT FINANS BANK':'Orient Finance bank','POYTAXT BANK':'Poytaxt bank','TBC BANK':'Tibisi bank',
'TENGE BANK':'Tenge bank','TRASTBANK':'Trast bank','TURONBANK':'Turon bank','UNIVERSAL BANK':'Universal bank','UZUM BANK':'Uzum bank',
'XALQ BANKI':'Xalq banki','YANGI BANK':'Yangi bank','ZIRAAT BANK UZBEKISTAN':'Ziraat Bank Uzbekistan'}
s['bank']=s.bank.map(MAP); s=s.dropna(subset=['bank'])
K={'Резервы по кредитам':'RES','Высоколиквидные активы':'HLA','Всего активы':'TA_s','Собственный капитал':'EQ_s','Чистая прибыль':'NP_ytd',
'Операционные расходы':'OPEX','Чистые процентные доходы':'NII','Беспроцентные доходы':'NIIn','Беспроцентные расходы':'NIEn',
'Доходоприносящие активы':'EA','Кредиты и лизинговые операции, брутто':'L_s'}
s=s[s.korsatkich.isin(K)].assign(k=lambda d:d.korsatkich.map(K)).pivot_table(index=['bank','davr'],columns='k',values='mlrd_som').reset_index()
# ma'lum xatolar (TEKSHIRUV.md)
bad=[('Mikrokreditbank','2021_IV'),('Mikrokreditbank','2023_IV'),('Asaka bank','2022_I'),('Asaka bank','2023_I'),('Invest Finance bank','2022_IV'),('Tenge bank','2023_II')]
for b,d in bad: s.loc[(s.bank==b)&(s.davr==d),'RES']=np.nan
s=s[~((s.bank=='O‘zmilliybank')&(s.davr=='2022_IV'))]
p=c.merge(s,on=['bank','davr'],how='left')
p['y']=p.davr.str[:4].astype(int); p['q']=p.davr.str[5:].map({'I':1,'II':2,'III':3,'IV':4}); p['t']=p.y*4+p.q-1
p=p.sort_values(['bank','t'])
# davlat banklari (Ipoteka-bank: OTP Bank ga sotilishi 2023-y. III chorak – shu chorakdan 0)
ST=['O‘zmilliybank','O‘zsanoatqurilishbank','Agrobank','Asaka bank','Xalq banki','Mikrokreditbank','Biznesni rivojlantirish banki','Aloqa bank','Turon bank','Poytaxt bank','Ipoteka-bank']
p['DAVLAT']=p.bank.isin(ST).astype(int); p.loc[(p.bank=='Ipoteka-bank')&(p.t>=2023*4+2),'DAVLAT']=0
g=p.groupby('bank')
def lag(col,k): return g[col].shift(k).where(g['t'].shift(k)==p.t-k)
# chorak sof foydasi (YTD dan)
p['NPq']=np.where(p.q==1,p.NP_ytd,p.NP_ytd-lag('NP_ytd',1))
p['TA_o']=(p.Aktiv+lag('Aktiv',1))/2; p['EQ_o']=(p.Kapital+lag('Kapital',1))/2
# indeks ko'rsatkichlari
p['LEV']=p.Kapital/p.Aktiv*100
p['S3']=p.NPL_ulush
p['NPLnet']=np.where(p.RES>0,(p.NPL-p.RES)/p.Kapital*100,np.nan)
p['HLA_r']=p.HLA/p.TA_s*100
p['LTD']=p.Kredit/p.Depozit*100
p['ROA']=p.NP_ytd*4/p.q/p.Aktiv*100
p['ROE']=p.NP_ytd*4/p.q/p.Kapital*100
p['CI']=p.OPEX/(p.NII+p.NIIn-p.NIEn)*100
p.loc[((p.NII+p.NIIn-p.NIEn)<=0)|(p.OPEX<=0)|(p.CI<10),'CI']=np.nan
P8={'LEV':'LEV','S3':'S3','NPLnet':'NPLnet','HLA':'HLA_r','LTD':'LTD','ROA':'ROA','ROE':'ROE','CI':'CI'}
# panelda kapital/aktiv (umumiy kapital) uchun langarlar 8/11/14/17 (O'zbekiston banklarida 12–17%)
#M.IND=[(k,b,n,dr,[8.0,11.0,14.0,17.0] if n=='LEV' else a) for k,b,n,dr,a in M.IND]
W={b:v for b,v in M.W['taklif'].items() if b!='V'}
def ix(r):
  v={n:(None if pd.isna(r[c_]) else r[c_]) for n,c_ in P8.items()}
  if sum(x is not None for x in v.values())<6: return pd.Series([np.nan]*5)
  m,bs=M.mbii(v,W); return pd.Series([m]+[bs.get(b,np.nan) for b in 'KALD'])
p[['MBII','bK','bA','bL','bD']]=p.apply(ix,axis=1)
# Z-score: (ROA + E/A)/sd(ROA), sd – chorak ROA (yillik) ning 8 chorakli siljuvchi standart chetlanishi
p['ROAq']=p.NPq*4/p.TA_o*100
p['sdROA']=g['ROAq'].transform(lambda x:x.rolling(8,min_periods=4).std())
p['Z']=(p.ROA+p.LEV)/p.sdROA; p['lnZ']=np.log(p.Z.where(p.Z>0))
# regressorlar
p['lnTA']=np.log(p.Aktiv)
p['KR_osish']=(p.Kredit/lag('Kredit',4)-1)*100
p['KR_osish_l4']=g['KR_osish'].shift(4).where(g['t'].shift(4)==p.t-4)
p['KR_osish_l8']=g['KR_osish'].shift(8).where(g['t'].shift(8)==p.t-8)
p['DEP_A']=p.Depozit/p.Aktiv*100
p['KR_A']=p.Kredit/p.Aktiv*100
p['NIM']=p.NII*4/p.q/p.EA*100
mk=pd.DataFrame(json.load(open('makro.json'))).T.reset_index().rename(columns={'index':'davr'})
p=p.merge(mk[['davr','YaIM_osish','Inflyatsiya','Asosiy_stavka','USD_osish','Kredit_YaIM','Kredit_osish','NPL_tizim']],on='davr',how='left')
for k in ['MBII','lnZ','S3']: p[k+'_l1']=p.groupby('bank')[k].shift(1)
p['tanlov']=(p.Aktiv>=1000)&(p.KR_A>=20)&(p.t>=2021*4+3)
p['tanlov']&=p.groupby('bank').tanlov.transform('sum')>=12  # kamida 12 chorak kuzatilgan banklar
p.to_csv('panel33.csv',index=False)
if __name__=='__main__':
  d=p[p.tanlov]
  print(len(d),'kuzatuv',d.bank.nunique(),'bank'); print(d.groupby('davr').bank.count().to_dict())
  print(d[['MBII','lnZ','LEV','S3','NPLnet','HLA_r','LTD','ROA','ROE','CI','NIM','KR_osish','lnTA']].describe().T.round(2).to_string())
