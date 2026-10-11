# Markaziy bank: banklar kesimida NPL va asosiy ko'rsatkichlar, chorak oxiri 2021_IV–2025_IV
import openpyxl,glob,re,json,csv
D='/home/user/tarjima/manbalar/markaziy_bank/banklar/cbu_oylik/'
Q={'01':('IV',-1),'04':('I',0),'07':('II',0),'10':('III',0)}
ALIAS={'Qishloq qurilish bank':'Biznesni rivojlantirish banki','Hi-tech bank':'Hi-Tech bank','KDB bank O‘zbekiston':'KDB Bank O‘zbekiston','Madad invest bank':'Madad Invest bank','Orient Finans bank':'Orient Finance bank','Trastbank':'Trast bank','Eron Soderot bank':'Soderot bank','Eron Soderot bankining ShB':'Soderot bank','O`zagroeksportbank':'O‘zagroeksportbank'}
def norm(n):
  b=re.sub(r'\s+',' ',str(n).strip().replace('’','‘').replace("'",'‘'))
  return ALIAS.get(b,b)
out=[]
for f in sorted(glob.glob(D+'*.xlsx')):
  kind,ym=re.match(r'.*/(\w+)_(\d{4}-\d\d)\.xlsx',f).groups()
  y,m=ym.split('-'); q,dy=Q[m]; davr=f'{int(y)+dy}_{q}'
  ws=openpyxl.load_workbook(f,read_only=True,data_only=True).worksheets[0]
  for r in ws.iter_rows(values_only=True):
    if isinstance(r[0],(int,float)) and r[1]:
      b=norm(r[1])
      if kind=='npl': out.append((davr,b,'Kredit',r[2])); out.append((davr,b,'NPL',r[3])); out.append((davr,b,'NPL_ulush',r[4]*100 if isinstance(r[4],(int,float)) else None))
      else:
        vals=[x for x in r[2:] if x is not None]
        # Aktiv, ulush, Kredit, ulush, Kapital, ulush, Depozit, ulush
        if len(vals)>=8: out.append((davr,b,'Aktiv',vals[0])); out.append((davr,b,'Kapital',vals[4])); out.append((davr,b,'Depozit',vals[6]))
with open('/home/user/tarjima/manbalar/markaziy_bank/banklar/cbu_banklar_choraklik.csv','w',newline='') as fh:
  w=csv.writer(fh); w.writerow(['davr','bank','korsatkich','qiymat']); w.writerows(out)
import collections
c=collections.Counter(d for d,b,k,v in out if k=='NPL_ulush'); print(sorted(c.items()))
print(sorted({b for d,b,k,v in out})[:60])
