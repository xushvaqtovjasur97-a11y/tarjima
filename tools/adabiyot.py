# Foydalanilgan adabiyotlar ro'yxatini barcha fn*.py lug'atlari va matnlardagi havolalardan yig'ish
import re,glob,json
R={}
for f in sorted(glob.glob('/home/user/tarjima/tools/fn*.py')):
  code=open(f).read(); i=code.find('\ndef ref')
  if i<0: continue
  blk=code[:i]; blk=blk[blk.find('R={}'):] if 'R={}' in blk else blk
  blk=re.sub(r"for f in \[.*?\n(?=R\.update|\S)", "", blk, flags=re.S)
  ns={'R':{}}; exec(blk,ns); R.update(ns['R'])
used=set()
for f in glob.glob('/home/user/tarjima/paragraflar/*.src.md'):
  used|=set(k for k,v in re.findall(r'%([A-Z][A-Z0-9]*):([^%]+)%',open(f).read()) if k!='AUTHOR')
for f in glob.glob('/home/user/tarjima/tools/fn*.py'):
  used|=set(re.findall(r"ref\('([A-Z][A-Z0-9]*)'",open(f).read()))
items={k:R[k][0].strip() for k in used if k in R}
# 2.1-§ (manba matni to'g'ridan-to'g'ri snoskalarda)
extra=[]
for line in open('/home/user/tarjima/paragraflar/2.1-paragraf.md'):
  m=re.match(r'^\[\^\d+\]: (.*)$',line.strip())
  if m:
    t=m.group(1)
    if t.startswith('Muallif tomonidan') or t.startswith('Tadqiqot natijasida'): 
      t=t.split(': ',1)[1] if ': ' in t else ''
    for part in re.split(r';\s+(?=[A-ZА-ЯЎҚҒҲO‘«"])',t):
      part=part.strip().rstrip('.')
      if len(part)>25: extra.append(part+'.')
json.dump({'keyed':items,'extra':extra},open('/home/user/tarjima/manbalar/adabiyot/xom.json','w'),ensure_ascii=False,indent=1)
print(len(items),len(extra))

# ---- guruhlash va ro'yxatni yozish
import unicodedata
T=lambda k: R[k][0].strip().rstrip('.')+'.'
LOCAL=re.search(r"'LOCAL':J\((.*?)\n \),?\n?\}",open('/home/user/tarjima/tools/fn_kirish.py').read(),re.S)
loc_src=open('/home/user/tarjima/tools/fn_kirish.py').read()
local=[m.strip().rstrip('.')+'.' for m in re.findall(r"\n '([^']+)',",loc_src)]
G=[]
G.append(('I. Normativ-huquqiy hujjatlar',[
 'O‘zbekiston Respublikasining 2019-yil 11-noyabrdagi O‘RQ-582-son «O‘zbekiston Respublikasining Markaziy banki to‘g‘risida»gi Qonuni // Qonunchilik ma’lumotlari milliy bazasi.',
 T('ZRU580'),T('ORQ1070'),T('PF5992'),
 'O‘zbekiston Respublikasi Prezidentining 2022-yil 28-yanvardagi PF-60-son «2022–2026-yillarga mo‘ljallangan Yangi O‘zbekistonning taraqqiyot strategiyasi to‘g‘risida»gi Farmoni // Qonunchilik ma’lumotlari milliy bazasi.',
 T('PQ4611'),T('N3697'),
 'Tijorat banklari kapitalining monandligiga qo‘yiladigan talablar to‘g‘risidagi nizom (O‘zbekiston Respublikasi Markaziy banki Boshqaruvining qarori, ro‘yxat raqami 2693, 06.07.2015). – https://lex.uz.',
 T('N2709'),T('N2696'),T('ROADMAP')[:-1]+' – Toshkent, 17.11.2025.',
],False))
G.append(('II. Xalqaro standartlar, xalqaro tashkilotlar va xorijiy nazorat organlarining hujjatlari',[T(k) for k in
 ['B3GRF','B3LCR','B3NSFR','B3FIN','DSIB','BCBS15','BCBS17','BCBS23','BCBSST','SVB','IFRS9','IMFFSI','IMF145','IMF227','IMF189','IMF','GFSR','CIH','CIH07','SCH','HKS','NK','WBFSA','WBPAD','OECDCI','EBA25','SREP','CRR3','PRA','ECBFS','FEDLFI','FFIEC','FEDST','B3US','PCA','CNLAW','CNCAP','NFRA','CBRCCYB','CBRNKL','CBRNNKL']]+
 ['International Monetary Fund. Uzbekistan: Staff Concluding Statement of the 2026 Article IV Mission. – 13.04.2026. – https://www.imf.org/en/news/articles/2026/04/13/mcs041326-uzbekistan.'],True))
vf=T('VOTF'); vparts=[p.strip().rstrip('.')+'.' for p in vf.split('; ')]
G.append(('III. Monografiyalar, darsliklar va o‘quv qo‘llanmalar',[T(k) for k in ['KS','BERD','LM','LAV','AFD','KARM','SOK','SHAT','FET','CAR','HEN','PEL','TAR','HE','IZZI','LANG','MOD','BALT','WOOL','HOS']]+vparts,True))
G.append(('IV. Ilmiy maqolalar',[T(k) for k in ['AB','ABAD','ABND','AHS','ATH','BETZ','BG','BHOL','BJP','BOR','BP','CG','CMR','COC','CORN','CP','CROCK','DAVK','DD','DELONG','DKD','DKH','FAW','FNW','GN','HAUS','HL','IL','INS','KLR','KR','LL','LS','LV','LVM','MERT','MIN','MISH','NIC','ROY','SHAN','SST','SV','SY','THAK','TIAN','TUT','SHUB','TIM','LAR']],True))
diss=local[:]
for k in ['SATT','SHAR','BOZ','SAD','MAD','ABD']:
  t=T(k); key=t.split(' ')[0]
  if not any(d.startswith(key) for d in diss): diss.append(t)
diss=[d for d in diss if 'darslik' not in d]
G.append(('V. Dissertatsiyalar va avtoreferatlar',diss,True))
G.append(('VI. Statistik ma’lumotlar, hisobotlar va internet manbalari',['O‘zbekiston Respublikasi Markaziy banki. Statistik byulleten: 2021-yil. – Toshkent, 2022. – https://cbu.uz/uz/statistics/buleten/.','O‘zbekiston Respublikasi Markaziy banki. Statistik byulleten: 2026-yil mart. – Toshkent, 2026. – https://cbu.uz/uz/statistics/buleten/.']+[T(k) for k in ['CBU22','CBU23','CBU24','CBU25','CBUAR23','CBUBS','SNS','SQBREP','MKREP','SQBIFRS','SQBIFRS21','SQBIFRS23','MKIFRS21','MKIFRS23','MKIFRS25','BRBIFRS','XALQIFRS','XALQIFRS22','IPOIFRS','IPOIFRS23','KILDE']]+[
 'The Central Bank of the Republic of Uzbekistan. Financial Stability Report for H1 2025. – Tashkent, 2025. – https://cbu.uz.',
 'The Financial Stability Report for 2024 has been published // The Central Bank of the Republic of Uzbekistan. – https://cbu.uz/en/press_center/releases/2489245/.',
 'Countercyclical capital buffer // The Central Bank of the Republic of Uzbekistan. – https://cbu.uz/en/financial-stability/macroprudential-policy-tools/.',
 '2026-yil uchun mahalliy tizimli ahamiyatga molik banklar belgilandi // O‘zbekiston Respublikasi Markaziy banki. – https://cbu.uz.',
 'https://www.bis.org','https://www.imf.org','https://www.worldbank.org','https://www.ifrs.org','https://cbu.uz','https://lex.uz'],False))
def skey(s):
  s=s.lstrip('«"„ ')
  cyr=bool(re.match(r'[А-Яа-яЁёЎўҚқҒғҲҳ]',s))
  return (1 if cyr else 0, unicodedata.normalize('NFKD',s).lower())
out=['## FOYDALANILGAN ADABIYOTLAR RO‘YXATI','']; n=0; seen=set()
for title,items,srt in G:
  items=[i for i in items if not (i in seen or seen.add(i))]
  if srt: items=sorted(items,key=skey)
  out.append(f'**{title}**'); out.append('')
  for it in items:
    n+=1; out.append(f'{n}\\. {it}'); out.append('')
open('/home/user/tarjima/paragraflar/Adabiyotlar.md','w').write('\n'.join(out)); print('jami',n)
