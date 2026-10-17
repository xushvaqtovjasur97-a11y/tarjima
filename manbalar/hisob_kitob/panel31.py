# MBII-milliy uchun 14 ko'rsatkichni prudensial shakllardan hisoblash (MKB 2016-2025, SQB 2020-2025)
import pickle,re,json,sys
SP='/tmp/claude-0/-home-user-tarjima/315c132a-88b7-5e49-bb67-7f6a05e43e7d/scratchpad'
P=pickle.load(open(SP+'/p/panel/data.pkl','rb'))
O=pickle.load(open(SP+'/mk1619/d.pkl','rb'))
def num(v):
  try: return float(v)
  except: return None
def lab(r): return ' '.join(str(x) for x in r[:4] if isinstance(x,str)).strip()
def find(rows,pat,col=2):
  for r in rows:
    if re.search(pat,lab(r),re.I):
      v=num(r[col]) if col<len(r) else None
      if v is not None: return v,r
  return None,None
def sheet(sh,*names):
  for n in names:
    for k in sh:
      if k.strip().lower()==n.lower(): return sh[k]
  return []
def get(b,y):
  sh=P[(b,y)] if (b,y) in P else O[y]
  en=sheet(sh,'Economic Normatives'); bal=sheet(sh,'Balance'); inc=sheet(sh,'Income')
  ac=sheet(sh,'Asset Classification','Loan Classification'); cc=sheet(sh,'Bank Capital Calc','Bank Capital Calc.')
  r={}
  r['CAR']=find(en,r'^1\.1')[0] if False else find(en,r'Регулятив капиталнинг монандлик|Коэффициент достаточности капитала$|достаточности капитала',8)[0]
  t1=find(en,r'I даражали капиталнинг монандлик|достаточности Капитала Уровня 1',8)[0]
  c1=find(en,r'асосий капиталнинг монандлик|основного Капитала Уровня 1',8)[0]
  r['CET1']=c1 if c1 is not None else t1
  r['LEV']=find(en,r'Левераж',8)[0]
  r['LCR']=find(en,r'қоплаш меъёри|покрытия ликвидности',8)[0]
  r['NSFR']=find(en,r'Соф барқарор|чистого стабильного',8)[0]
  for k in ('CAR','CET1','LEV','LCR','NSFR'):
    if r[k] is not None: r[k]*=100
  TA=find(bal,r'Жами активлар|Итого активов|ВСЕГО АКТИВОВ|Всего активов')[0]
  EQ=find(bal,r'Жами хусусий капитал|Итого собственн|Всего собственного капитала|ВСЕГО СОБСТВЕННОГО')[0]
  L,Lr=find(bal,r'Кредит ва лизинг операциялари, брутто|Кредиты и лизинговые операции, брутто|Кредиты и лизинг, брутто')
  TL,TLr=find(bal,r'Жами мажбуриятлар|Итого обязательств|ВСЕГО ОБЯЗАТЕЛЬСТВ|Всего обязательств')
  DD=find(bal,r'Талаб қилиб олингунча|Депозиты до востребования')[0] or 0
  DT=find(bal,r'Муддатли депозитлар|Срочные депозиты')[0] or 0
  DS=find(bal,r'Жамғарма депозитлари|Сберегательные депозиты')[0] or 0
  RS=find(bal,r'^8\.2')[0] or 0
  ST=find(bal,r'^13\.1 Стандарт|^31\.2|по стандартным')[0] or 0
  r['FXL']=num(Lr[4])/L*100 if Lr else None
  r['FXLI']=num(TLr[4])/TL*100 if TLr and num(TLr[4]) is not None else None
  r['LTD']=L/(DD+DT+DS)*100
  # tasnif
  for row in ac:
    if re.search(r'Жами кредит ва лизинг суммаси|Итого кредит|ВСЕГО КРЕДИТ|Всего кредит',lab(row)):
      tot=num(row[2]); sub=num(row[8]); npl=sum(num(row[i]) or 0 for i in (12,16,20)); break
  else: tot=sub=npl=None
  r['S3']=npl/tot*100 if tot else None; r['S2']=sub/tot*100 if tot else None
  RC=find(cc,r'^V\b|ЖАМИ РЕГУЛЯТИВ КАПИТАЛ|ИТОГО РЕГУЛЯТИВНЫЙ|РЕГУЛЯТИВНЫЙ КАПИТАЛ')[0]
  r['NPLnet']=(npl-RS-ST)/RC*100 if (RC and npl is not None) else None
  NP=None
  for row in inc:
    if lab(row).startswith('10.'): NP=num(row[2]); break
  ti=find(inc,r'^к\. Жами фоизли даромадлар|^л\. Итого процентных доходов|Итого процентных доходов$|Всего процентных доходов')[0]
  te=find(inc,r'^к\. Жами фоизли харажатлар|^л\. Итого процентных расходов$|Итого процентных расходов$|Всего процентных расходов$')[0]
  ni=find(inc,r'Жами фоизсиз даромадлар|Итого беспроцентных доходов|Всего беспроцентных доходов')[0]
  rr=find(inc,r'захираларни қайтарилиши|захираларнинг қайтарил|Возврат резервов')[0] or 0
  ne=find(inc,r'Жами фоизсиз харажатлар|Итого беспроцентных расходов|Всего беспроцентных расходов')[0]
  op=find(inc,r'Жами операцион харажатлар|Итого операционных расходов|Всего операционных расходов')[0]
  try: r['CI']=op/((ti-te)+(ni-rr)-ne)*100
  except Exception: r['CI']=None
  r['_TA']=TA; r['_EQ']=EQ; r['_NP']=NP; r['_L']=L; r['_RC']=RC
  return r
OUT={}
for b,ys in (('MK',range(2020,2026)),('SQ',range(2020,2026))):  # 2020 faqat o'rtacha qiymatlar uchun
  prev=None
  for y in ys:
    r=get(b,y)
    if prev and r['_NP'] is not None:
      r['ROA']=r['_NP']/((r['_TA']+prev['_TA'])/2)*100
      r['ROE']=r['_NP']/((r['_EQ']+prev['_EQ'])/2)*100 if (r['_EQ'] and prev['_EQ']) else None
    else:
      r['ROA']=r['_NP']/r['_TA']*100 if r['_NP'] else None; r['ROE']=r['_NP']/r['_EQ']*100 if (r['_NP'] and r['_EQ']) else None
    OUT[f'{b}{y}']=r; prev=r
    print(b,y,{k:(round(v,2) if isinstance(v,float) and abs(v)<1e5 else (round(v/1e6,1) if isinstance(v,float) else v)) for k,v in r.items()})
json.dump({k:{a:b for a,b in v.items()} for k,v in OUT.items()},open('/home/user/tarjima/manbalar/hisob_kitob/panel31.json','w'),indent=1)
