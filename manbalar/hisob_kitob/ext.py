import sys,pickle,re
S=sys.argv[1]; data=pickle.load(open(S+'/panel/data.pkl','rb'))
def num(v):
  try: return float(v)
  except: return None
def find(rows,pat,col,start=0):
  for r in rows[start:]:
    lab=' '.join(str(x) for x in r[:3] if isinstance(x,str))
    if re.search(pat,lab,re.I):
      v=num(r[col]) if col<len(r) else None
      if v is not None: return v
  return None
res={}
for k,sh in sorted(data.items()):
  en=sh.get('Economic Normatives',[]); bal=sh.get('Balance',[]); ac=sh.get('Asset Classification',[])
  d={}
  d['CAR']=find(en,r'Регулятив капиталнинг монандлик|Коэффициент достаточности капитала$',8)
  d['T1']=find(en,r'I даражали капиталнинг монандлик|достаточности Капитала Уровня 1',8)
  d['CET1']=find(en,r'асосий капиталнинг монандлик|основного Капитала Уровня 1',8)
  d['LEV']=find(en,r'Левераж',8)
  d['LCR']=find(en,r'қоплаш меъёри|покрытия ликвидности',8)
  d['NSFR']=find(en,r'Соф барқарор|чистого стабильного',8)
  d['HLA']=find(en,r'Юқори ликвидли активларнинг жами|высоколиквидных активов в общих',8)
  d['TA']=find(bal,r'Жами активлар|Итого активов',2)
  d['EQ']=find(bal,r'Жами хусусий капитал|Итого собственный капитал|Всего собственного капитала',2)
  d['NP']=find(sh.get('Income',[]),r'СОФ ФОЙДА \(ЗАРАР\)|ЧИСТАЯ ПРИБЫЛЬ',2)
  d['LOANG']=find(bal,r'Кредит ва лизинг операциялари, брутто|Кредиты и лизинговые операции, брутто',2)
  d['DEPD']=find(bal,r'Талаб қилиб олингунча|Депозиты до востребования',2)
  d['DEPT']=find(bal,r'Муддатли депозитлар|Срочные депозиты',2)
  for r in ac:
    lab=' '.join(str(x) for x in r[:3] if isinstance(x,str))
    if re.search(r'Жами кредит ва лизинг суммаси|Итого кредит',lab):
      d['ACtot']=num(r[2]); d['STD']=num(r[4]); d['SUB']=num(r[8]); d['UNS']=num(r[12]); d['DBT']=num(r[16]); d['LOSS']=num(r[20]); d['RES']=num(r[24]); break
  res[k]=d
  print(k,{a:(round(b,4) if b and abs(b)<10 else (round(b/1e6,1) if b else b)) for a,b in d.items()})
pickle.dump(res,open(S+'/panel/res.pkl','wb'))
