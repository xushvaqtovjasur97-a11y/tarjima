import sys,glob,openpyxl,xlrd,pickle
S=sys.argv[1]
MK=S+'/r1625/Микрокредитбанк-31.12.2016-31.12.2025 — копия'
SQ=S+'/r2025b/Ўзсаноатқурилишбанк-31.12.2020-31.12.2025 — копия'
WANT=('balance','баланс','economic','иқт','asset classification','loan classification','bank capital calc','liquidity analysis','income')
def sheets(path):
  out={}
  if path.endswith('.xls'):
    wb=xlrd.open_workbook(path)
    for sh in wb.sheets():
      if any(w in sh.name.strip().lower() for w in WANT):
        out[sh.name.strip()]=[[sh.cell_value(r,c) for c in range(min(sh.ncols,26))] for r in range(min(sh.nrows,140))]
  else:
    import shutil,tempfile,os
    tmp=tempfile.mktemp(suffix='.xlsx',dir=S+'/panel'); shutil.copy(path,tmp)
    import msoffcrypto
    with open(tmp,'rb') as fh:
      of=msoffcrypto.OfficeFile(fh); enc=of.is_encrypted()
      if enc:
        of.load_key(password='VelvetSweatshop'); tmp2=tmp+'d.xlsx'
        with open(tmp2,'wb') as o: of.decrypt(o)
    if enc: tmp=tmp2
    wb=openpyxl.load_workbook(tmp,data_only=True,read_only=True)
    for ws in wb.worksheets:
      if any(w in ws.title.strip().lower() for w in WANT):
        out[ws.title.strip()]=[list(r) for r in ws.iter_rows(max_row=140,max_col=26,values_only=True)]
  return out
files={}
for y in range(2020,2026):
  d={2022:'2022-йўқ'}.get(y,str(y))
  files[('MK',y)]=[f for f in glob.glob(MK+'/'+d+'/*') if 'crs005' in f and '~$' not in f and '(3)' not in f]
  files[('SQ',y)]=[f for f in glob.glob(SQ+'/'+str(y)+'/*') if 'crs003' in f and '~$' not in f]
data={}
for k,fl in sorted(files.items()):
  m={}
  for f in fl: m.update(sheets(f))
  data[k]=m; print(k,len(fl),sorted(m.keys()))
pickle.dump(data,open(S+'/panel/data.pkl','wb'))
