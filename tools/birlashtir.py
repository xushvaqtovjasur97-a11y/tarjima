# Titul + mundarija + barcha yozilgan bo'limlarni bitta docx ga yig'ish (LibreOffice UNO)
import uno, time, subprocess, os
from com.sun.star.beans import PropertyValue
P='/home/user/tarjima/paragraflar/'
OUT='/home/user/tarjima/Dissertatsiya_Maxmudov_O.T.docx'
ITEMS=[('H1','KIRISH',P+'Kirish.docx'),
('B','I BOB. TIJORAT BANKLARI MOLIYAVIY BARQARORLIGINI BAHOLASHNING NAZARIY ASOSLARI VA XALQARO STANDARTLARI',None),
(None,None,P+'1.1-paragraf.docx'),(None,None,P+'1.2-paragraf.docx'),(None,None,P+'1.3-xorijiy-tajriba.docx'),(None,None,P+'I_bob_xulosalar.docx'),
('B','II BOB. O‘ZBEKISTON BANK TIZIMIDA MOLIYAVIY BARQARORLIKNI BAHOLASH AMALIYOTINING XALQARO STANDARTLARGA MUVOFIQLIGI TAHLILI',None),
(None,None,P+'2.1-paragraf.docx'),(None,None,P+'2.2-paragraf.docx'),(None,None,P+'2.3-paragraf.docx'),(None,None,P+'II_bob_xulosalar.docx'),
('B','III BOB. XALQARO STANDARTLARGA ASOSLANGAN INTEGRAL BAHOLASH METODIKASI VA UNING EKONOMETRIK ASOSLANISHI',None),
(None,None,P+'3.1-paragraf.docx'),(None,None,P+'3.2-paragraf.docx'),(None,None,P+'3.3-paragraf.docx')]
def pv(n,v): x=PropertyValue(); x.Name=n; x.Value=v; return x
proc=subprocess.Popen(['soffice','--headless','--invisible','--norestore','--accept=socket,host=localhost,port=2002;urp;'])
ctx=None
for _ in range(60):
  try:
    lc=uno.getComponentContext(); r=lc.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver',lc)
    ctx=r.resolve('uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext'); break
  except Exception: time.sleep(1)
smgr=ctx.ServiceManager; desk=smgr.createInstanceWithContext('com.sun.star.frame.Desktop',ctx)
doc=desk.loadComponentFromURL(uno.systemPathToFileUrl('/tmp/claude-0/merge_titul.docx'),'_blank',0,(pv('Hidden',True),))
T=doc.Text
from com.sun.star.style.BreakType import PAGE_BEFORE
from com.sun.star.text.ControlCharacter import PARAGRAPH_BREAK
def newpara(style=None,text='',pagebreak=True):
  c=T.createTextCursorByRange(T.End)
  T.insertControlCharacter(c,PARAGRAPH_BREAK,False)
  c=T.createTextCursorByRange(T.End)
  if style: c.ParaStyleName=style
  else: c.ParaStyleName='Standard'
  if pagebreak: c.BreakType=PAGE_BEFORE
  if text: T.insertString(c,text,False)
  return c
prevB=False
for kind,title,path in ITEMS:
  if kind=='B':
    prevB=True
    c=newpara('Heading 1',title)
    c.gotoStartOfParagraph(False); c.gotoEndOfParagraph(True); c.ParaAdjust=3  # center
    continue
  c=newpara(None,'',pagebreak=not prevB); prevB=False
  c.insertDocumentFromURL(uno.systemPathToFileUrl(path),())
# KIRISH va bob xulosalari sarlavhalari: KIRISH -> Heading 1
en=T.createEnumeration()
while en.hasMoreElements():
  p=en.nextElement()
  if not p.supportsService('com.sun.star.text.Paragraph'): continue
  s=p.getString().strip()
  if s=='KIRISH' and p.ParaStyleName.startswith('Heading'): p.ParaStyleName='Heading 1'
# bo'sh sahifa yaratuvchi bo'sh paragraflarni olib tashlash: kiritilgan hujjatlar oldidagi bo'sh paragraf o'rniga keyingi paragrafga sahifa uzilishi
# mundarija
sd=doc.createSearchDescriptor(); sd.SearchString='{{TOC}}'; f=doc.findFirst(sd)
idx=doc.createInstance('com.sun.star.text.ContentIndex')
idx.CreateFromOutline=True; idx.Level=2; idx.Title=''
f.setString('')
T.insertTextContent(f,idx,False)
idx.update()
doc.refresh()
idx.update()
# mundarija uslublari va sahifa raqamlari
ps=doc.StyleFamilies.getByName('ParagraphStyles')
for nm,ind in (('Contents 1',0),('Contents 2',500)):
  if ps.hasByName(nm):
    st=ps.getByName(nm); st.CharFontName='Times New Roman'; st.CharHeight=13; st.ParaTopMargin=0; st.ParaBottomMargin=60; st.ParaLeftMargin=ind
pgs=doc.StyleFamilies.getByName('PageStyles')
for i in range(pgs.Count):
  pg=pgs.getByIndex(i)
  if pg.Name in ('First Page',): continue
  try:
    pg.FooterIsOn=True; ft=pg.FooterText; ft.setString('')
    fld=doc.createInstance('com.sun.star.text.TextField.PageNumber'); fld.NumberingType=4; fld.SubType=uno.Enum('com.sun.star.text.PageNumberType','CURRENT')
    cur=ft.createTextCursor(); cur.ParaAdjust=3; ft.insertTextContent(cur,fld,False)
  except Exception as e: pass
first=T.createEnumeration().nextElement(); first.PageDescName='First Page'
idx.update()
doc.storeToURL(uno.systemPathToFileUrl(OUT),(pv('FilterName','MS Word 2007 XML'),))
doc.storeToURL(uno.systemPathToFileUrl('/tmp/claude-0/birlashgan.pdf'),(pv('FilterName','writer_pdf_Export'),))
doc.close(True); proc.terminate(); print('ok')
