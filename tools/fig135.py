import sys
sys.path.insert(0,'/home/user/tarjima/tools')
import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import wordart as W
BLUE,ORANGE,GREEN,DBLUE='2A78D6','EB6834','70AD47','1F4E79'
d=docx.Document()
for s in d.sections:
  s.left_margin=Cm(3); s.right_margin=Cm(1.5); s.top_margin=Cm(2); s.bottom_margin=Cm(2)
p=d.add_paragraph('{{FIG}}'); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
c=d.add_paragraph(); c.alignment=WD_ALIGN_PARAGRAPH.CENTER
r=c.add_run('1.3.5-rasm. Likvidlikni qoplash koeffitsiyentiga qo‘yiladigan talablar: islohotning boshlanishi va davomi, %')
r.bold=True; r.font.name='Times New Roman'; r.font.size=Pt(14)
Y=list(range(2015,2027))
bas=[60,70,80,90]+[100]*8
rus=[None,70,80,90,100,100,100,None,None,100,100,100]
uzb=[None,80,90]+[100]*9
xml,rows=W.scatter_chart([
  ('Bazel qo‘mitasi',Y,bas,BLUE,28575,None,'circle',('t',[0,1,2,3,4],DBLUE)),
  ('Rossiya',Y,rus,ORANGE,28575,None,'square',('b',[1,2,3,4,9])),
  ('O‘zbekiston',Y,uzb,GREEN,28575,None,'triangle',('l',[1,2,3]))],
  xtitle='Yillar',ytitle='Minimal talab, %',xmin=2014,xmax=2027,xmajor=1,ymin=50,ymax=110,ymajor=10,xfmt='0',legend=True)
rid=W.add_chart(d,(xml,rows))
W.replace_placeholder(d,'{{FIG}}',W.chart_run_xml(rid,16.0,9.0,'1.3.5-rasm'))
d.save('/home/user/tarjima/paragraflar/rasmlar/1.3.5-rasm.docx'); print('ok')
