import re
SRC='/home/user/tarjima/paragraflar/1.2-paragraf.src.md'
DST='/home/user/tarjima/paragraflar/1.2-paragraf.md'
R={}
for f in ['fn11.py','fn13x.py','fn22.py']:
    code=open('/home/user/tarjima/tools/'+f).read(); code=code[:code.find('def ref')]
    ns={}; exec(code,ns); R.update(ns['R'])
R.update({
'KARM':('Карминский А.М. Кредитные рейтинги и их моделирование. – М.: Изд. дом НИУ ВШЭ, 2015.','С.'),
'MOD':('Modina M. Credit Rating and Bank–Firm Relationships: New Models to Better Evaluate SMEs. – Basingstoke: Palgrave Macmillan, 2015.','P.'),
'BERD':('Бердияров Б.Т. Ўзбекистон Республикаси тижорат банкларининг ликвидлилиги ва тўлов қобилиятини таъминлаш масалалари: монография. – Т.: BMA, 2020.','Б.'),
'SHAT':('Шаталова Е.П. Банковские рейтинги в системе риск-менеджмента: процедуры мониторинга кредитных рейтингов: учебное пособие. – 2020.','С.'),
'IZZI':('Izzi L., Oricchio G., Vitale L. Basel III Credit Rating Systems: An Applied Guide to Quantitative and Qualitative Models. – Basingstoke: Palgrave Macmillan, 2012.','P.'),
'SOK':('Соколинская Н.Э. Анализ и оценка финансового состояния коммерческого банка на основе международных стандартов финансовой отчетности: учебное пособие. – М.: РУСАЙНС.','С.'),
'LANG':('Langohr H., Langohr P. The Rating Agencies and Their Credit Ratings: What They Are, How They Work and Why They Are Relevant. – Chichester: John Wiley & Sons, 2008.','P.'),
'IMFFSI':('International Monetary Fund. 2019 Financial Soundness Indicators Compilation Guide. – Washington, DC: IMF, 2019.',''),
'B3LCR':('Basel Committee on Banking Supervision. Basel III: The Liquidity Coverage Ratio and liquidity risk monitoring tools. – Basel: BIS, January 2013.',''),
'B3NSFR':('Basel Committee on Banking Supervision. Basel III: The Net Stable Funding Ratio. – Basel: BIS, October 2014.',''),
'B3FIN':('Basel Committee on Banking Supervision. Basel III: Finalising post-crisis reforms. – Basel: BIS, December 2017.',''),
'KS':('Karaliyev T.M., Sayfiddinov I.F. Tijorat banklari faoliyatini tartibga solish va nazorati: darslik. – T.: Iqtisod-Moliya, 2019.','B.'),
})
def ref(k,p):
  t,pp=R[k]
  return t if p=='-' else f'{t} – {pp} {p}.'
def J(*xs): return '; '.join(x.rstrip('.') for x in xs)+'.'
SRCS='Tadqiqot natijasida muallif tomonidan quyidagi manbalar asosida ishlab chiqildi: '
A={
'T1':SRCS+J(ref('AFD','49–53'),ref('TAR','16–23'),ref('PEL','4, 17, 61–62'),ref('CAR','184–185')),
'T2':SRCS+J(ref('LM','54–57'),ref('LAV','48–61'),ref('KARM','20, 208'),ref('KS','195–213'),ref('CP','77–93'),ref('CAR','194')),
'F1':SRCS+J(ref('TAR','45–67'),ref('TIAN','3–28'),ref('KARM','20–24')),
'F2':SRCS+J(ref('MOD','46–47'),ref('TIAN','6–7, 15')),
'F3':SRCS+J(ref('TIAN','20–22')),
'T3':SRCS+J(ref('B3LCR','-'),ref('B3NSFR','-'),ref('LAR','69, 93–94')),
'F4':SRCS+J(ref('IFRS9','-'),ref('SOK','94–95, 104')),
'T4':'Tadqiqot natijasida muallif tomonidan ishlab chiqildi.',
'F5':'Tadqiqot natijasida muallif tomonidan ishlab chiqildi.',
}
txt=open(SRC).read(); notes=[]
def sub(m):
  k,v=m.group(1),m.group(2)
  notes.append(A[v] if k=='AUTHOR' else ref(k,v))
  return f'[^{len(notes)}]'
txt=re.sub(r'%([A-Z][A-Z0-9]*):([^%]+)%',sub,txt)
txt=txt.rstrip()+'\n\n'+'\n'.join(f'[^{i+1}]: {n}' for i,n in enumerate(notes))+'\n'
open(DST,'w').write(txt); print(len(notes),'notes')
