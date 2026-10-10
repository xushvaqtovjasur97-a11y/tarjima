import re
SRC='/home/user/tarjima/paragraflar/1.1-paragraf.src.md'
DST='/home/user/tarjima/paragraflar/1.1-paragraf.md'
LM='Оценка финансовой устойчивости кредитной организации: учебник / под ред. О.И. Лаврушина, И.Д. Мамоновой. – М.: КНОРУС, 2011.'
LAV='Устойчивость банковской системы и развитие банковской политики: монография / под ред. О.И. Лаврушина. – М.: КНОРУС, 2014.'
BR='Bank Risk, Governance and Regulation / ed. by E. Beccalli, F. Poli. – Basingstoke: Palgrave Macmillan, 2015.'
R={
'CAR':('Cargill T.F. The Financial System, Financial Regulation and Central Bank Policy. – Cambridge: Cambridge University Press, 2017.','P.'),
'LM':(LM,'С.'),'LAV':(LAV,'С.'),
'SCH':('Schinasi G.J. Defining Financial Stability: IMF Working Paper WP/04/187. – Washington, DC: International Monetary Fund, 2004. – 18 p.',''),
'MISH':('Mishkin F.S. Global Financial Instability: Framework, Events, Issues // Journal of Economic Perspectives. – 1999. – Vol. 13, No. 4. – P. 3–20.',''),
'HEN':('Hendrickson J.M. Regulation and Instability in U.S. Commercial Banking: A History of Crises. – Basingstoke: Palgrave Macmillan, 2011.','P.'),
'SHAR':('Sharipova N.H. Tijorat banklari moliyaviy barqarorligini mustahkamlashning konseptual asoslarini takomillashtirish: iqtisodiyot fanlari doktori (DSc) dissertatsiyasi avtoreferati. – Toshkent, 2026.','B.'),
'AFD':('Регулирование банковской сферы: учебник / под ред. О.Н. Афанасьевой, С.Е. Дубовой. – М.: КНОРУС.','С.'),
'PEL':('Peláez C.M., Peláez C.A. Regulation of Banks and Finance: Theory and Policy after the Credit Crisis. – Basingstoke: Palgrave Macmillan, 2009.','P.'),
'TAR':('Tarullo D.K. Banking on Basel: The Future of International Financial Regulation. – Washington, DC: Peterson Institute for International Economics, 2008.','P.'),
'DD':('Diamond D.W., Dybvig P.H. Bank Runs, Deposit Insurance, and Liquidity // Journal of Political Economy. – 1983. – Vol. 91, No. 3. – P. 401–419.',''),
'MIN':('Minsky H.P. The Financial Instability Hypothesis: Working Paper No. 74. – Annandale-on-Hudson: The Jerome Levy Economics Institute of Bard College, 1992. – 10 p.',''),
'THAK':('Thakor A.V. Leverage, System Risk and Financial System Health: How Do We Develop a Healthy Financial System? // Governance, Regulation and Bank Stability / ed. by T. Lindblom, S. Sjögren, M. Willesson. – Basingstoke: Palgrave Macmillan, 2014.','P.'),
'MERT':('Merton R.C. On the Pricing of Corporate Debt: The Risk Structure of Interest Rates // Journal of Finance. – 1974. – Vol. 29, No. 2. – P. 449–470.',''),
'ROY':('Roy A.D. Safety First and the Holding of Assets // Econometrica. – 1952. – Vol. 20, No. 3. – P. 431–449.',''),
'BG':('Boyd J.H., Graham S.L. Risk, Regulation, and Bank Holding Company Expansion into Nonbanking // Federal Reserve Bank of Minneapolis Quarterly Review. – 1986. – Vol. 10, No. 2. – P. 2–17.',''),
'CP':('Chiaramonte L., Poli F. Predicting European Bank Distress: Evidence from the Recent Financial Crisis // Governance, Regulation and Bank Stability / ed. by T. Lindblom, S. Sjögren, M. Willesson. – Basingstoke: Palgrave Macmillan, 2014.','P.'),
'BOR':('Borio C. Towards a Macroprudential Framework for Financial Supervision and Regulation?: BIS Working Papers No. 128. – Basel: Bank for International Settlements, 2003. – 32 p.',''),
'AB':('Adrian T., Brunnermeier M.K. CoVaR // American Economic Review. – 2016. – Vol. 106, No. 7. – P. 1705–1741.',''),
'CROCK':('Crockett A. Why Is Financial Stability a Goal of Public Policy? // Maintaining Financial Stability in a Global Economy. – Kansas City: Federal Reserve Bank of Kansas City, 1997. – P. 7–36.',''),
'IFRS9':('IFRS 9 Financial Instruments. – London: International Accounting Standards Board, 2014.',''),
}
def ref(k,p):
  t,pp=R[k]
  return t if p=='-' else f'{t} – {pp} {p}.'
def J(*xs): return '; '.join(x[:-1] for x in xs)+'.'
DEV='Tadqiqot natijasida muallif tomonidan ishlab chiqildi.'
SRCS='Tadqiqot natijasida muallif tomonidan quyidagi manbalar asosida ishlab chiqildi: '
A={
'FIG1':SRCS+J(ref('LM','14–15, 57–61'),ref('LAV','37–40')),
'TAB1':SRCS+J(ref('LM','24'),ref('LAV','38'),ref('SCH','-'),ref('MISH','-'),ref('HEN','6–7'),ref('THAK','13–16'),ref('SHAR','13–14')),
'FIG2':DEV,
'FIG3':SRCS+J(ref('LM','41–45'),ref('AFD','49–55')),
'FIG4':SRCS+J(ref('LM','31–41')),
'FIG5':'Tadqiqot natijasida muallif tomonidan quyidagi manba ma’lumotlari asosida ishlab chiqildi: '+ref('LM','44'),
'FIG6':'Tadqiqot natijasida muallif tomonidan quyidagi manba ma’lumotlari asosida ishlab chiqildi: '+ref('CP','92–95'),
'TAB2':SRCS+J(ref('DD','-'),ref('MIN','-'),ref('TAR','16'),ref('MERT','-'),ref('BG','-'),ref('BOR','-'),ref('AB','-'),ref('IFRS9','-')),
'FIG7':DEV,
}
txt=open(SRC).read(); notes=[]
def sub(m):
  k,v=m.group(1),m.group(2)
  notes.append(A[v] if k=='AUTHOR' else ref(k,v))
  return f'[^{len(notes)}]'
txt=re.sub(r'%([A-Z][A-Z0-9]*):([^%]+)%',sub,txt)
txt=txt.rstrip()+'\n\n'+'\n'.join(f'[^{i+1}]: {n}' for i,n in enumerate(notes))+'\n'
open(DST,'w').write(txt); print(len(notes),'notes')
