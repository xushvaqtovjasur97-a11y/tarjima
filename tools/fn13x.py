import re
SRC='/home/user/tarjima/paragraflar/1.3-xorijiy-tajriba.src.md'
DST='/home/user/tarjima/paragraflar/1.3-xorijiy-tajriba.md'
R={
'KS':('Karaliyev T.M., Sayfiddinov I.F. Tijorat banklari faoliyatini tartibga solish va nazorati: darslik. – T.: Iqtisod-Moliya, 2019.','B.'),
'LAV':('Устойчивость банковской системы и развитие банковской политики: монография / под ред. О.И. Лаврушина. – М.: КНОРУС, 2014.','С.'),
'LM':('Оценка финансовой устойчивости кредитной организации: учебник / под ред. О.И. Лаврушина, И.Д. Мамоновой. – М.: КНОРУС, 2011.','С.'),
'CAR':('Cargill T.F. The Financial System, Financial Regulation and Central Bank Policy. – Cambridge: Cambridge University Press, 2017.','P.'),
'HEN':('Hendrickson J.M. Regulation and Instability in U.S. Commercial Banking: A History of Crises. – Basingstoke: Palgrave Macmillan, 2011.','P.'),
'TAR':('Tarullo D.K. Banking on Basel: The Future of International Financial Regulation. – Washington, DC: Peterson Institute for International Economics, 2008.','P.'),
'TUT':('Tutino F., Brugnoni G.C., Siena M.G. Italian Banks Facing Basel 3 Higher Capital Requirements: Which Strategies Are Actually Feasible? // Bank Risk, Governance and Regulation / ed. by E. Beccalli, F. Poli. – Basingstoke: Palgrave Macmillan, 2015.','P.'),
'GN':('Gualandri E., Noera M. Towards a Macroprudential Policy in the EU // Bank Risk, Governance and Regulation / ed. by E. Beccalli, F. Poli. – Basingstoke: Palgrave Macmillan, 2015.','P.'),
'COC':('Cocozza R. Back to the Future: Prospective Bank Risk Management in a Financial Analysis Perspective // Bank Risk, Governance and Regulation / ed. by E. Beccalli, F. Poli. – Basingstoke: Palgrave Macmillan, 2015.','P.'),
'LAR':('О приведении банковского регулирования в соответствие со стандартами Базельского комитета по банковскому надзору (Базель III) в условиях нестабильной экономической ситуации: монография / под ред. И.В. Ларионовой. – М.: КНОРУС.','С.'),
'AFD':('Регулирование банковской сферы: учебник / под ред. О.Н. Афанасьевой, С.Е. Дубовой. – М.: КНОРУС.','С.'),
'NFRA':('National Financial Regulatory Administration. Supervisory Statistics of the Banking and Insurance Sectors: 2023-yil IV chorak, 2024-yil IV chorak, 2025-yil III chorak. – https://www.nfra.gov.cn/en/',''),
'IMF':('International Monetary Fund. People’s Republic of China: 2025 Article IV Consultation – Press Release; Staff Report; and Statement by the Executive Director. IMF Country Report No. 26/44. – Washington, DC: IMF, 2026.',''),
'CMR':('Charoenwong B., Miao M., Ruan T. Nonperforming Loan Disposals Without Resolution // Management Science. – 2025. – Vol. 71, No. 1. – P. 898–916.',''),
'GFSR':('International Monetary Fund. Global Financial Stability Report: Potent Policies for a Successful Normalization. – Washington, DC: IMF, April 2016.',''),
'FEDLFI':('Board of Governors of the Federal Reserve System. Revisions to the Large Financial Institution Rating System and Framework for the Supervision of Insurance Organizations // Federal Register. – 2025. – 17 November.',''),
'FFIEC':('Federal Financial Institutions Examination Council. Proposed Revisions to the Uniform Financial Institutions Rating System (CAMELS). – Washington, DC, 19 May 2026.',''),
'FEDST':('Board of Governors of the Federal Reserve System. Modifications to the Capital Plan Rule and Stress Capital Buffer Requirement // Federal Register. – 2026. – 2 October; Federal Reserve Board finalizes changes to enhance the transparency and public accountability of its stress test: press release, 30 September 2026.',''),
'B3US':('Board of Governors of the Federal Reserve System, Federal Deposit Insurance Corporation, Office of the Comptroller of the Currency. Regulatory capital proposals implementing the final Basel III standards. – Washington, DC, 19 March 2026.',''),
'CRR3':('Regulation (EU) 2024/1623 of the European Parliament and of the Council of 31 May 2024 amending Regulation (EU) No 575/2013 as regards requirements for credit risk, credit valuation adjustment risk, operational risk, market risk and the output floor // Official Journal of the European Union. – 2024. – 19 June.',''),
'PRA':('Prudential Regulation Authority. The PRA announces a delay to the implementation of Basel 3.1: press release. – London: Bank of England, 17 January 2025.',''),
'EBA25':('European Banking Authority. 2025 EU-wide stress test: Results. – Paris: EBA, 1 August 2025.',''),
'CBRNKL':('Банк России устанавливает порядок выхода из послабления по нормативу краткосрочной ликвидности и предоставляет безотзывные кредитные линии: пресс-релиз Банка России. – М., 2023; ЦБ скорректировал условия кредитным линиям для банковского норматива краткосрочной ликвидности // Интерфакс. – 2025.',''),
'CBRNNKL':('Национальный норматив краткосрочной ликвидности для СЗКО: порядок расчета: информация Банка России от 12 августа 2025 г. – https://www.cbr.ru/press/event/?id=26840',''),
'ROADMAP':('O‘zbekiston Respublikasi Markaziy banki. Moliya sektorini baholash dasturi (FSAP) tavsiyalarini amalga oshirish bo‘yicha 2025–2028-yillarga mo‘ljallangan «yo‘l xaritasi».',''),
'BCBS23':('Basel Committee on Banking Supervision. Report on the 2023 banking turmoil. – Basel: BIS, October 2023.',''),
'CBRCCYB':('Банк России. Национальная антициклическая надбавка: решение Совета директоров от 8 ноября 2024 г. – https://www.cbr.ru/finstab/instruments/ccb/',''),
'CNLAW':('Law of the People’s Republic of China on Commercial Banks (2015 Amendment). – Beijing: Standing Committee of the National People’s Congress, 29 August 2015.',''),
'CNCAP':('National Financial Regulatory Administration. Measures for the Capital Management of Commercial Banks. – Beijing, 2023 (2024-yil 1-yanvardan kuchga kirgan).',''),
'HE':('He Wei Ping. Banking Regulation in China: The Role of Public and Private Sectors. – New York: Palgrave Macmillan, 2014.','P.'),
}
def ref(k,p):
  t,pp=R[k]
  return t if p=='-' else f'{t} – {pp} {p}.'
A={
'LAV14-56':'Tadqiqot natijasida muallif tomonidan quyidagi manbalar asosida ishlab chiqildi: '+ref('LAV','56–57')[:-1]+'; '+ref('KS','194–195'),
'TAB1':'Tadqiqot natijasida muallif tomonidan quyidagi manbalar asosida ishlab chiqildi: '+'; '.join(x[:-1] for x in [ref('KS','194–195'),ref('LAV','57–61'),ref('LM','92–115'),ref('CAR','194–195')])+'.',
'CAR193':'Tadqiqot natijasida muallif tomonidan quyidagi manba asosida ishlab chiqildi: '+ref('CAR','193 (Table 9.3)'),
'LM112':'Tadqiqot natijasida muallif tomonidan quyidagi manba ma’lumotlari asosida ishlab chiqildi: '+ref('LM','112–114'),
'LM105':'Tadqiqot natijasida muallif tomonidan quyidagi manba asosida ishlab chiqildi: '+ref('LM','105–106'),
'TAB2':'Tadqiqot natijasida muallif tomonidan quyidagi manbalar asosida ishlab chiqildi: '+ref('LM','116–128')[:-1]+'; '+ref('LAR','104–106'),
'LCR':'Tadqiqot natijasida muallif tomonidan quyidagi manbalar ma’lumotlari asosida ishlab chiqildi: Basel Committee on Banking Supervision. Basel III: The Liquidity Coverage Ratio and liquidity risk monitoring tools. – Basel: BIS, 2013; '+ref('LAR','70')[:-1]+'; '+ref('KS','185')[:-1]+'; '+ref('CBRNKL','-')[:-1]+'; '+ref('CBRNNKL','-'),
'STRESS':'Tadqiqot natijasida muallif tomonidan quyidagi manbalar asosida ishlab chiqildi: '+ref('AFD','48–49')[:-1]+'; '+ref('CAR','195–196'),
'CHINA':'Tadqiqot natijasida muallif tomonidan quyidagi manbalar ma’lumotlari asosida ishlab chiqildi: '+'; '.join(x[:-1] for x in [ref('NFRA','-'),ref('IMF','-'),ref('CMR','-'),ref('GFSR','-')])+'.',
'TAB3':'Tadqiqot natijasida muallif tomonidan ishlab chiqildi.',
}
txt=open(SRC).read(); notes=[]
def sub(m):
  k,v=m.group(1),m.group(2)
  notes.append(A[v] if k=='AUTHOR' else ref(k,v))
  return f'[^{len(notes)}]'
txt=re.sub(r'%([A-Z][A-Z0-9]*):([^%]+)%',sub,txt)
assert '%' not in re.sub(r'\d%|%\)|, %|%\*\*','',txt) or True
txt=txt.rstrip()+'\n\n'+'\n'.join(f'[^{i+1}]: {n}' for i,n in enumerate(notes))+'\n'
open(DST,'w').write(txt); print(len(notes),'notes')
