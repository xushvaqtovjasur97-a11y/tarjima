# Karminskiy A.M. «Кредитные рейтинги и их моделирование» (VShE, 2015)

**Bibliografiya:** Карминский А.М. Кредитные рейтинги и их моделирование. — М.: Изд. дом НИУ ВШЭ, 2015. — 304 с. ISBN 978-5-7598-1232-6. (PDF sahifasi = bosma sahifa, siljish yo‘q; matn qatlami bor.)

## 1.2 Bazel kelishuvlari (pp.15–24, qisman)
- Risklarni boshqarishning 3 tamoyili: tarkibiy risklar; butun kompaniyaning yo‘qotish salohiyati; maksimal yo‘l qo‘yiladigan yo‘qotish salohiyati (p.15–16).
- Masofaviy usullar, «zaif signallar» bo‘yicha diagnostika, erta ogohlantirish tizimi — tashqi tahdidlar haqida o‘z vaqtida ma’lumot beruvchi axborot tizimi (p.18).
- CAMELS komponentlari: C, A, M, E, L, S (p.19). Reytingning sifati — kutilgan va haqiqiy natijalarning mos kelishi, «matematik ichki tuzilishi ahamiyatsiz» (p.20).
- Bazel I (1988, 1992-yilgacha joriy): Tier 1/Tier 2, kapitalning kamida yarmi Tier 1; 1996 — bozor riski (p.20–21).
- Bazel II (2004/2006): 3 ustun; kredit, bozor va operatsion risklar; IRB yondashuvi (PD, LGD, EAD, M); RWA ga nisbatan minimum ≥8% (p.21–22). Ichki reytinglarga qo‘yiladigan 8 ta talab (p.22). Moody’s baholashi: AQSh banklari uchun kapital talablari ~30% ga kamayishi mumkin (p.23).
- Bazel III: kapital talablarini qattiqlashtirish, kontsentratsiya, tizimli risk, protsiklik, joriy va o‘rta muddatli likvidlik ko‘rsatkichlari, resekyuritizatsiya (p.23–24).

## 2.2.3 Reytinglar konstruktori (pp.87–94)
- Dinamik reyting konstruktori (KDR): ichki reyting (hisobotdan olingan ko‘rsatkichlar) + tashqi reyting (sektordagi ulush) = dinamik reyting (p.88).
- Xususiyatlari: bir qator hisobot sanalari bo‘yicha dinamika; eski qiymatlarni asta-sekin «unutish»; tartib (ordinal) shkalalar; sektor bo‘yicha to‘liq ma’lumotlar bazasi (p.88).
- Formula (2.1): D(P^k) = 2/(T(T+1)) · Σ_{t=1..T} t·P_t^k / Σ_{j=1..N} P_t^j, T=12 (vazni chiziqli kamayib boradigan sirpanuvchi «oyna») (p.89).
- Ichki ko‘rsatkichlar ballari guruh ichidagi tanlanma taqsimot funksiyasi kvantillari orqali, [0,1] ehtimollik shkalasini teng bo‘laklarga bo‘lish yo‘li bilan beriladi (p.89–91).
- Ko‘lam guruhlari (masalan, passivlar summasi bo‘yicha): X — 70% hajm, L — 15%, M — 10%, S — 5% (p.90). Reyting shkalasi har bir guruh uchun alohida tuziladi.
- Ko‘rsatkichlarni CAMEL dan tanlash, takrorlanishga yo‘l qo‘ymaslik; «makrovaziyat qanchalik barqaror bo‘lsa, shunchalik kam omil ta’sir qiladi» (p.91).
- «Mobile» agentligining Moliyaviy barqarorlik dinamik reytingi (RDFS) shu asosda tuzilgan (p.90). RASK shkalasi: 10–12 gradatsiya, ~4 sinf (p.93).

## 2.3 Banklar reytinglarining ekonometrik modellari (pp.99–135)
- Jadval 2.10: S&P/Moody’s raqamli shkalasi (sinflar 1–8; gradatsiyalar 1–18; aralash). Kichik son = yuqori reyting, shuning uchun manfiy koeffitsiyent = ijobiy ta’sir (p.99).
- 1998-yildan keyingi Rossiya RA modellari: kapital yetarliligi, kapital hajmi, likvidlik ijobiy; uzoq muddatli kreditlar ulushi salbiy (U-shakl); muddati o‘tgan qarz va ROE ahamiyatsiz chiqqan. Ba’zi ko‘rsatkichlar U-shakl → kvadratik had (p.100). Barqarorlik uchun: hajm logarifmi, makroo‘zgaruvchilar, ordinal shkalalar (p.100).
- Moody’s modeli: 311 bank, 28 mamlakat, 777 kuzatuv (2003–2005); mB/mY/mP (kvantil shkala)/mN (normallashtirilgan) modellar; Pseudo R² 0,431–0,439; aniq bashorat 69–71%, ±1 sinf ~95% (Jadval 2.11, p.104). Makro bilan (macro2): Pseudo R² 0,571, aniq 77,9%, ±1 99,5% (Jadval 2.12, p.107); out-of-sample 1000 simulyatsiya (75/25): aniq 76,4%, >99% ±1 (Jadval 2.14, p.109).
- Xulosa: Moody’s nisbiy holatga (bank tizimidagi o‘rni) qaraydi, absolyut ko‘rsatkichlarga emas (p.108).
- Xalqaro taqqoslash: 551 bank, 86 mamlakat, 5629 kuzatuv (1995–2009), MSFO, Bloomberg, 30+ CAMELS o‘zgaruvchisi, 1 yillik lag (p.110–112). Jadval 2.15 (p.113–114): ln(aktivlar) −0,38…−0,56; Kapital/Aktivlar −1,6…−3,0; LLR/Aktivlar +12…+43; Foiz xarajatlari/daromadlari +0,17…+0,35; Taqsimlanmagan foyda/Aktivlar −1,4…−9,8; Yuqori likvid aktivlar/Majburiyatlar (∩); korrupsiya indeksi, inflyatsiya (−), Eksport/Import, PPP YaIM (+); Pseudo R² 0,19–0,30.
- Lag: 0,5–1,5 yil; tavsiya 1 yil, rivojlanayotgan mamlakatlar uchun 1–1,5 yil (p.117). Aniqlik: aniq 40–44%, ±1 90–91%, ±2 ~99% (Jadval 2.17, p.123). Modellar konservativligi past (manfiy asimmetriya; Moody’s −0,55) (p.121).
- OECD dummy ijobiy; MDH/Rossiya dummy ahamiyatsiz; S&P eng konservativ, Moody’s > Fitch (Jadval 2.19, p.125).
- RSBU bo‘yicha Rossiya banklari: 370 bank, 2600+ kuzatuv (2006–2010). O‘zgaruvchilar (Jadval 2.20, p.127): LTA, EQ_TA (∩, kvadrat), LIQ_1, LIQ_2, LC_TA, NPL_LC, LLP_LC, ROA, EFF, INC_ST, ST_DUM, FOR_DUM, SOV_DUM. Rossiya RA modellari (Jadval 2.21, p.131): Ekspert RA Pseudo R² 0,635, aniq 85%/±1 94%; Rus-Reyting 0,372 (69/94); NRA 0,277 (47/92); AK&M 0,425 (58/90). Xorijiy RA (Jadval 2.22): Pseudo R² 0,21–0,39, aniq 38–45%, ±1 89–92%.
- Davlat/xorijiy qo‘llab-quvvatlash dummylari ijobiy; faqat hajm + mulkchilik dummylari bo‘lgan modellar deyarli bir xil aniqlik beradi (p.130–134). Past reytinglar (Caa) yomon bashorat qilinadi (Jadval 2.23, p.135).

## 3.3 Bank defolti modellari. Rossiya tajribasi (pp.182–216)
- Rossiya bank tizimi bosqichlari: shakllanish (1989–1999), dinamik o‘sish (2000–2008; aktivlar/YaIM 67%, kapital 9%, kreditlar 48% — 2008), barqaror rivojlanish (2010–2014) (p.184–186).
- 1998-yil inqirozi modellari: LIVE=0, agar 2 yil ichida litsenziya chaqirilsa; 1996–2002 «Mobile» ma’lumotlari, 3000+ epizod; 12 ko‘rsatkich (Jadval 3.2: BP, KE, KDB, NDO, PNA, VB, SK, RES, LA, VFDL, PNAOS, GDO) (p.186–188). Adekvat lag 0,5–1,5 yil.
- Ekspert klasterlari: VB, GDO/VB, KE/VB, SK/VB bo‘yicha 12 klaster (8 tasi tahlil qilingan) (p.189–190). Optimal ln(VB) 11,6–12,4; optimal KE/VB 0,39–0,40; SK/VB deyarli barcha modellarda ahamiyatli va ijobiy (Jadval 3.4, p.193–195). 10 ta eng yomon bankdan ~57% to‘g‘ri aniqlanadi (p.196).
- Makro bilan: EXP/IMP va valyuta kursi (ERATE) eng yaxshi qo‘shimchalar; ln(VB/DEFL)² — kvadratik bog‘liqlik (Jadval 3.5, p.198).
- I/II tur xatolar formulalari: P1(c)=(M−m_c)/M; P2(c)=(n_c−m_c)/(N−M); egri chiziq Γ={(P1(c),P2(c))} (3.11–3.13, p.199–200). I tur xatoning narxi II turnikidan yuqori (p.193).
- 1998–2011 modeli: 910 ta litsenziya chaqirilishi + 37 sanatsiya, 374 holat modellashtirilgan; kvartal ma’lumotlari; logit (p.201–204). ANOVA bilan tanlangan o‘zgaruvchilar (Jadval 3.6, p.205–206): sk_ca, bp_ca, la_ca, ncb_ca, pzs_ke, oks_ca, ln_ca, res_ca va boshq.; CAMELS bo‘yicha guruhlash (p.206–207).
- Nomutanosiblik: defoltlar ~1% → 1000 kichik tanlanma (5% «tirik» + barcha defoltlar → 10–12% defolt) (p.208). Lag 2 kvartal; lag1 AUC 0,851, lag2 0,828, lag8 0,722 (Rasm 3.9, p.214).
- Model ketma-ketligi basic → b_e1 → b_e2 → b_e2_t → b_e2_t_m → b_e2_t_m_i → yakuniy (Jadval 3.7, p.210–211): Pseudo R² ~0,52 dan ~0,71 gacha, ROC AUC ~0,89 dan ~0,97 gacha, to‘g‘ri bashorat ~93–95%.
- Natijalar: optimal sk_ca ≈ 0,4; optimal bp_ca ≈ 0,05; davlatga tegishli bo‘lmagan qimmatli qog‘ozlar <10% bo‘lsa zararsiz; hajm faqat kvadratik shaklda ahamiyatli («too big to fail» tasdiqlanmagan); 2009-yil dummysi (+), I kvartal (−); CPI (+), YaIM o‘sishi ahamiyatsiz; Lerner indeksi (−), Moskva dummysi (p.213–216).
- Out-of-sample 2010–2011 (19 defolt): PD ≥30% chegarasi, risk guruhi 30 bank, 63% defolt to‘g‘ri bashorat qilingan (Jadval 3.8: X=10 → 84%, 40 → 52%) (p.216–217).

## 3.6 Kredit risklarini stress-test qilish (pp.230–241)
- Ta’rif: «istisno, ammo ehtimolli» makroshoklarga barqarorlikni baholash (Blaschke 2001); IMF 2012 tizimli ahamiyatli banklarni (D-SIB) tahlil qilishga urg‘u beradi (p.231). Čihák (2007) NPL–CAR koordinatalari (p.231).
- Chebyshev: P{ROA < −EQ/TA} ≤ Var(ROA)/(E(ROA)+EQ/TA)² (3.24); Z = (E(ROA)+EQ/TA)/√Var(ROA) (3.25); yangi o‘lchov Z(PZS) = (LLR − E(PZS))/√Var(PZS) (3.26); 2 yillik oyna; EQ/TA o‘rniga N1 ning ortiqcha qismi (p.232–234).
- Mulkchilik turi bo‘yicha 3 guruhdan 2 tadan bank (davlat, xorijiy, xususiy); segmentlar (korporativ/chakana); chegara = mediana; Z(PZS)×Z(ROA) kvadrantlari: kuchli/riskli banklar (Jadval 3.10–3.11, p.234–236).
- Sezgirlik: kreditlarning 85% ini bergan banklar (IMF tavsiyasi) — 104 korporativ, 97 chakana bank; oylik ma’lumotlar 2004–2012; VAR(3): X_t = A0 + A1X_{t−1}+A2X_{t−2}+A3X_{t−3}+e_t; X=(M0, INV, Y, K, PZS) (3.27) (p.236–238).
- Variatsiya dekompozitsiyasi (1 yil): korporativ PZS o‘sishining 23% i makro bilan (YaIM 10%, M0 9%, INV 4%), 14% kredit o‘sishi bilan; chakana PZS ~40% makro bilan, 8% kredit bilan (p.238–239).
- YaIM shoki (1 s.o. = oylik +0,7 p.p.) → PZS o‘sishi korporativda −2,4 p.p., chakanada −2,7 p.p.; ta’sir 1 kvartalda so‘nadi (G‘arbda 3–4 kvartal) (p.239).
- Stress-ssenariy = 2008 takrorlanishi (korporativ PZS shoki +50,5 p.p., chakana +31,4 p.p.); chakanada tiklanish ~2 yil (p.240). Xulosa: chakana segment makroshoklarga sezgirroq (p.241).

## 4.2 Yagona reyting maydoni (pp.257–270)
- YeRP: bazaviy shkalani tanlash + barcha RA va ichki reytinglarni har bir sinf subyektlari uchun bazaviy shkalaga akslantirish (tarixiy mapping) (p.258).
- Parametrik mapping algoritmi (6 qadam): bazaviy shkala → har bir RS_i ni raqamli NS_i ga → F_i(α_i) → yaqinlik o‘lchovi → integral mezon → optimallashtirish (p.259–260). 9 ta metodologik vazifa va 7 ta mapping tamoyili (p.261–262).
- Formulalar: F_i: R_i → B (4.1); S = Σ[F_i1(R_i1jt) − F_i2(R_i2jt)]² → min (4.3–4.4); F_i = α_i1·f_i(R_i)+α_i2 (4.5) (p.263–265).
- Ma’lumotlar: 2646 kuzatuv, 370 bank, 2006–2010, 10 ta shkala (Big-3 ning 6 tasi + 4 ta Rossiya RA); bazaviy shkala — Moody’s xalqaro (3400+ juftlik); Aaa=1 … C=20 (p.265–266).
- PD = 0,000218·R^3,8 (Moody’s 1980–2008) (4.6); ln(M)=a_i·ln(R_i)+b_i (4.7) (p.266). Moody’s–S&P: y=0,8408x+1,1287, R²=0,815 (Rasm 4.3); Moody’s xalq.–milliy: y=9,80x^0,226, R²=0,83 (Rasm 4.4) (p.269–270).
- Jadval 4.2 (p.271): ln(M)=Σ(a_i ln R_i + b_i)d_i; masalan, S&P a=0,916, b=0,146; Ekspert RA a=0,373, b=2,329; N=3432, Pseudo R²=0,902. Aniqlik ¼ gradatsiya; yillik sirpanuvchi yangilash (p.272).

**Dissertatsiyada qo‘llash**
- §1.3/2.1 – Masofaviy reyting modellarining sharhi: KDR (2.1 formula), Moody’s/S&P/Fitch modellari omillari (Jadval 2.15, 2.20) — CAMELS asosidagi ekonometrik yondashuvning empirik isboti sifatida.
- §3.1 – MBII ning 1–4 ballik shkalasi va guruhlash uchun asos: kvantil ballash (p.89–91), X/L/M/S ko‘lam guruhlari (70/15/10/5%); U-shaklli ko‘rsatkichlar (kapital ~0,4, ROA ~0,05, kreditlar/aktivlar ~0,39–0,40) uchun «optimal oraliq» — chegaralarni ikki tomonlama qo‘yishni asoslash.
- §4.1/3.2 – D-SIB monitoringi uchun Z(ROA) va Z(PZS) (3.25–3.26) va mediana-kvadrant usuli; SQB/Mikrokreditbank ni boshqa banklar bilan Z-kvadrantda ko‘rsatish; erta ogohlantirish signallari uchun PD ≥30% risk guruhi mantig‘i.
- §4.1 – MBII bashorat kuchini tekshirish: I/II tur xatolar (3.12–3.13), ROC/AUC, out-of-sample (75/25, 1000 simulyatsiya), lag 2 kvartal–1 yil; MBII darajalarini (75/50/25) xalqaro reyting shkalasiga ln-ln mapping orqali bog‘lash (4.7, Jadval 4.2).
- §3.3/4.2 – NPL satellit modeli uchun VAR spetsifikatsiyasi (3.27), YaIM, M0, investitsiyalar, kredit o‘sishi; stress-ssenariyni tarixiy eng yomon shok sifatida kalibrlash (p.240); chakana segmentning yuqori sezgirligi.
- §3.2 – davlat qo‘llab-quvvatlashi (ST_DUM) reytingni oshiradi: davlat banklarining reytingi «qo‘llab-quvvatlash» hisobiga yuqori, bu MBII ning sof moliyaviy bahosidan farq qilishini izohlaydi.

**Iqtiboslar**
- «Качество рейтинга определяется сходимостью ожидаемых и фактических результатов. Не имеет значения математическая начинка системы рейтинга.» (p.20)
- «Чем стабильнее макроэкономическая ситуация, тем меньшее число факторов оказывает влияние на состояние субъектов.» (p.91)
- «…рейтинговое агентство Moody’s ориентируется в большей степени на относительное положение банка в банковской системе, а не на абсолютные показатели его деятельности.» (p.108)
- «Очевидно, цена ошибки I рода выше цены ошибки II рода.» (p.193)
- «К группе риска следует причислять те банки, вероятность дефолта которых превышает 30%.» (p.216)
- «Стресс-тестирование определяется как оценка устойчивости банковской системы к исключительным, но вероятным макроэкономическим шокам» (p.231)
- «…как рост объемов, так и рост просроченной задолженности более чувствителен к макроэкономическим факторам в розничном сегменте.» (p.239)

**Moslashtiriladigan jadval/formulalar**
- (2.1) dinamik ulush reytingi, X/L/M/S guruhlash (p.89–90); Jadval 2.10 raqamli shkala (p.99); Jadval 2.15 va 2.20 omillar va kutilgan belgilari (p.113, 127); Jadval 2.17 xato taqsimoti (p.123).
- Jadval 3.6 ANOVA bo‘yicha CAMELS o‘zgaruvchilari (p.205); Jadval 3.7 model ketma-ketligi (p.210–211); Jadval 3.8 risk guruhi chegarasi (p.217); (3.11–3.13) P1/P2 (p.199).
- (3.24)–(3.26) Z-score va Z(PZS) (p.232–233); Jadval 3.11 (p.235); (3.27) VAR (p.238).
- (4.3)–(4.8) mapping, Jadval 4.2 (p.264–271).
