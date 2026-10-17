# Lindblom T., Sjögren S., Willesson M. (eds.) «Governance, Regulation and Bank Stability» (Palgrave, 2014)

**1. Bibliografiya:** Lindblom T., Sjögren S., Willesson M. (eds.). Governance, Regulation and Bank Stability. – Basingstoke: Palgrave Macmillan, 2014. Boblar:
- Thakor A.V. Leverage, System Risk and Financial System Health: How Do We Develop a Healthy Financial System? – pp. 9–19.
- Chiaramonte L., Poli F. Predicting European Bank Distress: Evidence from the Recent Financial Crisis. – pp. 77–99.
- Deng B., Casu B., Ferrari A. The Impact of Deregulation and Re-regulation on Bank Efficiency: Evidence from Asia. – pp. 100–128.

**2. Mazmun**

*5-bob. Chiaramonte va Poli (Z-score)*
- Distress indikatorlarining ikki guruhi (p.77–80):
  - bozorga asoslangan (DD, obligatsiya spredlari, CDS) – faqat listingdagi banklar uchun;
  - hisobga asoslangan (CAMEL, Z-score).
- Tanlanma (p.81–82): 12 mamlakatdagi 3 125 ta bank, 2001–2011. Davr inqirozdan oldingi (2001–2007) va inqiroz (2008–2011) davrlariga bo‘lingan.
- Distress ta’rifi (p.82): quyidagilardan biri.
  - bankrotlik, tugatish yoki faoliyatni to‘xtatish;
  - M&A, agar undan oldingi yili CR < 0 bo‘lsa, bu yerda CR = (kapital + kredit rezervi − NPL)/TA;
  - davlat yordami.
  - Natijada 149 bankda 185 ta hodisa aniqlangan.
- Model: probit, Pr(DB=1 | X_{t−1}) (5.1), barcha o‘zgaruvchilar bir yil lag bilan (p.83). Z-score = (ROAA + E/TA)/σ(ROAA), σ 3 yillik sirpanuvchi oynada (5.2), modelda ln_Z ishlatiladi (p.83–84).
- Nazorat o‘zgaruvchilari (Table 5.1, p.85): SIZE, GDPC, INFC, CR3, CRED = muammoli kreditlar/yalpi kreditlar, CIR, LIQ = sof kreditlar/(omonatlar + qisqa muddatli fondlar).
- Xatolar (p.84–85): I tur – distressni o‘tkazib yuborish; II tur – soxta signal. Kesish nuqtalari 1% va 10%, asosiy e’tibor 1% ga.
- Tavsifiy statistika (Table 5.2, p.86): ln_Z sog‘lom banklarda 4.551, distressdagilarda 3.222.
- Probit natijalari (Table 5.4, p.90–91):
  - ln_Z koeffitsiyenti barcha spetsifikatsiyalarda −0.14…−0.18 va ahamiyatli;
  - SIZE +0.17…+0.21, CIR +0.52, LIQ +0.19 (inqirozda +0.31);
  - makro o‘zgaruvchilar ahamiyatsiz;
  - Pseudo R² 0.232 dan 0.324 gacha oshadi.
- Prognoz aniqligi (p.92; Table 5.5, p.94–95):
  - faqat ln_Z: 185 tadan 139 ta distress to‘g‘ri topilgan (75%), sog‘lom banklar 81% to‘g‘ri;
  - to‘liq model: 81% va 87%;
  - inqiroz davrida to‘liq model: 91% va 83%;
  - 10% kesishda I tur xatosi 0.58–0.85 ga yetadi.
- Xulosa (p.93): Z-score "reliable and parsimonious". Bozor ma’lumotlari bo‘lmaganda PCA uchun yaroqli.

*2-bob. Thakor*
- Inqirozning sabablari (p.10–12): rag‘batlar buzilishi va e’tiqodlar buzilishi. Uzoq muvaffaqiyat bankirlar mahoratiga ishonchni oshiradi va riskni oshiradi. Shok kelganda funding to‘xtaydi: bu tashqaridan likvidlik inqiroziga o‘xshaydi, aslida to‘lovga qobiliyatsizlik inqirozi.
- Yuqori bank va iste’molchi leverage korrelyatsiyasi tizimli riskni oshiradi (p.12). AQSh banklarida kapital/aktiv 1893–2010 yillarda pasaygan (Fig. 2.1, p.13).
- Sog‘lom tizimning to‘rt belgisi (p.13–15, 16):
  1. vaqti-vaqti bilan bailoutsiz idiosinkratik bankrotliklar;
  2. yuqori kapital, soya bank tizimida ham;
  3. faol aksiyador boshqaruvi;
  4. regulyatsiyani shartnoma yorlig‘iga emas, iqtisodiy funksiyaga qarab yuritish.
- Yetti chora (p.15–16): bank bankrotlik kodeksi; RWAdan tashqari yuqori leverage ratio, balansdan tashqari moddalar bilan; dividend cheklovi orqali kontrsiklik kapital; soya bankka talablar; mulkdorlikni kengaytirish; qarz va kapital soliq rejimini tenglashtirish (Belgiya ACE).
- "Kapital kreditni kamaytiradi" degan da’vo xato: kapital talabi zaxira talabi emas (p.14).

*6-bob (qisman). Deng, Casu, Ferrari*
- Usul (p.113–115): stoxastik xarajat chegarasi (6.1), samarasizlik Battese–Coelli bo‘yicha (6.2), metafrontier formulasi Meta-E = CF-E × TGR (6.3), bootstrap.
- Ma’lumotlar (p.115–118): 8 ta Osiyo iqtisodiyoti, 2001–2010, 3 805 kuzatuv. Xarajat funksiyasi translog.
  - Kirishlar: PL, PC, PF. Chiqishlar: LN, OEA, NFC. Risk: CR, VOroa, LLP.
  - Regulyatsiya indekslari (Table 6.2, p.116–117): ACTR (1–4), CMD (0–10), CAPS (0–7), SUPP (0–15), MARD (0–7). Dereg = (ACTR+CMD)/2, Rereg = (CAPS+SUPP+MARD)/3.
- Natijalar:
  - texnologik bir xillik rad etilgan (p.118);
  - CAPS samaradorlikni pasaytiradi (p.121);
  - Dereg −0.557***, ya’ni +0.1 birlik xarajatni −5.57% ga kamaytiradi; Rereg ahamiyatsiz (Table 6.4, p.120–122);
  - kapital xarajatni oshirmaydi (signal effekti), LLP xarajatni oshiradi (p.122);
  - Meta-E o‘sgan, asosan TGR hisobiga; Malayziya istisno; Yaponiya eng samarali (p.122–124).
- Xulosa (p.124–125): barqarorlik va vositachilikni rag‘batlantiruvchi siyosatlarni uyg‘unlashtirish kerak.

**3. Dissertatsiyada qo‘llash**
- §4.1 – MBIIning bashorat kuchini Chiaramonte–Poli dizayni bilan tekshirish:
  - distress ta’rifi: litsenziyani chaqirib olish, sanatsiya, davlat kapitali yoki CR < 0;
  - probit modeli, MBII_{t−1} bilan;
  - I va II tur xatolari, 1% va 10% kesishda;
  - ln Z-score benchmark sifatida.
- §1.3/2.1 – CAMEL va bozorga asoslangan indikatorlar taqqosi (p.79–81). Bu O‘zbekistonda listingdagi banklar kamligi sababli buxgalteriya indeksini tanlashni asoslaydi.
- §3.1 – CR = (E + LLR − NPL)/TA ni aktivlar blokiga qo‘shimcha "NPL qamrovi" indikatori sifatida kiritish mumkin. Bu IFRS 9 ECL bilan bog‘liq.
- §1.1 – Thakorning to‘rt belgisi "sog‘lom moliya tizimi" va barqarorlik ta’riflari bo‘limi uchun. Barqarorlik "bankrotlik nolga teng" degani emas.
- §3.3 – Deng va boshqalarning regulyatsiya indekslarini (CAPS, SUPP, MARD) GMM modelida institutsional o‘zgaruvchi sifatida ishlatish mumkin.

**4. Iqtiboslar**
- "the Z-score is a reliable and parsimonious measure of bank distress because while it requires few data, it does not lose predictive power" (p.93).
- "the Z-score could be used by supervisors as an early-warning tool to detect the symptoms of bank distress and adopt prompt corrective actions (PCA)" (p.93).
- "supervisors, from a prudential perspective, are primarily concerned about missing a distressed bank" (p.84–85).
- "It is a crisis that looks like a liquidity crisis, but it is in fact a solvency crisis" (p.11).
- "A healthy financial system is not one that is never susceptible to a crisis" (p.13).
- "This confuses capital requirements with cash asset reserve requirements" (p.14).
- "These results highlight the importance of combining policies which aim to foster financial stability with policies which promote financial intermediation" (p.125).

**5. Moslashtiriladigan jadval/formulalar**
1. Z-score formulasi (5.2) va probit modeli (5.1), p.83.
2. O‘zgaruvchilar va kutilgan ishoralar, Table 5.1 (p.85).
3. Prognoz va xatolar matritsasi, Table 5.5 (p.94–95).
4. ln_Z sog‘lom va distressdagi banklarda, Table 5.2 (p.86).
5. Metafrontier formulasi (6.3) va o‘zgaruvchilar, Table 6.2 (p.116–117).
6. Thakorning to‘rt belgisi (p.16).
