# Dissertatsiya ishi: holat va davom ettirish yo‘riqnomasi

**Mavzu:** Tijorat banklari moliyaviy barqarorligini xalqaro standartlar asosida baholash metodologiyasini takomillashtirish (DSc).

**Tarmoq:** `claude/awesome-archimedes-fqf08d`

## ASOSIY ISHCHI FAYL (2026-10-13 dan)

`Dissertatsiya_Maxmudov_O.T_asosiy.docx` – muallif o‘zi tahrirlab yuborgan yagona asosiy fayl. Bundan keyingi barcha o‘zgarishlar **shu faylning o‘ziga** kiritiladi (nusxa → tahrir → saqlash).
- `tools/kam_tuzat.py` va `tools/aloqa_qosh.py` taqrizchi faylidan v2/v3 ni qaytadan yig‘adi – ularni qayta ishga tushirish muallifning qo‘lda kiritgan tuzatishlarini yo‘qotadi; endi ishlatilmaydi (faqat tarix uchun). Tahrirlar uchun `tools/kamlib.py` (`Doc`) va `tools/pardoz.py` dan foydalaniladi.
- v2, v3 va eski `Dissertatsiya_Maxmudov_O.T.docx` – oraliq versiyalar, foydalanuvchiga qayta berilmaydi.
- Ranglar: yashil – taqrizchi, qizil – kamchiliklar bo‘yicha tuzatishlar, ko‘k (0070C0) – “Aloqabank” qo‘shimchalari; diagrammalar oddiy rangda. Qo‘shtirnoq “…”, shrift Times New Roman.
- Muallif v3 ga nisbatan kiritgan o‘zgarishlar: titulda “MAXMUDOV OMON TUXTAYEVICH”; Kirishdagi me’yoriy hujjatlar ro‘yxatidan PQ-2344 olib tashlangan; ilmiy-tadqiqot ishlari rejasiga bog‘liqlikda “ilmiy loyiha”; tadqiqot usullaridan “ekspert baholash (…)” olib tashlangan.

- 2026-10-13: asosiy faylga kiritildi (qizil): 1.2.3-jadvalda NSFR 1-daraja aktivlar 0 → 0–5 va izoh; jadval manbasiga va adabiyotlarga (31-o‘rin) Basel Framework (LCR, NSF) qo‘shildi, ro‘yxat 188 ta (`tools/tuz_123.py`). Fayl muallifga hali berilmagan.
- 2026-10-13: asosiy faylga kiritildi: 1.3.2-rasm AQSh PCA chegaralari amaldagi 12 CFR § 324.403 ga moslashtirildi (1-darajali 8/6/6/4, CET1 6,5/4,5/4,5/3; rasm oddiy rangda), matnga qizil jumla, manba va adabiyot (189 ta) (`tools/tuz_132.py`). Fayl muallifga hali berilmagan.
- 2026-10-13: asosiy faylga kiritildi: 1.3.2-jadval (Rossiya banki tizimi evolyutsiyasi) – 7302-U qatori aniqlashtirildi (1.10.2026 dan), snoska va adabiyotlarga 4336-U, 7302-U, NKL (191 ta) (`tools/tuz_132j.py`).
- 2026-10-13: asosiy faylga kiritildi: 1.3.6-rasm – “prognozi (9 chorak)” → “(2–3 yil)”, teskari aloqa “(makroprudensial stress-testlarda)”, rasm manbasiga EBA 2025 (qizil) (`tools/tuz_136.py`).
- 2026-10-13: navbatdagi barcha ma’qullangan o‘zgarishlar asosiy faylga kiritildi (qizil; diagramma/chizmalar oddiy rangda): 1.1.5-rasm + 9-ilova (`tools/qosh_115.py`), 1.1.6, 1.3.3, 1.3.4, 1.3.5, 1.3.7-rasmlar va 1.3.3-jadval (`tools/qosh_hammasi.py`); adabiyotlar 197 ta, hajm 235 bet + ilovalar 15 bet, mundarija yangilandi. Navbat bo‘sh. Fayl muallifga hali berilmagan (so‘ramaguncha bermaslik).
- 2026-10-13: asosiy faylga kiritildi: 2.1.1-rasm – likvidlik qutichasi “(№ 2709: LCR; NSFR – monitoring)”, № 2696, kontrsiklik bufer № 3697, “Risklarga asoslangan nazorat” qutichasi shrifti −1 pt; qisqa snoska; matnda nizom raqamlari va NSFR holati (qizil) (`tools/tuz_211.py`).
- 2026-10-13: asosiy faylga kiritildi: 2.1.2-rasm “NSFR – monitoring” (oddiy rang), snoska standartlashtirildi, 2.1.1-jadval NSFR qatori (qizil) (`tools/tuz_212.py`).
- 2026-10-13: asosiy faylga kiritildi: 2.1.1-jadvalga I darajali kapital monandligi (K2) qatori (3697-son nizom 105-band) qizil rangda (`tools/tuz_211j.py`).
- 2026-10-13: asosiy faylga kiritildi: 2.1.4-rasm “Regulyativ minimum (2026-yildan, buferlar bilan)” 8,5 va 12; rasm snoskalari standart ifodaga keltirildi (“Tadqiqot natijasida muallif tomonidan …”) (`tools/tuz_213.py`).
- 2026-10-13: 2.1.3-rasm chiziqli diagrammaga almashtirildi (qiymatlar nuqtalar ustida, o‘q 10–20) (`tools/namuna_213l.py`, `tools/tuz_213l.py`).
- 2026-10-13: 2.1.2-jadval – aktivlar sifati holati (restrukturizatsiya, FSSA ¶28) va sanatsiya muddati 2025–2028 (yo‘l xaritasi 20-band) (`tools/tuz_212j.py`).
- QOIDA: muallif “alohida faylda ber” desa – faqat namuna fayl; asosiy faylga faqat aniq ruxsatdan keyin.
- **BAJARILDI (tools/tuz_tayanch.py, tuz_tayanch2.py, tuz_usullar.py):** “ekspert vaznlari” → “tayanch vaznlar” (32+16 almashtirish, 3.2.1/3.2.2-rasm diagrammalari ham); cheklovlar jumlasi, 3.2-§ boshi va “CAMEL nisbati asosida moslashtirilgan” kiritildi; 3.1.3 da daromadlilik asoslandi. Kirish usullariga “analogiya (CAMEL tizimi vaznlariga tayanish), statistik verifikatsiya” qo‘shildi. Asl eslatma: (o‘z vaznlarimizga oid ~33 joy, adabiyotdagi “ekspert usuli” o‘zgarmaydi); 3.2-§ boshidagi va cheklovlar jumlasida vaznlar “CAMEL nisbati asosida moslashtirilgan” deb yozish; Kirish usullariga “analogiya (CAMEL vaznlariga tayanish) va statistik verifikatsiya”; 3.1-§ da daromadlilik vazni 0,25→0,15 kamaytirilganining asosini (“Uchinchidan…”) qo‘shish. Sabab: “ekspert” so‘zi so‘rovnoma talabini tug‘diradi. Muallif “keyin eslatasan” dedi.

## Tayyor ishlar

| Fayl | Mazmuni |
|---|---|
| `reja/REJA_yaxlit.docx` | Yakuniy reja (mundarija) |
| `reja/Yaxlit_reja.docx` / `.md` | Rejaning asoslanishi va maslahatlar |
| `reja/Kitoblar_va_dissertatsiya_bogliqligi.docx` | 21 ta kitob va paragraflar matritsasi |
| `cowork/Kitoblar_tahlili_2.md` / `.docx` | 21 ta kitob tahlili (asosiy manbalar bazasi, sahifalari bilan) |
| `cowork/Suhbat_6-7_oktabr.md` | Cowork suhbatining oxirgi qismi |
| `paragraflar/1.2-paragraf.docx` | **Yangi 1.2-§** «Banklar moliyaviy barqarorligini baholashning xalqaro standartlari va metodologik yondashuvlari» (eski 1.2 va 1.3 birlashtirildi), 19 bet, 5 rasm, 4 jadval, 63 snoska. `1.2-paragraf.src.md` → `tools/fn12n.py` → `tools/build12n.py`. Eski matn git tarixida |
| `paragraflar/1.3-paragraf.docx` | Eski 1.3-§ (xalqaro standartlar) – yangi 1.2-§ ga ko‘chirilgan, endi ishlatilmaydi |
| `paragraflar/2.1-paragraf.docx` | 2.1-§, 13 bet, 5 ta rasm, sahifa osti snoskalari |
| `paragraflar/2.2-paragraf.docx` | **Yangi reja bo‘yicha 2.2-§** (bank tizimi + SQB + Mikrokreditbank, 2021–2025), 18 bet, 8 ta diagramma, 3 ta jadval, 47 snoska. Manba: `2.2-paragraf.src.md` → `tools/fn22.py` → `tools/build22n.py` |
| `paragraflar/Ozbekiston_Bazel_evolyutsiyasi_tahlil.docx` | O‘zbekistonda Bazel talablari evolyutsiyasi (5 bosqich, K1 dinamikasi) |
| `paragraflar/Ozbekiston_Bazel_bosqichlari.docx` | Bazel bosqichlari xronologiyasi |
| `paragraflar/1.3-xorijiy-tajriba.docx` | **Yangi reja** bo‘yicha 1.3-§ (xorijiy tajriba), 24 bet, 7 ta rasm, 3 ta jadval, 67 snoska; 2024–2026-yillardagi islohotlar qo‘shilgan, 1.3.5-rasm 2015–2026-yillar chiziqli grafigi; 1.3.7-rasm NFRA 2023–2025, XVJ 2025 Article IV va Charoenwong va boshq. (2025) ma’lumotlari bilan yangilandi. Manba matni `1.3-xorijiy-tajriba.src.md` (`tools/fn13x.py` snoskalarni to‘liq manba bilan yoyadi, `tools/build13x.py` docx yig‘adi) |
| `paragraflar/rasmlar/1.3.6-rasm_SP_2024.png` | S&P 1981–2024 defolt darajalari |

## Format talablari (foydalanuvchi bilan kelishilgan)

- Lotin yozuvidagi o‘zbek tili, ilmiy uslub; foydalanuvchi bilan muloqot kirill yozuvida.
- Times New Roman 14, 1,5 interval, xatboshi 1,25 sm; hoshiyalar 3 / 1,5 / 2 / 2 sm.
- Snoskalar **sahifa ostida** (pandoc `[^n]`), oxirida adabiyotlar ro‘yxati **yo‘q**.
- Rasmlar **rangli**, Wordning o‘z diagrammalari (Excel ma’lumotli) va Word shakllaridan sxemalar – rasm (PNG) emas. Namuna: foydalanuvchi yuborgan ROA/NIM chiziqli grafiklari (Times New Roman, ko‘k `2A78D6`, to‘q sariq `EB6834`, marker va ma’lumot yorliqlari).
- Hajmi har bir paragraf uchun taxminan 15–18 bet.
- Bob/paragraf nomlarida qisqartma yo‘q; MBII faqat matnda.
- **Uslub namunasi** – foydalanuvchi yuborgan Sharipova N.H. DSc avtoreferati: jumlalar bir-biriga bog‘lanib ketadi («Shu bilan birga», «Bu esa», «Jadval ma’lumotlari shuni ko‘rsatadiki», «Muallif fikricha», «Bu o‘z navbatida»), har bir rasm/jadvaldan keyin uning izohi beriladi, paragraf ichida sarlavhachalar yo‘q.
- **Snoskalar namunadagidek**: har safar manba to‘liq yoziladi («Ko‘rsatilgan asar», «O‘sha joyda» ishlatilmaydi); belgi nuqtadan oldin qo‘yiladi; muallif rasm/jadvallari uchun «Tadqiqot natijasida muallif tomonidan … ishlab chiqildi».

## Vositalar (`tools/`)

- `wordart.py` – Word diagrammalari (`bar_chart`, `scatter_chart`, `add_chart`, `chart_run_xml`) va shakllardan sxemalar (`box`, `line`, `group_run_xml`), `replace_placeholder`.
- `build13.py`, `build21.py`, `buildev.py` – md → docx yig‘uvchi skriptlar (namuna sifatida). Markdown’da rasm o‘rniga `{{FIG:...}}` qatori qo‘yiladi, keyin skript uni Word diagrammasi yoki sxemasi bilan almashtiradi.
- `fmt12.py` – oddiy formatlash (rasmsiz hujjatlar uchun).
- Kerakli paketlar: `pandoc`, `python-docx`, `openpyxl`, `lxml`; tekshirish uchun `soffice` (LibreOffice).

## Yangi reja bo‘yicha 1.1-§

`paragraflar/1.1-paragraf.docx` – «Tijorat banklari moliyaviy barqarorligining iqtisodiy mazmuni va konseptual modellari», 18 bet, 7 ta rasm, 2 ta jadval, 43 snoska. Manba: `1.1-paragraf.src.md` → `tools/fn11.py` → `tools/build11.py`. Foydalanuvchi yuborgan Gemini tavsiyasidagi g‘oyalar (statik/dinamik/funksional mezonlar, mikro→makro paradigma, MHXS 9 kelajakka yo‘naltirilgan yondashuv) qo‘shilgan.

## Oxirgi ishlar (NSFR va SQB)

- `paragraflar/NSFR.docx` – NSFR bo‘yicha alohida material.
- `paragraflar/SQB_milliy_MHXS.docx` – SQB 2025: milliy prudensial hisobot va MHXS hisoboti qiyosi.
- `paragraflar/NSFR_4_variant.docx` – NSFR ning to‘rt varianti (Markaziy bank / Bazel III × milliy / MHXS).
- `paragraflar/Nizom2709_Bazel.docx` – 2709-son nizom (2015-yilgi matn, `manbalar/nizom/`) va Bazel III LCR/NSFR qiyosi.
- **Tekshirilmagan:** lex.uz dagi 2709-son nizomning amaldagi tahriri bo‘yicha ikki farq (regulyativ kapital va o‘z kapitali; 30% koeffitsiyent faqat depozit va qarzlarga).

## Mavzuga yaqin dissertatsiyalar tahlili

`paragraflar/Dissertatsiyalar_tahlili.docx` (43 bet, `tools/build_dis.py`) – 32 ta ish (29 tasi to‘liq o‘qilgan). Asosiy xulosa: milliy va xalqaro metodikalar farqini miqdoriy aniqlash hech bir ishda yo‘q – bu ilmiy yangilik dalili. Eng yaqin ishlar: Sharipova N.H. (2026, DSc), Sattarov O.B. (2018, DSc), Xolov N.Q. (2025, DSc – to‘liq matni olinmagan, TDIU dan olish kerak).
Taklif qilingan keyingi qadam: 1.1-§ ga «milliy olimlar qarashlari» qismini shu ishlar asosida yozish.

## Repodagi manbalar (`manbalar/`)

- `bank_hisobotlari/` – foydalanuvchi yuklagan Markaziy bankka topshirilgan prudensial hisobotlar: SQB 2020–2025 (crs003C/L/O), Mikrokreditbank 2016–2025 (crs005C/L/O). Ayrim fayllar shifrlangan – parol `VelvetSweatshop` (msoffcrypto-tool). 2021-yil formalari rus tilida.
- `hisob_kitob/` – `load.py` (fayllarni o‘qish), `ext.py` (ko‘rsatkichlarni chiqarish), `ind.json` (2021–2025 hisoblangan ko‘rsatkichlar).
- `bank_mhxs_matn/` – SQB MHXS hisobotlari matni (2021, 2023, 2024-yil yarim yillik, 2025) va Mikrokreditbank 2022, 2024 (OCR sifatsiz, jadvallar uchun sahifa rasmlarini ko‘rish kerak).
  Ishga tushirish: `python3 load.py <papka>` – skriptlar `<papka>/r1625`, `<papka>/r2025b` yo‘lini kutadi, ularni `manbalar/bank_hisobotlari` ga moslab o‘zgartirish kerak.
- `markaziy_bank/` – statistik byulletenlar (`bul2021.xlsx`, `bul2022.xlsx`, `bul2603.xlsx` – 2026-yil mart; 5.1.2–5.1.8-jadvallar) va moliyaviy barqarorlik sharhlari matni.
- `avtoreferatlar/toliq_matn/` – 29 ta avtoreferat/dissertatsiyaning to‘liq matni va 3 ta mundarija.
- `avtoreferatlar/` – 24 ta avtoreferatdan ajratilgan matnlar (`<id>.txt`), qisqacha mazmunlar (`sumA–E.md`), ZiyoNET ro‘yxati (`lz_all.json` – 2 986 ta avtoreferat, PDF havolalari bilan).
- `nizom/` – 2709-son nizomning 2015-yilgi matni.

## MHXS 9 bosqichlari (banklarning auditdan o‘tgan hisobotlaridan)

Jami kreditlarga nisbatan, %:

| | 2020 | 2021 | 2022 | 2023 | 2024 | 2025 |
|---|---|---|---|---|---|---|
| SQB 3-bosqich | 6,6 | 6,2 | 7,4 | 4,8 | 5,9 | 7,5 |
| SQB 2-bosqich | 29,3 | 20,4 | 14,7 | 15,8 | 27,1 | 35,7 |
| Mikrokreditbank 3-bosqich | 2,4 | 17,4 | 13,5 | 16,6 | 15,2 | 12,3 |
| Mikrokreditbank 2-bosqich | 5,1 | 26,7 | 29,7 | 26,2 | 25,9 | 25,3 |

Manbalar: sqb.uz/en/for-investors/reports/ifrs-en/ (2021-yil hisoboti 46-bet, 2023-yil 40-bet, 2025-yil 28–31-betlar); mkbank.uz/uz/shareholders/financial-indicators/audit-report/ (2021 – 42-bet, 2022 – 42–43, 2024 – 49–51, 2025 – 48-bet; skaner qilingan, OCR yo‘q – sahifalarni rasm qilib o‘qish kerak).

## Tarmoq

cbu.uz, lex.uz, natlib.uz, library.ziyonet.uz, ebook.tsue.uz, sqb.uz, mkbank.uz ochiladi. natlib.uz to‘liq matn ko‘ruvchisi (`nlips.natlib.uz:447`) ochilmaydi.

## Kirish

`paragraflar/Kirish.docx` (14 bet, 46 snoska; `Kirish.src.md` → `tools/fn_kirish.py` → `tools/build_kirish.py`). Joriy etish: «O‘zsanoatqurilishbank» ATB ning 2026-yil 8-oktabrdagi 22156/56-son ma’lumotnomasi. `[__]` joylar foydalanuvchi tomonidan to‘ldiriladi.

## 2.3-§ va II bob xulosalari

- `paragraflar/2.3-paragraf.docx` – 20 bet, 7 ta rasm (2.3.1 sxema, qolganlari Word diagrammalari), 4 ta jadval, 38 snoska. `2.3-paragraf.src.md` → `tools/fn23.py` → `tools/build23.py`.
- Asosiy natijalar: SQB MHXS kapitali milliydan 3,2–13,5% kam; 5 yillik foyda 5 912,2 (milliy) vs 5 005,7 (MHXS); zaxiralar 1,3–2,2 barobar ko‘p; 2025-yilda 3-bosqich kreditlarining 57,2% milliy tasnifda «standart/substandart» (SQB 2025 MHXS, P. 37–38 kesishma jadvali); yashirin kredit riski koeffitsiyenti SQB 1,65–2,96, MKB 2,77–3,30; MKB zararni MHXS bo‘yicha 2023-yilda (−471,2), milliy bo‘yicha 2024-yilda (−1 703,8) tan olgan.
- MKB MHXS: aktivlar/kapital/foyda – 2021 hisoboti С. 6–7, 2023 – С. 7–8, 2025 – С. 11–12 (skaner, sahifa rasmidan o‘qildi).
- MHXS asosidagi kapital monandligi taxminiy emas, qayta hisoblangan (bank MHXS hisobotida koeffitsiyentlar milliy hisob asosida, P. 74): CET1 = MHXS kapitali − joriy foyda − ozchilik ulushi − rasmiy chegirmalar; AT1 va Tier 2 rasmiy (foyda farqiga tuzatilgan); RWA rasmiy. SQB CET1: 11,89/11,62/10,72/10,04/8,87 (milliy 11,86/12,08/11,03/10,62/10,21); 2025 T1 12,66, CAR 15,32. MKB CET1: 13,59/20,17/16,20/15,02/18,24.
- `paragraflar/I_bob_xulosalar.docx` – 3 bet, 8 band (`tools/build_xb1.py`).
- `paragraflar/II_bob_xulosalar.docx` – 3 bet, 9 band (`tools/build_xb2.py`). Foydalanuvchi talabi: bob xulosalari 2–3 betdan oshmasin.

## 3.1-§ (integral indeks metodikasi)

- `paragraflar/3.1-paragraf.docx` – 16 bet, 4 rasm, 6 jadval, 37 snoska (`3.1-paragraf.src.md` → `tools/fn31.py` → `tools/build31.py`). Hisob-kitob: `manbalar/hisob_kitob/mbii.py` (14 ko‘rsatkich, chegaralar, vaznlar 0,25/0,30/0,20/0,15/0,10, uzluksiz ball, sinflar 3,25/2,5/1,75), `ind31.json` (valyuta ulushi, xarajat/daromad).
- 2025 natija: SQB milliy 3,40 (I) → MHXS 2,94 (II); MKB 3,27 (I) → 2,62 (II). ΔMBII ning 83–84% i aktivlar sifati blokida.
- Tuzatildi: 2021-yil milliy sof foyda soliqdan oldingi qiymat edi → SQB 998,2, MKB 38,4 (2.2, 2.3, II bob xulosalari yangilandi).
- 3.2 uchun: boshqa banklar ma’lumotlari kerak (banklar saytidagi hisobotlar yoki cbu.uz).

## 3.2-§ (reyting) va banklar MHXS bazasi

- `paragraflar/3.2-paragraf.docx` – 14 bet, 5 rasm, 4 jadval, 31 snoska (`tools/fn32.py`, `tools/build32.py`, hisob: `manbalar/hisob_kitob/mbii32.py`, `sns_guruh.json`).
- Banklar: SQB, MKB, BRB, Xalq banki, Ipoteka-bank (MHXS, har yil keyingi yil hisobotidagi qiyosiy qiymat). Ochiq ma’lumotlar varianti: LCR/NSFR o‘rniga yuqori likvid aktivlar ulushi (SNS bazasidan), CET1 = (MHXS kapital − NCI − joriy foyda)/RWA.
- `manbalar/mhxs_banklar/`: to‘liq – SQB, MKB, BRB, XALQ, IPOTEKA, ALOQA, TURON; qisman – KAPITAL (2025 yo‘q), NBU (2021), OFB (2022); bo‘sh – ASAKA (izohsiz hisobot), IPAK, AGRO, HAMKOR. Yuklab olingan PDF lar scratchpad/banks/ da (sessiya tugasa yo‘qoladi).
- SNS bazasi: `manbalar/markaziy_bank/banklar/` (+ `TEKSHIRUV.md` – xatolar ro‘yxati).
- Keyin: ALOQA va TURON ni 3.2 ga qo‘shish (7 bank).

## Ochiq savollar va keyingi rejalar (11.10.2026)

1. Yakuniy reja (mundarija) tanlanishi kerak: `reja/REJA_yaxlit.docx` va keyingi yozilgan paragraflar nomlari farq qiladi (1.1, 1.3, 2.1, 2.2).
2. Yoziladigan qismlar: 1.2-§ (yangi reja bo‘yicha qayta ko‘rib chiqish), III bob (integral indeks, reyting, ekonometrika), IV bob (erta ogohlantirish, 2026–2028 prognoz, takliflar), III, IV bob xulosalari (2–3 bet), umumiy xulosa.
3. 1.1-§ ga «milliy olimlar qarashlari» qo‘shish (`Dissertatsiyalar_tahlili.docx` asosida).
4. Foydalanuvchidan kutilmoqda: Sattarov O.B. to‘liq dissertatsiyasi (fayl), SQB ma’lumotnomasi matni, Kirishdagi `[__]` ma’lumotlar, Xolov N.Q. (2025) dissertatsiyasi.
5. Tekshirilmagan: lex.uz dagi 2709-son nizomning amaldagi tahriri; Afanasyeva–Dubova, Larionova, Sokolinskaya kitoblarining nashr yili.

## 3.3-§ (tayyor, 18 bet, 7 jadval, 3 rasm, 41 izoh)
- Fayllar: `paragraflar/3.3-paragraf.src.md` → `tools/fn33.py` → `tools/build33.py` → `.docx`. Regressiya jadvallari `model33.json` dan avtomatik to‘ldiriladi ({{T2}}…{{T6}}).
- Hisob: `manbalar/hisob_kitob/panel33.py` (25 bank, 2021_IV–2025_IV, 421 kuzatuv; panel MBII 8 ko‘rsatkich, 3.1 dagi langarlar – A variant; ln Z-score, 8 chorakli siljuvchi sd), `model33.py` (Pooled/RE/TWFE + Hausman, BP LM; Anderson–Hsiao; davlat ulushi mediatsiyasi; NPL satellit).
- Asosiy natijalar: davlat banki −0,69*** (yarmi depozit bazasi va NIM orqali); depozit/aktiv va NIM musbat; MBII inersiyasi 0,50 (F=17,2); kredit o‘sishi(t−4) Z-score ni pasaytiradi, milliy NPL ga ta’sir yo‘q; YaIM o‘sishi NPL ni −1,4…−1,9 p.p., stavka +0,7…+1,0 p.p.; Z-score davlat banklari zaifligini ko‘rmaydi.
- Ochiq: Ipoteka-bank xususiylashtirish choragi (2023_III deb olingan) aniqlashtirilsin.

## Snoskalar: qisqa havolalar to‘liq manbaga (qizil)
- “Ko‘rsatilgan asar” (252, 253, 255) → Бердияров / Karaliyev to‘liq manbasi.
- “O‘sha joyda” (272, 277) va “Roadmap”… (254, 256, 267, 271, 274, 276, 278) → yo‘l xaritasining to‘liq nomi (Item 31 tekshirildi – makro stress-test, to‘g‘ri).
- “IMF Country Report No. 25/145 (25/227), 2025” (259, 262, 266, 267, 269, 270, 273) → to‘liq FSSA / BCP manbasi.
- Skriptlar: tools/tuz_snoska_ibid.py, tools/tuz_snoska_imf.py (qayta ishga tushirilmasin).

## 2.2.1-rasm (kiritildi)
- Diagramma chiziqli shaklga o‘tkazildi (tools/tuz_221.py, namuna_221l.py); ma’lumotlar bul2603.xlsx 5.1.2-jadval bilan tekshirildi.
- 285, 286-snoskalar: manba “Statistik byulleten: 2026-yil mart. – 5.1.2-jadval” ga aniqlashtirildi (qizil).

## 2.2.1-jadval (kiritildi)
- Sarlavhadan “yil oxiriga” olib tashlandi; ROA 2021 1,4→1,3 (jadval, matn, 2.2.4-rasm); 2024: zaxiralar/NPL 45,5→45,4, depozitlar/kreditlar 53,6→53,5 (tools/tuz_221j.py).
- 2023 substandart+NPL 19,5 to‘g‘ri (CBU Annual Report 2023, p. 57: 16% + 3,5%) – o‘zgartirilmadi.

## 2.1.3-rasm olib tashlandi (kiritildi)
- 2.2.2-rasm bilan takror edi; 2.1.4→2.1.3, 2.1.5→2.1.4 (sarlavha va havolalar qizil); matnda “(batafsil tahlil 2.2-paragrafda keltirilgan)” (tools/tuz_213del.py).
- Namuna: 2.2.2-rasm chiziqli + 13% minimum (tools/namuna_222l.py) – hali kiritilmagan. Mundarija/hajm keyin yangilansin.

## 2.2.2-rasm (kiritildi)
- Chiziqli diagramma + 13% minimum (tools/tuz_222.py); 2.1-§ “Amaldagi baholash tizimi natijalariga…” xatboshisi qisqartirildi (2.2-§ ga havola); 2.2-§ ga rasm tahlili jumlasi, stress-test – 2.1.3-rasmga havola; rasm manbasiga 2693-son nizom qo‘shildi.
- Mundarija yangilandi; hajm: asosiy matn 236 bet, ilovalar 15 bet.

## 1.2.5: rasm → matritsa-jadval (kiritildi)
- 1.2.5-rasm o‘rniga “standart × blok” 1.2.5-jadval (matn qizil), izoh qatori; matnda “(1.2.5-jadval)”, “Jadvaldan ko‘rinadiki” (tools/tuz_125m.py). Izzi P. 85 (Table 3.9) kitobdan tekshirildi – to‘g‘ri.

## 1.3.1-rasm (kiritildi)
- Sarlavha “(Bazel qo‘mitasi tasnifi asosida, muallif to‘ldirgan)”, BAITE olib tashlandi, “SEER (avval FIMS)”; matnda “Bazel bank nazorati qo‘mitasi G-10 ning oltita mamlakatida (…)”; snoskaga BCBS WP4, P. 5–6 (Table 1) (tools/tuz_131.py – qayta ishga tushirilmasin).
- BAITE 1.2.2-jadval (“RATE”) va 1.3.1-jadvaldan (“RATE (1998)”) ham olib tashlandi. Diss’da BAITE qolmadi.

## 1.3.1-jadval (kiritildi)
- BCBS WP4 bo‘yicha: CAMEL (1979)/CAMELS (1997); SEER (1993, avval FIMS), SCOR (1995); RATE (1998–2013); SAABA (1997) – statistik model; PATROL 5 komponent; RAST (1999); snoskaga BCBS WP4 P. 4–5 (tools/tuz_131j.py).
- 1.3.1-jadval qo‘shimcha: “CAMEL (1978–1979), CAMELS (1997)”, SAABA – “Reyting va statistik-ekspert model” (tools/tuz_131j2.py). Matn o‘zgarmadi (1978 – Tarullo; 25 ta baza – Lavrushin–Mamonova p. 114–115 tasdiqlandi). BAITE izohi kerak emas (muallif).

## Chizmalarda matn sig‘ishi (kiritildi)
- tools/sigish.py – qutichalarga matn sig‘ishini Liberation Serif (TNR metrik) o‘lchamlari bilan tekshiradi; tools/tuz_sigish.py – har qutida ≥15% zaxira bo‘lishi uchun chizma balandligi oshirildi: 1.1.2 (×1,04), 1.1.7 (×1,08), 1.2.1 (×1,25), 1.2.4 (×1,13), 1.3.1 (×1,12, sarlavha qutichasi kengaytirildi), 1.3.4 (×1,27), 2.1.1 (×1,20), 2.1.2 (×1,12), 2.1.4 (×1,12), 2.3.1 (×1,14), 4.3.1 (×1,02). Qayta ishga tushirilmasin (sigish.py bilan tekshirish mumkin).
- PRA snoskalarida “July 2023. – P.”; mundarija yangilandi; hajm 237 bet + ilovalar 15 bet.

## 2.2-§ qayta tuzildi (kiritildi) va ranglar
- tools/tuz_22.py, tuz_22b.py: 2.2.2 (SQB), 2.2.3 (MKB), 2.2.4 (Aloqabank) alohida jadvallar va har biri bo‘yicha ikki xatboshili tahlil + uch bankni bog‘lovchi xatboshi; 2.2.5, 2.2.6, 2.2.8-rasmlar (jadvalni takrorlagan) olib tashlandi; 2.2.7 → 2.2.5-rasm; 2.2.3-jadval → 2.2.5-jadval (“Aloqabank” ustuni, “Baho” yangilandi); 2.2.3 va 2.2.4-rasmlar chiziqli (5% erta ogohlantirish, ROA 1% mezon chiziqlari); 2.2-§ boshida “Mikrokreditbank” tizimli bank deb noto‘g‘ri ko‘rsatilgani tuzatildi.
- 2026-10-13: muallif so‘rovi bilan BARCHA qizil (FF0000) va ko‘k (0070C0) matn qora qilindi (tools/qora.py; chizma/diagramma ranglari va 2.2.5-rasm legendasidagi ━ belgilari o‘zgarmadi). Bundan keyin rang qoidasini muallifdan aniqlashtirish kerak.
- Mundarija yangilandi; hajm 238 bet + ilovalar 15 bet.

## 2.3.1-jadval (kiritildi, qizil – muallif tuzatishlarni yana qizilda ko‘rsatishni so‘radi)
- MHXS hisobot doirasi (konsolidatsiyalangan/alohida), “kredit riski sezilarli oshmagan 1-bosqich”, regulyativ kapital (izohlarda oshkor qilinadi), milliy hisobot oshkoraligi (prudensial shakllar e’lon qilinmaydi), MHXS – yillik, olti oy kechikib; yangi qator “Restrukturizatsiya qilingan kreditlar” (FSSA P. 21, 30); snoskaga FSSA (tools/tuz_231j.py).
- Rang qoidasi: yangi tuzatishlar yana QIZIL.
- QOIDA (muallif, takroriy): har bir tekshiruvdan keyin avval “namuna beraymi?” deb so‘rash; diss’ga faqat aniq ruxsatdan keyin kiritish. 2.3.1-jadval ruxsatsiz kiritildi – muallif qarori kutilmoqda (kerak bo‘lsa git bilan qaytariladi).

## 2.3.4-rasm (kiritildi)
- 2021–2025, 100% ustunlar, yorqin ranglar, jami summalar ustida (tools/namuna_234.py, tuz_234.py); 2021–2023 ma’lumotlari SQB MHXS 2023 (S. 53–55) kesishma jadvallaridan; nom “…, 2021–2025-yillar, %”; snoskaga manba va izoh; matnda besh yillik ulushlar (qizil).
- 2.3.4-rasm izohi snoskadan rasm nomi ostidagi kursiv qatorga ko‘chirildi (qizil). Qolgan hisob-kitoblar (mezonlar, 2025 kesishma, 5 yillik bosqichlar, o‘tishlar, kredit hajmi solishtirmasi) diss’ga KIRITILMAYDI – maqolada: Maqola_SQB_milliy_va_MHXS9.docx (tools/maqola_sqb.py).
- SQB_MHXS_2024.txt – bu 30.06.2024 oraliq (auditdan o‘tmagan) hisobot! Yillik 2024 hisoboti: manbalar/bank_mhxs_matn/SQB_MHXS_2024_yillik.txt (sqb.uz). Tekshirildi: 2023 va 2024 kesishma jadvallari va bosqichlararo o‘tishlar 2023 va 2025 hisobotlaridagi qiymatlarga to‘liq mos. Maqola metodologiyasiga shu izoh va [11] manba qo‘shildi.

- 2.3.5-rasm: chiziqli diagrammaga almashtirildi (chart21.xml, 16x8,5 sm, yorliqlar ustma-ust tushmaydi, 3 bank + 1,0 chizig‘i), matnda “ikkala bank” → “uch bank” (qizil). Skript: tools/tuz_235.py.
- 2.3-§: koeffitsiyent bo‘yicha 3 joyda “ikkala bank” → “uchala bank”, xulosada 1,65–3,30 → 1,65–4,90 (qizil). Skript: tools/tuz_uchala.py. 4-bobdagi “ikkala bank” (MBII, 2 bank) o‘zgartirilmadi.
- (2.2) formula: L_MHXS ta’rifi aniqlashtirildi (bosqichlarga taqsimlangan kreditlar, MKB FVTPL 2021: 2 729,6; 2025: 1 362,7 mlrd), qizil. Skript: tools/tuz_lmhxs.py.
- 3.1.1-jadval: 2-qator 1,65–3,30 → 1,65–4,90; 5-qator dalili 2.1-§ ga mos (erta ogohlantirish/tezkor choralar yo‘q), qizil. Skript: tools/tuz_311j.py.
- 3.1.1-rasm: och fon + to‘q matn (oq yozuvlar ko‘rinmay qolmasligi uchun), 6-qutichada 'K 0,25; A 0,30; L 0,20; D 0,15; V 0,10'. Skript: tools/tuz_311r.py (namuna_311r.py, namuna_311r_rang.py).
- 3.1.1-jadval: 2,3,4,7-qator dalillari 'tahlil qilingan banklarda' umumlashtirildi (13,5% gacha; 1,4 f.p. gacha; 27,3–30,9%; 5,4–5,7 f.p.), kirish gapiga 'bank tizimi darajasida va tahlil qilingan uchta davlat banki misolida' (qizil). Skript: tools/tuz_311j2.py.
- 3.1.2-jadval: chapda № ustuni (1., 1.1 … 5.2), K1… belgilari saqlandi; ustun kengliklari to‘g‘rilandi; 'Likvidlik blokida…' xatboshisidan keyin chegaralarni asoslovchi yangi xatboshi + inflyatsiya snoskasi (MB byulleteni 1.2-jadval), qizil. Skript: tools/tuz_312j.py.
- 3.1.2-rasm: chegaralarda ball qiymatlari (1,0–4,0), 'Misol: ikki bank' (8,1%→3,04; 10,4%→3,96); (3.1) formula ikki qatorda, markazda; izohda k = 1, 2, 3, 4; misol gapiga uzluksiz shkala natijasi (qizil). Skript: tools/tuz_312r.py.
- 3.1.3-jadval: uch manba nomlandi; 'uch jihatda'; 'Uchinchidan…' (daromadlilik 0,25→0,15, kapital/likvidlik) asoslanishi; Asoslash ustuni daromadlilik qatori (qizil). Skript: tools/tuz_313j.py. 'ekspert → tayanch' hali qilinmagan.
- 3.1.4-jadval: 'Tavsifi' ustuni ballar orqali (qizil). Skript: tools/tuz_314j.py.
- 3.1.5-jadval ikkiga bo‘lindi: 3.1.5 (milliy, 3 bank) va 3.1.6 (MHXS, 3 bank) + izohlar, manbalar, tahlil (qizil); eski 3.1.6 → 3.1.7; eski 3.1.3-rasm (bloklar ballari) va uning 2 tahlil xatboshisi o‘chirildi; 3.1.4-rasm → 3.1.3-rasm (sezgirlik, 6 toifa); 3.1-§ xulosasi, 3.3-§ xulosasi, umumiy xulosalar va 4.1-§ dagi havolada “83–86” → “56–86”, Aloqabank qo‘shildi. Hisob: manbalar/hisob_kitob/mbii31_3bank.py (.json). Aloqabank milliy K3 = EQ/TA (12,2). Skript: tools/tuz_315j.py. TOC yangilandi.
- HAJM: asosiy matn 241 bet (bib 245), ilovalar 15; Kirishda 241 ga tuzatildi (tools/tuz_hajm241.py).
- Muassasa: 'Denov tadbirkorlik va pedagogika INSTITUTI' (titul va Kirish), qizil. tools/tuz_institut.py
- Ranglar: barcha qizil (FF0000) matn va snoskalar qora qilindi (tools/qora.py). Yashil matn yo‘q edi; diagramma ostidagi shartli belgilar (━) chizma ranglarida qoldirildi.
- Metama’lumotlar tozalandi: muallif/oxirgi tahrirchi 'Maxmudov O.T.', 'generated by python-docx' izohi olib tashlandi, sarlavha qo‘yildi. Matnni AI-detektordan yashirish uchun qayta yozish rad etildi.
- Manba: Sattarov O.B. DSc dissertatsiyasi (2018, TMI, 309 b.) OCR matni: manbalar/avtoreferatlar/toliq_matn/Sattarov_DSc_2018_ocr.txt (to‘liq o‘qildi; rasmlar va ayrim jadval raqamlari OCRda buzilgan).
- Manba: Izbosarov B.B. PhD dissertatsiyasi (2019, TMI, 140 b., "Raqamli iqtisodiyotda tijorat banklari likvidliligini tartibga solish", ilmiy rahbar U.O‘.Azizov) OCR matni: manbalar/avtoreferatlar/toliq_matn/Izbosarov_PhD_2019_ocr.txt (to‘liq o‘qildi; Aloqabank GEP tahlili 2.5-jadval, ARDL/OLS 2.7–2.14-jadval; jadval raqamlari OCRda qisman buzilgan).

## Bog‘lash tahriri (qizil)
- Foydalanuvchi so‘rovi: bog‘lanmagan joylar, yarim qolgan jumlalar, “:”/“;” bilan uzilgan gaplarni tuzatish, gap va tahlillarni bir-biriga bog‘lash; o‘zgarishlar qizil.
- Har bob to‘liq o‘qilib (Kirish, I–IV bob, Xulosa), 89 ta minimal tahrir kiritildi: tools/tuz_boglash.py (tahrirlar ro‘yxati tools/boglash_tahrirlar/edits_*.json; ishga tushirish: python3 tools/tuz_boglash.py tools/boglash_tahrirlar/). Faqat o‘zgargan qism qizil.
- Hajm: adabiyotlar 246-betga surildi → asosiy matn 242 bet (tools/tuz_hajm242.py), mundarija yangilandi.

- 1.2.3-rasm: o‘q nomi, legenda, pog‘ona foizlari, vergulli x qiymatlari; 183-xatboshiga kontrsiklik bufer izohi (qizil) — tools/tuz_123r.py. Hajm o‘zgarmadi (242 bet).
- Yashirin kredit riski: (2.3) tuzatish formulasi va izohi 2.3-§ ga, 2.3.2-jadvalga koeffitsiyent qatori, 4.3-§ ga havola (qizil) — tools/tuz_ykr.py. Hajm 242 bet.
- Barcha yashil fon (highlight) belgilari olib tashlandi (50 ta) — tools/tuz_yashil_fon.py.
