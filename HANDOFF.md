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
- **Qo‘shish navbati (muallif “keyin hammasini birdan qo‘shasan” dedi):**
  1. 1.1.5-rasm (faqat 2021–2025, G-7, 12 bank) – muallif ma’qulladi; `tools/qosh_115.py` bilan qo‘shiladi (9-ilova, 4 manba, hajm/mundarija yangilanadi).
  2. 1.1.6-rasm: Correia, Luck, Verner (QJE 2026) AQSh 1959–2024 AUC – `Namuna_1.1.6-rasm_yangi.docx`, `tools/namuna_116.py`; muallif MA’QULLADI, navbatda (O‘zbekiston Z-score varianti rad etilgan – 4.1-§ ni takrorlaydi). Qo‘shilsa: Chiaramonte & Poli (2014) boshqa joyda keltirilmasa adabiyotlardan olib tashlash.
  3. 1.3.3-rasm: EBA Staff Paper No. 21 (2025), 176 bank, 2017–2023, tanlanmadan tashqari 2022 III – 2023 III – `Namuna_1.3.3-rasm_yangi.docx`, `tools/namuna_133.py`; muallif MA’QULLADI, navbatda. Qo‘shilsa adabiyotga EBA Staff Paper qo‘shiladi.
  4. 1.3.4-rasm: BAITE → RATE (1998, FSA), PRA PIF 5 bosqich (2023) – `Namuna_1.3.4-rasm_yangi.docx`, `tools/namuna_134.py`; muallif ko‘rib chiqmoqda. Qo‘shilsa adabiyotga PRA (2023) qo‘shiladi (BCBS WP4 allaqachon bor).
- 1.1.5-rasm yangi varianti faqat namuna faylda: `Namuna_1.1.5-rasm_yangi.docx` (2021–2025, G-7, 12 bank). Muallif ko‘rib chiqib ruxsat bermaguncha asosiy faylga QO‘SHILMAYDI; qo‘shish skripti tayyor: `tools/qosh_115.py`.
- QOIDA: muallif “alohida faylda ber” desa – faqat namuna fayl; asosiy faylga faqat aniq ruxsatdan keyin.
- **Kutilayotgan (muallifga eslatish kerak):** “ekspert vaznlari” → “tayanch vaznlar” (o‘z vaznlarimizga oid ~33 joy, adabiyotdagi “ekspert usuli” o‘zgarmaydi); 3.2-§ boshidagi va cheklovlar jumlasida vaznlar “CAMEL nisbati asosida moslashtirilgan” deb yozish; Kirish usullariga “analogiya (CAMEL vaznlariga tayanish) va statistik verifikatsiya”; 3.1-§ da daromadlilik vazni 0,25→0,15 kamaytirilganining asosini (“Uchinchidan…”) qo‘shish. Sabab: “ekspert” so‘zi so‘rovnoma talabini tug‘diradi. Muallif “keyin eslatasan” dedi.

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
