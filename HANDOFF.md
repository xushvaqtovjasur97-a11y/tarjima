# Dissertatsiya ishi: holat va davom ettirish yo‘riqnomasi

**Mavzu:** Tijorat banklari moliyaviy barqarorligini xalqaro standartlar asosida baholash metodologiyasini takomillashtirish (DSc).

**Tarmoq:** `claude/awesome-archimedes-fqf08d`

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
