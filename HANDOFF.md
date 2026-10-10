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
| `paragraflar/1.2-paragraf.docx` | 1.2-§, 18 bet (havolalar muallif-yil ko‘rinishida) |
| `paragraflar/1.3-paragraf.docx` | 1.3-§, 18 bet, 7 ta rasm, sahifa osti snoskalari |
| `paragraflar/2.1-paragraf.docx` | 2.1-§, 13 bet, 5 ta rasm, sahifa osti snoskalari |
| `paragraflar/2.2-paragraf.docx` | 2.2-§, 13 bet, 6 ta diagramma; ma’lumotlar qidiruv orqali (IMF FSAP, MB sharhlari), Drive/cbu.uz ma’lumotlari bilan tekshirilishi kerak |
| `paragraflar/Ozbekiston_Bazel_evolyutsiyasi_tahlil.docx` | O‘zbekistonda Bazel talablari evolyutsiyasi (5 bosqich, K1 dinamikasi) |
| `paragraflar/Ozbekiston_Bazel_bosqichlari.docx` | Bazel bosqichlari xronologiyasi |
| `paragraflar/1.3-xorijiy-tajriba.docx` | **Yangi reja** bo‘yicha 1.3-§ (xorijiy tajriba), 24 bet, 7 ta rasm, 3 ta jadval, 67 snoska; 2024–2026-yillardagi islohotlar qo‘shilgan, 1.3.5-rasm 2015–2026-yillar chiziqli grafigi; 1.3.7-rasm NFRA 2023–2025, XVJ 2025 Article IV va Charoenwong va boshq. (2025) ma’lumotlari bilan yangilandi. Manba matni `1.3-xorijiy-tajriba.src.md` (`tools/fn13x.py` snoskalarni to‘liq manba bilan yoyadi, `tools/build13x.py` docx yig‘adi) |
| `paragraflar/rasmlar/1.3.6-rasm_SP_2024.png` | S&P 1981–2024 defolt darajalari |

## Keyingi vazifa

**2.2-§ birinchi varianti yozildi** – foydalanuvchining Drive’idagi va cbu.uz’dagi asl ma’lumotlar bilan solishtirib, bo‘sh kataklarni (2021, 2024 yillar) to‘ldirish kerak. Keyingisi: 2.3-§.

Kerakli ma’lumotlar: XVJ FSI mezonlari bo‘yicha 2021–2025-yillar (kapital, CET1, NPL, ROA, ROE, LCR, NSFR, likvid aktivlar, valyuta kreditlari va depozitlari ulushi). Manbalar: Markaziy bankning moliyaviy barqarorlik sharhlari (cbu.uz) va foydalanuvchining Google Drive’idagi fayllar.

Oldingi sessiyada `cbu.uz`, `lex.uz` tarmoqdan bloklangan, Google Drive ulagichi esa suhbatda yoqilmagan edi. Foydalanuvchi sozlamani o‘zgartirdi – yangi sessiyada ishlashi kerak. Ishlamasa, foydalanuvchidan fayllarni biriktirishni so‘rang.

Qidiruvda topilgan, lekin hali tekshirilmagan raqamlar: kapital monandligi 17,5% (2023), 17,0% (2024), 17,4% / CET1 14,6% (2025 I yarim), 18,3% / CET1 14,7% (2025), 18,5% / CET1 15,0% / NPL 3,6% (01.07.2026); 2025 I yarim: NPL 3,8%, ROE 10,8%, ROA 2%, LCR 195%, NSFR 117%, yuqori likvid aktivlar 18%; valyuta kreditlari 47% (2021) → 42% (2024-iyun), depozit dollarlashuvi 40% → 30%.

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
