# Banklar MHXS ko'rsatkichlari fayli sxemasi (<kod>.json)

{
 "bank": "To'liq nomi", "kod": "NBU", "mulk": "davlat|xususiy|xorijiy",
 "manbalar": {"2021": {"url": "...", "hujjat": "to'liq nomi (rus/ingliz)", "auditor": "..."}, ...},
 "yillar": {
   "2020": {...faqat TA, EQ...},
   "2021": {
     "TA": jami aktivlar, "EQ": jami kapital, "NCI": ozchilik ulushi, "NP": yil sof foydasi (zarar manfiy),
     "LOANS_GROSS": mijozlarga kreditlar brutto (amortizatsiyalangan qiymat + FVTPL), "ECL": kreditlar bo'yicha zaxira (musbat son),
     "S1","S2","S3": bosqichlar bo'yicha brutto kreditlar (amortizatsiyalangan qiymat), "S_TOTAL": ularning yig'indisi,
     "DEP": mijozlar mablag'lari (customer accounts),
     "CAR","T1","CET1": hisobotda oshkor qilingan kapital monandligi, % (qaysi asosda – milliy yoki Bazel – "CAP_BASIS" da yozing),
     "RWA": risk bo'yicha tortilgan aktivlar (agar oshkor qilingan), "REG_CAP": regulyativ kapital, "T1_CAP": I darajali kapital,
     "LEV": leverage, %, "LCR": %, "NSFR": % (agar oshkor qilingan bo'lsa),
     "NII": zaxiradan oldingi sof foizli daromad, "FEE_NET": sof komission daromad, "FX_NET": valyuta operatsiyalaridan sof daromad (savdo + qayta baholash),
     "OTHER_INC": boshqa operatsion daromadlar, "OPEX": ma'muriy va boshqa operatsion xarajatlar (xodimlar + boshqa),
     "FX_LOANS_SHARE": xorijiy valyutadagi kreditlar ulushi, % (valyuta riski izohidan), "FX_LIAB_SHARE": xorijiy valyutadagi majburiyatlar ulushi, %,
     "betlar": {"TA": "2", "S3": "45", ...}  // hisobotning o'z sahifa raqamlari
   }, ...
 },
 "izohlar": "qayta hisoblash (restated), birlik (mln/ming so'm), topilmagan ko'rsatkichlar va sabablari"
}
Barcha summalar mlrd so'mda (1 kasr). Har yil uchun o'sha yilning o'z hisobotidagi qiymat olinadi; agar yil hisoboti topilmasa, keyingi yil hisobotidagi qiyosiy (qayta hisoblangan) qiymat olinadi va izohda yoziladi.
