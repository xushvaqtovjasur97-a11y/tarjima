# NSFR va LCR qayta hisob-kitobini tekshirish (2025-yil oxiri): SQB, MKB, Aloqabank.
# Savol: barqaror moliyalashtirishda xususiy kapital o‘rniga regulyativ kapital olinganda kapital instrumentlari
# (muddatsiz obligatsiyalar – AT1, subordinar qarz – T2) ikki marta hisoblanmaydimi?
# Shakl: 16-jadval (Likvidlilik tahlili) – 2.5.1 xususiy kapital, 2.5.2 bir yildan ortiq majburiyatlar (100 %),
#        2.5.3 muddatsiz majburiyatlar (30 %), 2.5.4 bir yilgacha majburiyatlar (30 %); 18-jadval – majburiyatlar muddatlari;
#        5-jadval – regulyativ kapital.
import json, re, openpyxl

B = '/home/user/tarjima/manbalar/bank_hisobotlari/'
FILES = {'SQB': (B + 'Ўзсаноатқурилишбанк-31.12.2020-31.12.2025 — копия/2025/crs003L.XLSX', B + 'Ўзсаноатқурилишбанк-31.12.2020-31.12.2025 — копия/2025/crs003C.XLSX'),
         'MKB': (B + 'Микрокредитбанк-31.12.2016-31.12.2025 — копия/2025/crs005L.XLSX', B + 'Микрокредитбанк-31.12.2016-31.12.2025 — копия/2025/crs005C.XLSX'),
         'ALOQA': (B + 'Aloqabank-31.12.2020-31.12.2025/2025/crs031L.XLSX', B + 'Aloqabank-31.12.2020-31.12.2025/2025/crs031C.xlsx')}


def rows(ws):
    for r in ws.iter_rows(values_only=True):
        yield [x for x in r]


def num(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return None


def find(ws, code=None, text=None):
    for r in rows(ws):
        cells = [str(x).strip() if x is not None else '' for x in r]
        if code is not None and cells and cells[0] == code:
            return r
        if text is not None and any(text in c for c in cells):
            return r
    return None


def first_num(r, skip=1):
    vals = [num(x) for x in r[skip:]]
    vals = [v for v in vals if v is not None]
    return vals[0] if vals else None


def sheet(wb, *names):
    for n in names:
        for s in wb.sheetnames:
            if s.strip().startswith(n):
                return wb[s]
    raise KeyError(names)


OUT = {}
for bank, (fl, fc) in FILES.items():
    wl = openpyxl.load_workbook(fl, data_only=True)
    wc = openpyxl.load_workbook(fc, data_only=True)
    la = sheet(wl, 'Liquidity Analysis')
    g = lambda code: first_num(find(la, code=code), 2) / 1e6           # mlrd so‘m
    nsfr = first_num(find(la, code='2.4'), 2)
    asf = {k: g(k) for k in ('2.5', '2.5.1', '2.5.2', '2.5.3', '2.5.4')}
    rsf = g('2.6')
    hqla = g('2.3.1'); out30 = g('2.3.2')
    # regulyativ kapital (5-jadval)
    cc = sheet(wc, 'Bank Capital Calc')
    cap = {}
    for r in rows(cc):
        c0 = str(r[0]).strip() if r[0] is not None else ''
        txt = ' '.join(str(x) for x in r[1:3] if x is not None)
        v = first_num(r, 2)
        if v is None:
            continue
        if c0 == 'I' and 'АСОСИЙ' in txt: cap['CET1'] = v / 1e6
        elif c0 == '1.6.': cap['CET1_gross'] = v / 1e6
        elif c0 == 'II' and 'ҚЎШИМЧА' in txt: cap['AT1'] = v / 1e6
        elif c0 == '2.1.': cap['AT1_instr'] = v / 1e6
        elif c0 == '3.1.': cap['T2_profit'] = v / 1e6
        elif c0 == '3.2.': cap['T2_reserves'] = v / 1e6
        elif c0 == '3.4.': cap['T2_subdebt'] = v / 1e6
        elif c0 == '3.6.': cap['T2_cap'] = v / 1e6
        elif c0 == 'IV': cap['T2'] = v / 1e6
        elif c0 == 'V': cap['REG'] = v / 1e6
    # majburiyatlar muddatlari (18-jadval): chiqarilgan qimmatli qog‘ozlar va subordinar qarz
    lb = sheet(wl, 'Bank Liabilities Analysis')
    def maturity(code):
        r = find(lb, code=code)
        if r is None: return None
        vals = [num(x) or 0 for x in r[2:19]]
        # C muddatsiz, E 1–7, G 8–30, I 31–90, K 91–180, M 181–365, O 366–730, Q >2 yil (sum ustunlari juft o‘rinlarda)
        s = vals[0::2]
        return {'muddatsiz': s[0] / 1e6, '≤1 yil': sum(s[1:6]) / 1e6, '>1 yil': sum(s[6:8]) / 1e6}
    sec, sub = maturity('6.'), maturity('7.')
    own = asf['2.5.1']; reg = cap['REG']
    naive = (asf['2.5'] - own + reg) / rsf * 100
    # kapitalga kiritilgan instrumentlar 2.5.2 (100 %) yoki 2.5.3 (30 %) da allaqachon bor – bir marta hisoblash
    at1_in = cap.get('AT1_instr') or 0
    t2_sub = cap.get('T2_subdebt') or 0
    dbl = 0.0
    for amt, mat in ((at1_in, sec), (t2_sub, sub)):
        if not amt or not mat: continue
        tot = sum(mat.values()) or 1
        dbl += amt * (mat['>1 yil'] / tot * 1.0 + mat['muddatsiz'] / tot * 0.3 + mat['≤1 yil'] / tot * 0.3)
    corr_after = (asf['2.5'] - own + reg - dbl) / rsf * 100
    reg_gross = reg + (cap.get('CET1_gross', 0) - cap.get('CET1', 0)) - (cap.get('T2_cap') or 0)
    corr_before = (asf['2.5'] - own + reg_gross - dbl) / rsf * 100
    OUT[bank] = dict(NSFR_shakl=round(nsfr * 100, 1), ASF=round(asf['2.5'], 1), RSF=round(rsf, 1), xususiy=round(own, 1), regulyativ=round(reg, 1),
                     farq=round(reg - own, 1), AT1_instr=round(at1_in, 1), T2_subdebt=round(t2_sub, 1), T2_zaxira=round(cap.get('T2_reserves') or 0, 1),
                     chiqarilgan_qq=sec and {k: round(v, 1) for k, v in sec.items()}, subordinar=sub and {k: round(v, 1) for k, v in sub.items()},
                     ikki_marta=round(dbl, 1), NSFR_oddiy_almashtirish=round(naive, 1), NSFR_tuzatilgan=round(corr_after, 1),
                     NSFR_tuzatilgan_chegirmagacha=round(corr_before, 1), LCR=round(hqla / out30 * 100, 1), HQLA=round(hqla, 1), sof_chiqim=round(out30, 1))
print(json.dumps(OUT, ensure_ascii=False, indent=1))
json.dump(OUT, open('/home/user/tarjima/manbalar/hisob_kitob/nsfr_tekshiruv.json', 'w'), ensure_ascii=False, indent=1)
