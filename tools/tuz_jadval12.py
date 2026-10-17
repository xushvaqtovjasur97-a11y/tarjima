"""Barcha jadvallar matnini 12 pt qilish; ustun sarlavhalari (birinchi qator va takrorlanuvchi sarlavha qatorlari) markazda."""
import sys
sys.path.insert(0, '/home/user/tarjima/tools')
import docx
from docx.oxml.ns import qn
from lxml import etree
from pardoz import pardoz

F = '/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx'
d = docx.Document(F)


def set_sz(rpr):
    rf = rpr.find(qn('w:rFonts'))
    if rf is None:
        rf = etree.SubElement(rpr, qn('w:rFonts'))
    for a in list(rf.attrib):
        if a.endswith('Theme'):
            del rf.attrib[a]
    for a in ('ascii', 'hAnsi', 'cs', 'eastAsia'):
        rf.set(qn('w:' + a), 'Times New Roman')
    for tag in ('w:sz', 'w:szCs'):
        e = rpr.find(qn(tag))
        if e is None:
            e = etree.SubElement(rpr, qn(tag))
        e.set(qn('w:val'), '24')
    # rPr elementlari tartibi: sz/szCs ni to‘g‘ri joyga ko‘chirish (Word sxemasi)
    order = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow', 'emboss',
             'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs',
             'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout',
             'specVanish', 'oMath']
    kids = list(rpr)
    kids.sort(key=lambda e: order.index(e.tag.split('}')[1]) if e.tag.split('}')[1] in order else 99)
    for e in kids:
        rpr.remove(e)
    for e in kids:
        rpr.append(e)


def ensure_rpr(parent, before_tags=()):
    rpr = parent.find(qn('w:rPr'))
    if rpr is None:
        rpr = etree.Element(qn('w:rPr'))
        parent.insert(0, rpr) if parent.tag == qn('w:r') else parent.append(rpr)
    return rpr


nt = nr = nh = 0
import collections
nb = collections.Counter()
for tbl in d.element.body.iter(qn('w:tbl')):
    nt += 1
    rows = tbl.findall(qn('w:tr'))
    for r in tbl.iter(qn('w:r')):
        if r.find(qn('w:drawing')) is not None:
            continue
        set_sz(ensure_rpr(r)); nr += 1
    for p in tbl.iter(qn('w:p')):
        ppr = p.find(qn('w:pPr'))
        if ppr is None:
            ppr = etree.Element(qn('w:pPr')); p.insert(0, ppr)
        rpr = ppr.find(qn('w:rPr'))
        if rpr is None:
            rpr = etree.SubElement(ppr, qn('w:rPr'))
        set_sz(rpr)
    # sarlavha qatorlari: birinchi qator + tblHeader belgisi qo‘yilgan qatorlar
    hdr = [rows[0]] if rows else []
    for tr in rows[1:]:
        trpr = tr.find(qn('w:trPr'))
        if trpr is not None and trpr.find(qn('w:tblHeader')) is not None:
            hdr.append(tr)
    for tr in hdr:
        for tc in tr.findall(qn('w:tc')):
            tcpr = tc.find(qn('w:tcPr'))
            if tcpr is None:
                tcpr = etree.Element(qn('w:tcPr')); tc.insert(0, tcpr)
            va = tcpr.find(qn('w:vAlign'))
            if va is None:
                va = etree.SubElement(tcpr, qn('w:vAlign'))
            va.set(qn('w:val'), 'center')
            for p in tc.findall(qn('w:p')):
                ppr = p.find(qn('w:pPr'))
                jc = ppr.find(qn('w:jc'))
                if jc is None:
                    jc = etree.Element(qn('w:jc'))
                    # jc pPr ichida rPr dan oldin turishi kerak
                    rp = ppr.find(qn('w:rPr'))
                    (rp.addprevious(jc) if rp is not None else ppr.append(jc))
                jc.set(qn('w:val'), 'center')
                ind = ppr.find(qn('w:ind'))
                if ind is not None and ind.get(qn('w:firstLine')):
                    ind.set(qn('w:firstLine'), '0')
                nh += 1
    # qolgan kataklar: qisqa yoki raqamli matn – markazda, uzun matn – eni bo‘yicha
    hset = set(id(x) for x in hdr)
    for tr in rows:
        if id(tr) in hset:
            continue
        for ci, tc in enumerate(tr.findall(qn('w:tc'))):
            for p in tc.findall(qn('w:p')):
                txt = ''.join(t.text or '' for t in p.iter(qn('w:t'))).strip()
                if not txt or p.find('.//' + qn('w:drawing')) is not None:
                    continue
                ppr = p.find(qn('w:pPr'))
                num = sum(ch.isdigit() for ch in txt) >= 0.4 * len(txt.replace(' ', ''))
                short = num or (ci > 0 and len(txt) <= 30)
                if ci == 0 and not num and len(tr.findall(qn('w:tc'))) > 1:
                    short = False   # birinchi ustun (ko‘rsatkich nomi) – eni bo‘yicha
                jc = ppr.find(qn('w:jc'))
                if jc is None:
                    jc = etree.Element(qn('w:jc'))
                    rp = ppr.find(qn('w:rPr'))
                    (rp.addprevious(jc) if rp is not None else ppr.append(jc))
                jc.set(qn('w:val'), 'center' if short else 'both')
                ind = ppr.find(qn('w:ind'))
                if ind is not None:
                    for a in ('firstLine', 'hanging'):
                        if ind.get(qn('w:' + a)):
                            ind.set(qn('w:' + a), '0')
                nb[('c' if short else 'j')] += 1
d.save(F)
print('tana kataklari:', dict(nb))
print('jadvallar:', nt, 'run:', nr, 'sarlavha paragraflari:', nh, pardoz(F))
