# Kamchiliklar jadvali bo'yicha tuzatishlarni tayyor docx faylga qizil rangda kiritish vositalari
import copy, re
from lxml import etree
import docx
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
RED = 'FF0000'
LOG = []


class Doc:
    def __init__(self, path):
        self.d = docx.Document(path)
        self.fpart = None
        for r in self.d.part.rels.values():
            if r.reltype.endswith('/footnotes'):
                self.fpart = r.target_part
        if self.fpart is not None:
            self.fx = etree.fromstring(self.fpart.blob)
            ids = [int(f.get(qn('w:id'))) for f in self.fx.findall(qn('w:footnote'))]
            self.next_fn = max(ids) + 1
        else:
            self.fx = None
        self.charts = {r.target_part.partname: r.target_part for r in self.d.part.rels.values() if r.reltype.endswith('/chart')}

    # ---------- qidirish
    def paras(self):
        out = []
        for el in self.d.element.body.iterchildren():
            if el.tag == qn('w:p'):
                out.append(Paragraph(el, self.d))
            elif el.tag == qn('w:tbl'):
                for row in Table(el, self.d).rows:
                    for c in row.cells:
                        out.extend(c.paragraphs)
        return out

    def para(self, sub, after=None, nth=0):
        ps = [p for p in self.paras() if sub in p.text]
        if after:
            idx = [i for i, p in enumerate(self.paras()) if after in p.text]
            allp = self.paras()
            start = idx[0]
            ps = [p for p in allp[start:] if sub in p.text]
        if not ps:
            raise KeyError('topilmadi: ' + sub[:80])
        if len(ps) > 1 and nth is None:
            raise KeyError(f'{len(ps)} ta: ' + sub[:80])
        return ps[nth]

    def table(self, header_sub, nth=0):
        ts = [Table(el, self.d) for el in self.d.element.body.iterchildren(qn('w:tbl'))]
        ts = [t for t in ts if header_sub in ' | '.join(c.text for c in t.rows[0].cells)]
        if not ts:
            raise KeyError('jadval topilmadi: ' + header_sub)
        return ts[nth]

    def table_after(self, caption_sub):
        found = False
        for el in self.d.element.body.iterchildren():
            if el.tag == qn('w:p') and caption_sub in Paragraph(el, self.d).text:
                found = True
            elif found and el.tag == qn('w:tbl'):
                return Table(el, self.d)
        raise KeyError('jadval topilmadi: ' + caption_sub)

    # ---------- run darajasidagi amallar
    @staticmethod
    def _runs(p):
        return [r for r in p._p.iter(qn('w:r')) if r.find(qn('w:t')) is not None]

    @staticmethod
    def _red_run(tmpl, text):
        r = etree.SubElement(etree.Element('dummy'), qn('w:r'))
        rpr = tmpl.find(qn('w:rPr')) if tmpl is not None else None
        if rpr is not None:
            rpr = copy.deepcopy(rpr)
            for tag in ('w:color', 'w:highlight', 'w:shd', 'w:vertAlign', 'w:rStyle'):
                for e in rpr.findall(qn(tag)):
                    rpr.remove(e)
        else:
            rpr = etree.Element(qn('w:rPr'))
        c = etree.Element(qn('w:color'))
        c.set(qn('w:val'), RED)
        # CT_RPr ketma-ketligi: w:color 'w:vanish/webHidden' dan keyin, 'w:spacing ... ' dan oldin turishi shart
        after = {qn('w:' + x) for x in ('spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd',
                                        'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath')}
        nxt = [e for e in rpr if e.tag in after]
        if nxt:
            nxt[0].addprevious(c)
        else:
            rpr.append(c)
        r.append(rpr)
        t = etree.SubElement(r, qn('w:t'))
        t.text = text
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        return r

    def _fn_ref_run(self, text):
        fid = self.next_fn
        self.next_fn += 1
        proto = None
        for f in self.fx.findall(qn('w:footnote')):
            if int(f.get(qn('w:id'))) > 0:
                proto = f
                break
        f = copy.deepcopy(proto)
        f.set(qn('w:id'), str(fid))
        p = f.find(qn('w:p'))
        for e in list(p):
            if e.tag != qn('w:pPr'):
                p.remove(e)
        r1 = etree.SubElement(p, qn('w:r'))
        rp = etree.SubElement(r1, qn('w:rPr'))
        etree.SubElement(rp, qn('w:rStyle')).set(qn('w:val'), 'aff8')
        etree.SubElement(r1, qn('w:footnoteRef'))
        r2 = etree.SubElement(p, qn('w:r'))
        rp2 = etree.SubElement(r2, qn('w:rPr'))
        etree.SubElement(rp2, qn('w:color')).set(qn('w:val'), RED)
        t = etree.SubElement(r2, qn('w:t'))
        t.text = ' ' + text
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        self.fx.append(f)
        r = etree.Element(qn('w:r'))
        rp = etree.SubElement(r, qn('w:rPr'))
        etree.SubElement(rp, qn('w:rStyle')).set(qn('w:val'), 'aff8')
        etree.SubElement(rp, qn('w:color')).set(qn('w:val'), RED)
        etree.SubElement(r, qn('w:footnoteReference')).set(qn('w:id'), str(fid))
        return r

    def _build(self, tmpl, text):
        """Matnni qizil run'larga aylantiradi; [[FN:...]] – yangi izoh."""
        out = []
        for part in re.split(r'(\[\[FN:.*?\]\])', text):
            if not part:
                continue
            if part.startswith('[[FN:'):
                out.append(self._fn_ref_run(part[5:-2]))
            else:
                out.append(self._red_run(tmpl, part))
        return out

    def _split_at(self, p, pos):
        """Paragraf matnining pos indeksida run'ni ikkiga bo'ladi; pos dan boshlanuvchi run'ni qaytaradi (yoki None – oxiri)."""
        acc = 0
        for r in self._runs(p):
            t = r.find(qn('w:t'))
            s = t.text or ''
            if acc <= pos < acc + len(s):
                k = pos - acc
                if k == 0:
                    return r
                r2 = copy.deepcopy(r)
                t.text = s[:k]
                t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                t2 = r2.find(qn('w:t'))
                t2.text = s[k:]
                t2.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                r.addnext(r2)
                return r2
            acc += len(s)
        return None

    @staticmethod
    def _normalize(p):
        """Bir nechta w:t (yoki w:tab, w:br) saqlagan run'larni har birida bitta element qoladigan run'larga ajratadi."""
        for r in list(p._p.iter(qn('w:r'))):
            cont = [e for e in r if e.tag != qn('w:rPr')]
            if len(cont) <= 1:
                continue
            rpr = r.find(qn('w:rPr'))
            prev = r
            for e in cont[1:]:
                nr = etree.Element(qn('w:r'))
                if rpr is not None:
                    nr.append(copy.deepcopy(rpr))
                nr.append(e)
                prev.addnext(nr)
                prev = nr

    def _text_index(self, p, sub, occ=0):
        self._normalize(p)
        txt = ''.join((r.find(qn('w:t')).text or '') for r in self._runs(p))
        i = -1
        for _ in range(occ + 1):
            i = txt.find(sub, i + 1)
            if i < 0:
                raise KeyError(f'matnda yo‘q: {sub[:80]} || {txt[:120]}')
        return i, txt

    def replace(self, p, old, new, occ=0, tag=''):
        """old matnni o'chirib, o'rniga new ni qizil rangda qo'yadi."""
        if isinstance(p, str):
            p = self.para(p)
        i, txt = self._text_index(p, old, occ)
        start = self._split_at(p, i)
        end = self._split_at(p, i + len(old))
        # o'chiriladigan run'lar
        runs = self._runs(p)
        tmpl = start
        dele = []
        on = False
        for r in runs:
            if r is start:
                on = True
            if r is end:
                break
            if on:
                dele.append(r)
        anchor = dele[0]
        tmplc = copy.deepcopy(anchor)
        for e in self._build(tmplc, new):
            anchor.addprevious(e)
        for r in dele:
            r.getparent().remove(r)
        LOG.append((tag, 'replace', old[:60], new[:80]))
        return p

    def insert(self, p, after, new, occ=0, tag=''):
        """after matnidan keyin new ni qizil rangda qo'shadi (after='' – paragraf oxiri)."""
        if isinstance(p, str):
            p = self.para(p)
        if after == '':
            runs = self._runs(p)
            last = runs[-1]
            # izoh havolasidan keyin bo'lsa ham oxiriga qo'shamiz
            lastall = [r for r in p._p if r.tag == qn('w:r')][-1]
            for e in reversed(self._build(last, new)):
                lastall.addnext(e)
        else:
            i, txt = self._text_index(p, after, occ)
            nxt = self._split_at(p, i + len(after))
            tmpl = self._runs(p)[0]
            prev = None
            for r in self._runs(p):
                if r is nxt:
                    break
                prev = r
            # oldingi run va keyingi run orasida izoh havolasi bo'lsa – undan keyin qo'yamiz
            el = prev
            nx = el.getnext()
            while nx is not None and nx.tag == qn('w:r') and nx.find(qn('w:t')) is None and nx.find(qn('w:footnoteReference')) is not None and after.endswith(('.', ')')) is False and False:
                el = nx
                nx = el.getnext()
            for e in reversed(self._build(prev, new)):
                el.addnext(e)
        LOG.append((tag, 'insert', after[-40:], new[:80]))
        return p

    def new_para_after(self, p, text, tag='', style_from=None):
        if isinstance(p, str):
            p = self.para(p)
        np_ = copy.deepcopy(p._p)
        for e in list(np_):
            if e.tag != qn('w:pPr'):
                np_.remove(e)
        src = style_from._p if style_from is not None else p._p
        tmpl = [r for r in src.iter(qn('w:r')) if r.find(qn('w:t')) is not None]
        tmpl = tmpl[-1] if tmpl else None
        for e in self._build(tmpl, text):
            np_.append(e)
        p._p.addnext(np_)
        LOG.append((tag, 'new_para', '', text[:80]))
        return Paragraph(np_, p._parent)

    def delete_para(self, p, tag=''):
        if isinstance(p, str):
            p = self.para(p)
        p._p.getparent().remove(p._p)
        LOG.append((tag, 'delete_para', p.text[:60], ''))

    # ---------- mavjud izohlar
    def fn_para(self, fid=None, contains=None):
        for f in self.fx.findall(qn('w:footnote')):
            t = ''.join(x.text or '' for x in f.iter(qn('w:t')))
            if (fid is not None and f.get(qn('w:id')) == str(fid)) or (contains and contains in t):
                return Paragraph(f.find(qn('w:p')), None)
        raise KeyError(f'izoh topilmadi: {fid} {contains}')

    def fn_ids(self, contains):
        out = []
        for f in self.fx.findall(qn('w:footnote')):
            t = ''.join(x.text or '' for x in f.iter(qn('w:t')))
            if contains in t:
                out.append(f.get(qn('w:id')))
        return out

    # ---------- jadvallar
    def set_cell(self, t, r, c, text, tag=''):
        cell = t.rows[r].cells[c]
        ps = cell.paragraphs
        p = ps[0]
        runs = self._runs(p)
        tmpl = copy.deepcopy(runs[0]) if runs else None
        for pp in ps[1:]:
            pp._p.getparent().remove(pp._p)
        for r_ in list(p._p.iter(qn('w:r'))):
            r_.getparent().remove(r_)
        for e in self._build(tmpl, text):
            p._p.append(e)
        LOG.append((tag, 'cell', f'{r},{c}', text))

    def cell_replace(self, t, r, c, old, new, tag=''):
        p = [pp for pp in t.rows[r].cells[c].paragraphs if old in pp.text][0]
        self.replace(p, old, new, tag=tag)

    def add_row(self, t, after_r, values, tag=''):
        tr = t.rows[after_r]._tr
        ntr = copy.deepcopy(tr)
        tr.addnext(ntr)
        from docx.table import _Row
        row = _Row(ntr, t)
        idx = list(t._tbl.iter(qn('w:tr'))).index(ntr)
        for c, v in enumerate(values):
            self.set_cell(t, idx, c, v, tag=tag)
        return idx

    def add_col(self, t, after_c, values, tag=''):
        """Jadvalga ustun qo'shadi (after_c ustundan nusxa), values – har bir qator uchun."""
        merged = set()
        for ri, tr in enumerate(t._tbl.findall(qn('w:tr'))):
            tcs = tr.findall(qn('w:tc'))
            if len(tcs) <= after_c:   # birlashtirilgan qator: oxirgi katakni kengaytiramiz
                tcpr = tcs[-1].find(qn('w:tcPr'))
                gs = tcpr.find(qn('w:gridSpan')) if tcpr is not None else None
                if gs is not None:
                    gs.set(qn('w:val'), str(int(gs.get(qn('w:val'))) + 1))
                merged.add(ri)
                continue
            ntc = copy.deepcopy(tcs[after_c])
            tcs[after_c].addnext(ntc)
        grid = t._tbl.find(qn('w:tblGrid'))
        if grid is not None:
            cols = grid.findall(qn('w:gridCol'))
            ng = copy.deepcopy(cols[after_c])
            cols[after_c].addnext(ng)
            # kengliklarni qayta taqsimlash
            ws = [int(g.get(qn('w:w'))) for g in grid.findall(qn('w:gridCol'))]
            tot = sum(ws) - int(ng.get(qn('w:w')))
            k = tot / sum(ws)
            for g in grid.findall(qn('w:gridCol')):
                g.set(qn('w:w'), str(int(int(g.get(qn('w:w'))) * k)))
            nws = [int(g.get(qn('w:w'))) for g in grid.findall(qn('w:gridCol'))]
            for ri, tr in enumerate(t._tbl.findall(qn('w:tr'))):
                if ri in merged:
                    continue
                for tc, w in zip(tr.findall(qn('w:tc')), nws):
                    tcw = tc.find(qn('w:tcPr') + '/' + qn('w:tcW'))
                    if tcw is not None:
                        tcw.set(qn('w:w'), str(w))
                        tcw.set(qn('w:type'), 'dxa')
        for ri, v in enumerate(values):
            if ri in merged:
                continue
            self.set_cell(t, ri, after_c + 1, v, tag=tag)

    # ---------- diagrammalar (kesh qiymatlari)
    def chart_xml(self, name):
        part = self.charts[name]
        return etree.fromstring(part.blob), part

    def chart_set(self, name, ser_idx, values, fmt=lambda v: repr(v), tag=''):
        x, part = self.chart_xml(name)
        ns = {'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'}
        sers = x.findall('.//c:ser', ns)
        s = sers[ser_idx]
        pts = s.findall('./c:val//c:pt', ns) or s.findall('./c:yVal//c:pt', ns)
        for pt in pts:
            i = int(pt.get('idx'))
            if i < len(values) and values[i] is not None:
                pt.find('c:v', ns).text = str(values[i])
        part._blob = etree.tostring(x, xml_declaration=True, encoding='UTF-8', standalone=True)
        LOG.append((tag, 'chart', name, str(values)))

    def _reorder_footnotes(self):
        # Izohlarni hujjatdagi havolalar tartibiga keltirish (LibreOffice izohlarni fayl tartibida o'qiydi)
        order = []
        for r in self.d.element.body.iter(qn('w:footnoteReference')):
            i = r.get(qn('w:id'))
            if i not in order:
                order.append(i)
        fns = self.fx.findall(qn('w:footnote'))
        special = [f for f in fns if f.get(qn('w:type')) in ('separator', 'continuationSeparator', 'continuationNotice')]
        by = {f.get(qn('w:id')): f for f in fns if f not in special}
        for f in fns:
            self.fx.remove(f)
        for f in special:
            self.fx.append(f)
        for i in order:
            if i in by:
                self.fx.append(by.pop(i))
        for f in by.values():
            self.fx.append(f)
        # id'larni havolalar tartibida ketma-ket qayta raqamlash (1, 2, 3, ...)
        base = 1
        new = {}
        for f in self.fx.findall(qn('w:footnote')):
            if f in special:
                continue
            new[f.get(qn('w:id'))] = str(base)
            f.set(qn('w:id'), str(base))
            base += 1
        for r in self.d.element.body.iter(qn('w:footnoteReference')):
            r.set(qn('w:id'), new[r.get(qn('w:id'))])

    def save(self, path):
        if self.fx is not None:
            self._reorder_footnotes()
            self.fpart._blob = etree.tostring(self.fx, xml_declaration=True, encoding='UTF-8', standalone=True)
        self.d.save(path)
