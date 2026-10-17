"""3.1.1-rasm: och fon + to‘q matn (yozuvlar har qanday ko‘ruvchida ko‘rinadi), 6-qutichada blok harflari (namuna)."""
import copy, re, sys
sys.path.insert(0, '/home/user/tarjima/tools')
from lxml import etree
from docx.oxml.ns import qn
import namuna_311r as N

A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
# asl fon: (yangi och fon, matn rangi)
PAL = {'1F4E79': ('DEEAF6', '1F4E79'), '2A78D6': ('DDEBF7', '1A4F8F'), 'EB6834': ('FBE5D6', 'A0400F'),
       '70AD47': ('E2F0D9', '375623'), '7030A0': ('E9E1F2', '4B1F6F'), 'C00000': ('F8D9D9', '8B0000')}


def rangla(fig):
    n = 0
    for sp in fig.iter('{http://schemas.microsoft.com/office/word/2010/wordprocessingShape}wsp'):
        f = sp.find('{http://schemas.microsoft.com/office/word/2010/wordprocessingShape}spPr/' + A + 'solidFill/' + A + 'srgbClr')
        if f is None or f.get('val') not in PAL:
            continue
        fon, mat = PAL[f.get('val')]
        f.set('val', fon)
        ln = sp.find('{http://schemas.microsoft.com/office/word/2010/wordprocessingShape}spPr/' + A + 'ln')
        if ln is not None:
            for c in ln.iter(A + 'srgbClr'):
                c.set('val', mat)
        for c in sp.iter(qn('w:color')):
            if c.get(qn('w:val')) == 'FFFFFF':
                c.set(qn('w:val'), mat)
        n += 1
    # VML zaxira (fallback)
    for el in fig.iter():
        fc = el.get('fillcolor')
        if fc and fc.lstrip('#').upper() in PAL:
            fon, mat = PAL[fc.lstrip('#').upper()]
            el.set('fillcolor', '#' + fon)
            if el.get('strokecolor'):
                el.set('strokecolor', '#' + mat)
            for c in el.iter(qn('w:color')):
                if c.get(qn('w:val')) == 'FFFFFF':
                    c.set(qn('w:val'), mat)
    return n


if __name__ == '__main__':
    import docx
    d = docx.Document(N.DST)
    fig = [p for p in d.paragraphs if p._p.find('.//' + qn('w:drawing')) is not None][0]._p
    print('shakllar:', rangla(fig))
    d.save(N.DST)
    from pardoz import pardoz
    print(pardoz(N.DST))
