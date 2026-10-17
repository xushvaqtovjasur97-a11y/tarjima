"""Yakuniy pardoz: «…» → “…” (o‘zbekcha qo‘shtirnoq) va barcha matn shrifti – Times New Roman.

docx fayl zip darajasida qayta yoziladi: asosiy matn, izohlar, kolontitullar, uslublar,
diagrammalar va ularning ichki Excel jadvallari. Ranglarga tegilmaydi.
"""
import io
import re
import sys
import zipfile

TNR = 'Times New Roman'
QUOTES = {'«': '“', '»': '”'}
# belgili shriftlar (ro‘yxat belgilari) o‘zgartirilmaydi
SYMBOL = {'Symbol', 'Wingdings', 'Wingdings 2', 'Wingdings 3', 'Courier New'}
THEME_ATTR = re.compile(r'\s+w:(asciiTheme|hAnsiTheme|eastAsiaTheme|cstheme)="[^"]*"')


def _quotes(s):
    for a, b in QUOTES.items():
        s = s.replace(a, b)
    return s


def _rfonts(m):
    tag = m.group(0)
    explicit = dict(re.findall(r'w:(ascii|hAnsi|eastAsia|cs)="([^"]*)"', tag))
    if any(v in SYMBOL for v in explicit.values()):
        return tag
    rest = re.sub(r'\s+w:(ascii|hAnsi|eastAsia|cs|asciiTheme|hAnsiTheme|eastAsiaTheme|cstheme)="[^"]*"', '', tag)
    rest = rest[len('<w:rFonts'):]
    return f'<w:rFonts w:ascii="{TNR}" w:hAnsi="{TNR}" w:eastAsia="{TNR}" w:cs="{TNR}"' + rest


def _word_xml(name, s):
    stats = {}
    stats['q'] = s.count('«') + s.count('»')
    s = _quotes(s)
    if name.startswith('word/charts/'):
        n = len(re.findall(r'typeface="(?!%s)[^"+][^"]*"' % TNR, s))
        s = re.sub(r'(<a:(?:latin|cs|ea) typeface=")(?!\+)[^"]*(")', r'\g<1>%s\g<2>' % TNR, s)
        stats['f'] = n
    else:
        before = s
        s = re.sub(r'<w:rFonts\b[^>]*/>', _rfonts, s)
        stats['f'] = sum(1 for a, b in zip(re.findall(r'<w:rFonts\b[^>]*/>', before), re.findall(r'<w:rFonts\b[^>]*/>', s)) if a != b)
    if name == 'word/styles.xml':
        # 3-sarlavha uslubidagi ko‘k (accent1) rang – qora
        s = s.replace('<w:color w:val="4F81BD" w:themeColor="accent1"/>', '<w:color w:val="000000"/>')
    return s, stats


def _theme(s):
    return re.sub(r'(<a:(?:majorFont|minorFont)><a:latin typeface=")[^"]*(")', r'\g<1>%s\g<2>' % TNR, s)


def _xlsx(blob):
    zin = zipfile.ZipFile(io.BytesIO(blob))
    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zo:
        for it in zin.infolist():
            data = zin.read(it.filename)
            if it.filename.endswith('.xml'):
                data = _quotes(data.decode('utf-8')).encode('utf-8')
            zo.writestr(it, data)
    return out.getvalue()


def pardoz(path, dst=None):
    dst = dst or path
    zin = zipfile.ZipFile(path)
    items = [(it, zin.read(it.filename)) for it in zin.infolist()]
    zin.close()
    tot = {'q': 0, 'f': 0, 'xlsx': 0}
    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zo:
        for it, data in items:
            n = it.filename
            if n.startswith('word/theme/') and n.endswith('.xml'):
                data = _theme(data.decode('utf-8')).encode('utf-8')
            elif n.startswith('word/') and n.endswith('.xml'):
                s, st = _word_xml(n, data.decode('utf-8'))
                data = s.encode('utf-8')
                tot['q'] += st['q']; tot['f'] += st['f']
            elif n.startswith('word/embeddings/') and n.endswith('.xlsx'):
                data = _xlsx(data); tot['xlsx'] += 1
            zo.writestr(it, data)
    open(dst, 'wb').write(out.getvalue())
    return tot


if __name__ == '__main__':
    for p in sys.argv[1:]:
        print(p, pardoz(p))
