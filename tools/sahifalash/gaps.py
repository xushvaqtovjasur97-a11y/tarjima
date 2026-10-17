"""PDF sahifalaridagi pastki bo‘sh joylarni o‘lchash (Liberation Serif = TNR metrikasi)."""
import re, sys, subprocess, collections, json
pdf = sys.argv[1]
first = int(sys.argv[2]) if len(sys.argv) > 2 else 1
last = int(sys.argv[3]) if len(sys.argv) > 3 else 9999
html = subprocess.run(['pdftotext', '-bbox', '-f', str(first), '-l', str(last), pdf, '-'], capture_output=True, text=True).stdout
pages = re.split(r'<page width="[\d.]+" height="[\d.]+">', html)[1:]
SEC = re.compile(r'^(KIRISH|XULOSA|FOYDALANILGAN|ILOVALAR|[IV]+ BOB|[1-4]\.[1-3]-§|[IV]+ bob bo‘yicha xulosalar|\d+-ilova)')
res = []
for k, pg in enumerate(pages):
    pno = first + k
    W = [(float(a), float(b), float(c), float(d), t) for a, b, c, d, t in
         re.findall(r'<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">([^<]*)</word>', pg)]
    lines = collections.OrderedDict()
    for w in W:
        lines.setdefault(round(w[3]), []).append(w)
    L = []
    for y, ws in sorted(lines.items()):
        h = max(w[3] - w[1] for w in ws)
        L.append((y, h, min(w[1] for w in ws), ' '.join(w[4] for w in sorted(ws))))
    L = [l for l in L if l[0] < 795]          # sahifa raqami
    # izohlar bloki: pastdan yuqoriga kichik shriftli qatorlar
    fn_top = None
    i = len(L) - 1
    while i >= 0 and L[i][1] <= 11.6:
        fn_top = L[i][2]; i -= 1
    body = L[:i + 1]
    # izohlar bo‘lmasa (pastda grafik yorliqlari bo‘lishi mumkin) – tekshiruv
    if fn_top is not None and body and fn_top - body[-1][0] > 60:
        # pastdagi kichik qatorlar izoh emas, grafik bo‘lishi mumkin: izoh boshlanishi superscript bilan
        pass
    limit = (fn_top - 14) if fn_top is not None else 786
    bottom = body[-1][0] if body else 57
    gap = limit - bottom
    firsttxt = body[0][3] if body else ''
    res.append(dict(p=pno, gap=round(gap), lines=round(gap / 24, 1), first=firsttxt[:60],
                    last=(body[-1][3][:60] if body else ''), sec=bool(SEC.match(firsttxt))))
# keyingi sahifa bo‘lim boshi bo‘lsa – oxirgi sahifa (bo‘sh joy tabiiy)
for a, b in zip(res, res[1:]):
    a['sec_end'] = b['sec']
res[-1]['sec_end'] = True
json.dump(res, open(sys.argv[4] if len(sys.argv) > 4 else '/dev/stdout', 'w'), ensure_ascii=False, indent=0)
