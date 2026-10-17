"""DOCX → PDF LibreOffice (UNO) orqali, Word sahifalashiga yaqinlashtirish sozlamalari bilan.
Foydalanish: python3 render_uno.py <docx> <pdf> [--fnline MM] [--fntext MM] [--info]
Word izoh ajratgichi – oddiy xatboshi (11 pt, 1 interval ≈ 12,65 pt); LibreOffice’da u sahifa uslubidagi
FootnoteLineDistance/FootnoteLineTextDistance bilan beriladi."""
import sys, time, subprocess, os
import uno
from com.sun.star.beans import PropertyValue
DOCX, PDF = os.path.abspath(sys.argv[1]), os.path.abspath(sys.argv[2])
def arg(n, d):
    return float(sys.argv[sys.argv.index(n) + 1]) if n in sys.argv else d
FNLINE, FNTEXT = arg('--fnline', None), arg('--fntext', None)
port = 2600 + os.getpid() % 500
prof = f'/tmp/lo_r_{port}'
proc = subprocess.Popen(['soffice', '--headless', '--invisible', '--norestore', f'-env:UserInstallation=file://{prof}',
                         f'--accept=socket,host=127.0.0.1,port={port};urp;'], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
local = uno.getComponentContext()
res = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
for _ in range(120):
    try:
        ctx = res.resolve(f'uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext'); break
    except Exception:
        time.sleep(1)
desk = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
def pv(n, v):
    p = PropertyValue(); p.Name = n; p.Value = v; return p
doc = desk.loadComponentFromURL(uno.systemPathToFileUrl(DOCX), '_blank', 0, (pv('Hidden', True),))
ps = doc.getStyleFamilies().getByName('PageStyles')
for i in range(ps.getCount()):
    s = ps.getByIndex(i)
    if '--info' in sys.argv and s.isInUse():
        print(s.getName(), 'LineDist', s.FootnoteLineDistance, 'TextDist', s.FootnoteLineTextDistance, 'Height', s.FootnoteHeight, 'Weight', s.FootnoteLineWeight)
    if FNLINE is not None: s.FootnoteLineDistance = int(FNLINE * 100)
    if FNTEXT is not None: s.FootnoteLineTextDistance = int(FNTEXT * 100)
doc.storeToURL(uno.systemPathToFileUrl(PDF), (pv('FilterName', 'writer_pdf_Export'),))
doc.close(True)
try: desk.terminate()
except Exception: pass
proc.wait(timeout=30)
print('ok')
