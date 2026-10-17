"""LibreOffice (UNO) orqali har bir jadval va rasmning boshlanish/tugash sahifasini aniqlash.
Foydalanish: python3 bloklar_uno.py <docx> <json>
Natija: [{kind, idx, first_text, p0, p1, h_mm, cap}] – jadvallar uchun birinchi va oxirgi katak sahifasi,
rasmlar (sarlavha “-rasm.”) uchun rasm langari va sarlavha sahifasi."""
import sys, json, time, subprocess, os, re
import uno
from com.sun.star.beans import PropertyValue

DOCX, OUT = os.path.abspath(sys.argv[1]), sys.argv[2]
port = 2002 + os.getpid() % 500
prof = f'/tmp/lo_uno_{port}'
proc = subprocess.Popen(['soffice', '--headless', '--invisible', '--norestore', f'-env:UserInstallation=file://{prof}',
                         f'--accept=socket,host=127.0.0.1,port={port};urp;'], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
local = uno.getComponentContext()
res = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
for _ in range(120):
    try:
        ctx = res.resolve(f'uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext'); break
    except Exception:
        time.sleep(1)
smgr = ctx.ServiceManager
desk = smgr.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
def pv(n, v):
    p = PropertyValue(); p.Name = n; p.Value = v; return p
doc = desk.loadComponentFromURL(uno.systemPathToFileUrl(DOCX), '_blank', 0, (pv('Hidden', False),))
ctrl = doc.getCurrentController()
vc = ctrl.getViewCursor()
def page_of(rng):
    vc.gotoRange(rng, False); return vc.getPage()
def pos(rng):
    vc.gotoRange(rng, False); return vc.getPage(), vc.getPosition().Y
out = []
tables = doc.getTextTables()
for i in range(tables.getCount()):
    t = tables.getByIndex(i)
    names = t.getCellNames()
    c0 = t.getCellByName(names[0]); cl = t.getCellByName(names[-1])
    try:
        p0, y0 = pos(c0.getStart()); p1, y1 = pos(cl.getEnd())
    except Exception as e:
        continue
    # jadvaldan oldingi ikki xatboshi (sarlavha)
    anc = t.getAnchor()
    out.append(dict(kind='tbl', name=t.getName(), first=c0.getString()[:40], p0=p0, p1=p1, y0=y0, y1=y1, rows=t.getRows().getCount()))
# rasmlar: sarlavha xatboshisi va undan oldingi rasm
enum = doc.getText().createEnumeration()
prevp = None
while enum.hasMoreElements():
    el = enum.nextElement()
    if el.supportsService('com.sun.star.text.Paragraph'):
        s = el.getString()
        m = re.match(r'^(\d\.\d\.\d+)-rasm\.', s)
        if m:
            pc = page_of(el.getStart())
            out.append(dict(kind='cap', num=m.group(1), p=pc, pe=page_of(el.getEnd()), prev=prevp))
        prevp = (page_of(el.getStart()), s[:30])
        m = re.match(r'^(\d\.\d\.\d+)-jadval$', s.strip())
        if m:
            out.append(dict(kind='jhead', num=m.group(1), p=page_of(el.getStart())))
    elif el.supportsService('com.sun.star.text.TextTable'):
        out.append(dict(kind='tblpos', name=el.getName()))
        prevp = ('tbl', el.getName())
# rasm obyektlari langari
for coll in (doc.getGraphicObjects(), doc.getEmbeddedObjects(), doc.getTextFrames()):
    for i in range(coll.getCount()):
        o = coll.getByIndex(i)
        try:
            a = o.getAnchor(); out.append(dict(kind='obj', name=o.getName(), p=page_of(a.getStart()), y=o.getPropertyValue('VertOrientPosition')))
        except Exception:
            pass
out.append(dict(kind='pages', n=ctrl.getPropertyValue('PageCount')))
json.dump(out, open(OUT, 'w'), ensure_ascii=False, indent=0)
doc.close(True)
try:
    desk.terminate()
except Exception:
    pass
proc.wait(timeout=30)
print('ok', len(out))
