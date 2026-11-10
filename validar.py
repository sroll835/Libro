import glob, zipfile, re, sys
from lxml import etree
sys.argv=['x']
import construir_word as cw
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
def chk(el, orden):
    idx=[orden.index(c.tag.replace(W,'w:')) for c in el if c.tag.replace(W,'w:') in orden]
    return idx==sorted(idx)
for f in sorted(glob.glob('libros/*.docx')):
    z=zipfile.ZipFile(f); x=z.read('word/document.xml'); r=etree.fromstring(x)
    e=sum(not chk(p,cw.ORDEN_PPR) for p in r.iter(W+'pPr'))+sum(not chk(p,cw.ORDEN_RPR) for p in r.iter(W+'rPr'))
    e+=not chk(etree.fromstring(z.read('word/settings.xml')),cw.ORDEN_SETTINGS)
    t=re.sub(r'<[^>]+>',' ',x.decode('utf-8'))
    print(f,'errores:',e,'palabras:',len(t.split()),'**:',t.count('**'))
