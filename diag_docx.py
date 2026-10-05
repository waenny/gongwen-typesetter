import sys
from docx import Document
from docx.oxml.ns import qn


def ea(r):
    rf = r._r.find(qn("w:rPr"))
    if rf is None:
        return None
    e = rf.find(qn("w:rFonts"))
    return e.get(qn("w:eastAsia")) if e is not None else None


def size(r):
    return (r.font.size.pt if r.font.size else None)


for path in sys.argv[1:]:
    print("\n########", path.split("\\")[-1])
    d = Document(path)
    for i, p in enumerate(d.paragraphs):
        if not (p.text or "").strip():
            continue
        pPr = p._p.pPr
        out = "-"
        ind = "-"
        if pPr is not None:
            o = pPr.find(qn("w:outlineLvl"))
            if o is not None:
                out = o.get(qn("w:val"))
            e = pPr.find(qn("w:ind"))
            if e is not None:
                ind = e.get(qn("w:firstLineChars"))
        al = str(p.alignment) if p.alignment is not None else "None"
        fonts = []
        for r in p.runs:
            fonts.append((r.text[:10], ea(r), size(r), r.font.bold))
        print("[%d] out=%s align=%s indChars=%s | %r" % (i, out, al, ind, p.text[:34]))
        for f in fonts:
            print("      run", f)
