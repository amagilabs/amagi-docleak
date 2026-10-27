"""PowerPoint (.pptx/.pptm) checks."""

from .docx import _drawing_fill_dark
from .model import Finding
from .ooxml import q, text_of
from .util import join_names, plural, quoted, secret

SKIP_PH = {"sldImg", "sldNum", "hdr", "ftr", "dt"}
SHAPES = {"p:sp", "p:pic", "p:graphicFrame", "p:grpSp", "p:cxnSp"}


def slides(pkg):
    """[(number, part)] in show order."""
    pres = pkg.main_part()
    rels = {r.id: r for r in pkg.rels(pres)}
    out = []
    lst = pkg.xml(pres).find(q("p:sldIdLst"))
    for i, sid in enumerate(lst if lst is not None else [], 1):
        rel = rels.get(sid.get(q("r:id")))
        if rel is not None and not rel.external and pkg.has(rel.part):
            out.append((i, rel.part))
    return out


def notes_text(root):
    chunks = []
    for sp in root.iter(q("p:sp")):
        ph = sp.find(".//" + q("p:ph"))
        if ph is not None and ph.get("type") in SKIP_PH:
            continue
        t = text_of(sp)
        if t.strip():
            chunks.append(t)
    return " ".join(chunks)


def _xfrm(shape):
    for path in ("p:spPr/a:xfrm", "p:grpSpPr/a:xfrm", "p:xfrm"):
        el = shape
        for step in path.split("/"):
            el = el.find(q(step)) if el is not None else None
        if el is not None:
            off, ext = el.find(q("a:off")), el.find(q("a:ext"))
            if off is not None and ext is not None:
                x, y = int(off.get("x", "0")), int(off.get("y", "0"))
                return (x, y, x + int(ext.get("cx", "0")), y + int(ext.get("cy", "0")))
    return None


def _placeholder_box(pkg, slide_part, shape):
    """Position of a placeholder that inherits it from the layout or master."""
    ph = shape.find(".//" + q("p:ph"))
    if ph is None:
        return None
    part = slide_part
    for kind in ("/slideLayout", "/slideMaster"):
        nxt = next(iter(pkg.rel_parts(part, kind)), None)
        if not nxt:
            return None
        part = nxt
        for sp in pkg.xml(part).iter(q("p:sp")):
            other = sp.find(".//" + q("p:ph"))
            if other is None:
                continue
            same = (ph.get("idx") and other.get("idx") == ph.get("idx")) or \
                   (not ph.get("idx") and other.get("type", "body") == ph.get("type", "body"))
            if same and _xfrm(sp):
                return _xfrm(sp)
    return None


def _area(b):
    return max(0, b[2] - b[0]) * max(0, b[3] - b[1])


def _overlap(a, b):
    return _area((max(a[0], b[0]), max(a[1], b[1]), min(a[2], b[2]), min(a[3], b[3])))


def check(pkg, findings):
    pres = pkg.xml(pkg.main_part())
    size = pres.find(q("p:sldSz"))
    cx, cy = (int(size.get("cx")), int(size.get("cy"))) if size is not None else (12192000, 6858000)

    notes, hidden, offslide, boxes, comments = [], [], [], [], []
    authors = {}
    for ap in pkg.rel_parts(pkg.main_part(), "/commentAuthors"):
        for a in pkg.xml(ap).iter(q("p:cmAuthor")):
            authors[a.get("id")] = a.get("name", "")
    for ap in pkg.rel_parts(pkg.main_part(), "/authors"):
        for a in pkg.xml(ap).iter(q("p188:author")):
            authors[a.get("id")] = a.get("name", "")

    for num, part in slides(pkg):
        root = pkg.xml(part)
        if root.get("show") in ("0", "false"):
            title = ""
            for sp in root.iter(q("p:sp")):
                ph = sp.find(".//" + q("p:ph"))
                if ph is not None and ph.get("type") in ("title", "ctrTitle"):
                    title = text_of(sp)
                    break
            hidden.append("slide %d%s" % (num, ": " + quoted(title, 50) if title.strip() else ""))
        for np_ in pkg.rel_parts(part, "/notesSlide"):
            t = notes_text(pkg.xml(np_))
            if t.strip():
                notes.append((num, t))
        for cp in pkg.rel_parts(part, "/comments"):
            croot = pkg.xml(cp)
            for cm in list(croot.iter(q("p:cm"))) + list(croot.iter(q("p188:cm"))) + list(croot.iter(q("p188:reply"))):
                text = text_of(cm, "p:text") if cm.tag == q("p:cm") else text_of(cm)
                comments.append((authors.get(cm.get("authorId"), ""), num, text))

        tree = root.find("%s/%s" % (q("p:cSld"), q("p:spTree")))
        if tree is None:
            continue
        texts = []  # (box, text) of shapes earlier in z-order
        for shape in tree:
            if shape.tag not in {q(s) for s in SHAPES}:
                continue
            box = _xfrm(shape) or _placeholder_box(pkg, part, shape)
            text = text_of(shape)
            cnv = shape.find(".//" + q("p:cNvPr"))
            name = cnv.get("name", "shape") if cnv is not None else "shape"
            if box and (box[2] <= 0 or box[3] <= 0 or box[0] >= cx or box[1] >= cy):
                offslide.append("slide %d: %s%s" % (num, name, ": " + quoted(text, 50) if text.strip() else ""))
                continue
            sppr = shape.find(q("p:spPr"))
            fill = sppr.find(q("a:solidFill")) if sppr is not None else None
            if box and fill is not None and _drawing_fill_dark(fill) and not text.strip():
                covered = [t for b, t in texts if _overlap(box, b) > 0.3 * min(_area(box), _area(b))]
                if covered:
                    boxes.append("slide %d: %s over %s" % (num, name, " / ".join(secret(c[:200]) for c in covered)))
            if box and text.strip():
                texts.append((box, text))

    if notes:
        findings.append(Finding("speaker-notes", "speaker notes on " + plural(len(notes), "slide") + " ("
                                + join_names(["%d" % n for n, _ in notes], 8) + ")",
                                ["slide %d: %s" % (n, quoted(t)) for n, t in notes[:8]]))
    if comments:
        who = join_names([a for a, _, _ in comments])
        findings.append(Finding("comments", plural(len(comments), "comment") + (" by " + who if who else ""),
                                ["slide %d %s: %s" % (n, a or "?", quoted(t)) for a, n, t in comments[:8]]))
    if hidden:
        findings.append(Finding("hidden-slides", plural(len(hidden), "hidden slide") + ": "
                                + join_names([h.split(":")[0] for h in hidden], 8), hidden))
    if offslide:
        findings.append(Finding("off-slide-content", plural(len(offslide), "shape") + " parked outside the slide",
                                offslide[:10]))
    if boxes:
        findings.append(Finding("possible-redaction-box", plural(len(boxes), "black box", "black boxes")
                                + " drawn over text; the text underneath can still be copied", boxes[:10], secret=True))
