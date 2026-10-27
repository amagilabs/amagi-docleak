"""Write a cleaned copy of a .docx, .xlsx, .pptx or .pdf file.

clean_file() returns a plain-English list of what it changed. The command line then runs the
normal checks on the copy, so anything left behind is still reported.
"""

import io
import os
import posixpath
import zipfile

from lxml import etree

from . import docx as wd
from . import pdf as pd
from . import xlsx as xl
from .ooxml import Package, q, rels_path, resolve
from .scan import fmt_of
from .util import is_blocked_out, is_dark, plural

BLOCK = "█"
ZIP_DATE = (1980, 1, 1, 0, 0, 0)
DPI = 200


def clean_file(src, dst, keep=frozenset(), unhide=False):
    """Clean src into dst. keep: check ids to leave alone. Returns the actions taken."""
    fmt = fmt_of(src)
    if fmt not in ("docx", "xlsx", "pptx", "pdf"):
        raise ValueError("only .docx, .xlsx, .pptx and .pdf files can be cleaned")
    keep = set(keep)
    tmp = dst + ".part"
    try:
        if fmt == "pdf":
            actions = clean_pdf(src, tmp, keep, unhide)
        else:
            actions = clean_ooxml(src, tmp, fmt, keep, unhide)
        os.replace(tmp, dst)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)
    return actions or ["nothing to clean"]


# ---- Office packages ----------------------------------------------------------------------

def _source_of(rels):
    """'word/_rels/document.xml.rels' -> 'word/document.xml'; '_rels/.rels' -> ''."""
    d, b = posixpath.split(rels)
    b = b[:-len(".rels")]
    return posixpath.join(posixpath.dirname(d), b) if b else ""


class Editor:
    """A package opened for editing: parts can be changed, dropped, then written out."""

    def __init__(self, path):
        self.pkg = Package(path)
        self.dirty, self.removed = set(), set()

    def xml(self, name):
        return self.pkg.xml(name)

    def touch(self, *names):
        self.dirty.update(names)

    def has(self, name):
        return self.pkg.has(name) and name not in self.removed

    def rels(self, part):
        return [r for r in self.pkg.rels(part) if not r.part or r.part not in self.removed]

    def rel_parts(self, part, suffix):
        return [p for p in self.pkg.rel_parts(part, suffix) if p not in self.removed]

    def drop_rel(self, source, rid):
        root = self.xml(rels_path(source))
        if root is None:
            return
        for el in root.iter(q("pr:Relationship")):
            if el.get("Id") == rid:
                root.remove(el)
                self.touch(rels_path(source))
                return

    def drop_part(self, name):
        """Delete a part, its own relationships, every relationship to it and its content type."""
        if not self.has(name):
            return
        self.removed.add(name)
        self.removed.add(rels_path(name))
        for other in self.pkg.names():
            if not other.endswith(".rels") or other in self.removed:
                continue
            root = self.xml(other)
            source = _source_of(other)
            for el in list(root.iter(q("pr:Relationship"))):
                if el.get("TargetMode") != "External" and resolve(source, el.get("Target", "")) == name:
                    root.remove(el)
                    self.touch(other)
        ct = self.xml("[Content_Types].xml")
        for el in list(ct.iter(q("ct:Override"))):
            if el.get("PartName", "").lstrip("/") == name:
                ct.remove(el)
                self.touch("[Content_Types].xml")

    def save(self, path):
        names = [n for n in self.pkg.names() if n not in self.removed and not n.endswith("/")]
        names.sort(key=lambda n: n != "[Content_Types].xml")
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as out:
            for n in names:
                if n in self.dirty:
                    data = etree.tostring(self.xml(n), xml_declaration=True, encoding="UTF-8", standalone=True)
                else:
                    data = self.pkg.read(n)
                info = zipfile.ZipInfo(n, ZIP_DATE)
                info.compress_type = zipfile.ZIP_DEFLATED
                out.writestr(info, data)

    def close(self):
        self.pkg.close()


CORE_DROP = ["dc:creator", "cp:lastModifiedBy", "cp:lastPrinted", "dc:title", "dc:subject", "dc:description",
             "cp:keywords", "cp:category", "cp:contentStatus", "cp:revision"]
APP_DROP = ["ep:Company", "ep:Manager", "ep:HyperlinkBase", "ep:TotalTime", "ep:Template"]


def _remove(el):
    el.getparent().remove(el)


def clean_properties(ed, keep, actions):
    if "author-metadata" not in keep:
        gone = []
        for part, tags in (("docProps/core.xml", CORE_DROP), ("docProps/app.xml", APP_DROP)):
            root = ed.xml(part) if ed.has(part) else None
            if root is None:
                continue
            for tag in tags:
                for el in root.findall(q(tag)):
                    if (el.text or "").strip():
                        gone.append(tag.split(":")[1])
                    _remove(el)
                    ed.touch(part)
        if gone:
            actions.append("cleared author, company and other document properties (%s)"
                           % ", ".join(dict.fromkeys(gone)))
    if "custom-properties" not in keep and ed.has("docProps/custom.xml"):
        ed.drop_part("docProps/custom.xml")
        actions.append("removed custom document properties")
    if "thumbnail" not in keep:
        thumbs = [n for n in ed.pkg.names() if n.startswith("docProps/thumbnail") and ed.has(n)]
        for n in thumbs:
            ed.drop_part(n)
        if thumbs:
            actions.append("removed the saved preview thumbnail")


def clean_ooxml(src, dst, fmt, keep, unhide):
    ed = Editor(src)
    try:
        actions = []
        {"docx": clean_docx, "xlsx": clean_xlsx, "pptx": clean_pptx}[fmt](ed, keep, unhide, actions)
        clean_properties(ed, keep, actions)
        ed.save(dst)
    finally:
        ed.close()
    return actions


# ---- Word ---------------------------------------------------------------------------------

# the order Word requires inside w:rPr; an element out of place makes Word call the file corrupt
RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline",
             "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing",
             "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText",
             "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath"]
COMMENT_RELS = ("/comments", "/commentsExtended", "/commentsIds", "/commentsExtensible", "/people")


def _set_rpr(run, name, value):
    rpr = run.find(q("w:rPr"))
    if rpr is None:
        rpr = etree.Element(q("w:rPr"))
        run.insert(0, rpr)
    old = rpr.find(q("w:" + name))
    if old is not None:
        rpr.remove(old)
    el = etree.Element(q("w:" + name))
    el.set(q("w:val"), value)
    rank = RPR_ORDER.index(name)
    pos = 0
    for i, child in enumerate(rpr):
        local = etree.QName(child).localname
        if local in RPR_ORDER and RPR_ORDER.index(local) < rank:
            pos = i + 1
    rpr.insert(pos, el)


def _unwrap(el):
    parent = el.getparent()
    i = parent.index(el)
    for child in list(el):
        parent.insert(i, child)
        i += 1
    parent.remove(el)


def _paragraph_text(p):
    return "".join(t.text or "" for t in p.iter(q("w:t")))


def accept_changes(root):
    """Accept every tracked change in one story part. Returns how many changes were accepted."""
    n = 0
    mark_parents = (q("w:rPr"), q("w:trPr"))
    for tag in ("w:moveFromRangeStart", "w:moveFromRangeEnd", "w:moveToRangeStart", "w:moveToRangeEnd"):
        for el in list(root.iter(q(tag))):
            _remove(el)
    for tag in ("w:del", "w:moveFrom"):
        for el in list(root.iter(q(tag))):
            if el.getparent() is not None and el.getparent().tag not in mark_parents:
                _remove(el)
                n += 1
    for tag in ("w:ins", "w:moveTo"):
        for el in list(root.iter(q(tag))):
            if el.getparent() is not None and el.getparent().tag not in mark_parents:
                _unwrap(el)
                n += 1
    # rows: a deleted row goes, an inserted row stays
    for el in list(root.iter(q("w:del"))):
        if el.getparent().tag == q("w:trPr"):
            tr = el.getparent().getparent()
            if tr.getparent() is not None:
                _remove(tr)
            n += 1
    for el in list(root.iter(q("w:ins"))):
        if el.getparent().tag == q("w:trPr"):
            _remove(el)
            n += 1
    for el in list(root.iter(q("w:cellDel"))):
        tc = el.getparent().getparent()
        if tc.getparent() is not None:
            _remove(tc)
        n += 1
    for tag in ("w:cellIns", "w:cellMerge"):
        for el in list(root.iter(q(tag))):
            _remove(el)
            n += 1
    # paragraph marks: an inserted mark stays; a deleted mark joins the paragraph to the next one
    for el in list(root.iter(q("w:ins"))):
        if el.getparent().tag == q("w:rPr"):
            _remove(el)
            n += 1
    for el in list(root.iter(q("w:del"))):
        if el.getparent().tag != q("w:rPr"):
            continue
        rpr = el.getparent()
        rpr.remove(el)
        n += 1
        p = rpr.getparent().getparent() if rpr.getparent() is not None else None
        if p is None or p.tag != q("w:p") or p.getparent() is None:
            continue
        nxt = p.getnext()
        if not _paragraph_text(p).strip():
            if nxt is not None and nxt.tag == q("w:p"):
                _remove(p)
        elif nxt is not None and nxt.tag == q("w:p"):
            for child in list(nxt):
                if child.tag != q("w:pPr"):
                    p.append(child)
            _remove(nxt)
    for tag in wd.FORMAT_CHANGE_TAGS:
        if tag in ("w:cellIns", "w:cellDel", "w:cellMerge"):
            continue
        for el in list(root.iter(q(tag))):
            _remove(el)
            n += 1
    return n


def remove_comment_marks(root):
    for tag in ("w:commentRangeStart", "w:commentRangeEnd"):
        for el in list(root.iter(q(tag))):
            _remove(el)
    for el in list(root.iter(q("w:commentReference"))):
        run = el.getparent()
        run.remove(el)
        if run.tag == q("w:r") and all(c.tag == q("w:rPr") for c in run):
            _remove(run)


def clean_docx(ed, keep, unhide, actions):
    main = ed.pkg.main_part()
    parts = wd.story_parts(ed.pkg, main)
    settings_part = next(iter(ed.rel_parts(main, "/settings")), None)
    settings = ed.xml(settings_part) if settings_part else None

    if "tracked-changes" not in keep:
        n = sum(accept_changes(ed.xml(p)) for p in parts)
        ed.touch(*parts)
        if settings is not None:
            for el in settings.findall(q("w:trackRevisions")):
                _remove(el)
                ed.touch(settings_part)
                actions.append("switched Track Changes off")
        if n:
            actions.insert(0, "accepted " + plural(n, "tracked change"))

    if "comments" not in keep:
        n, dropped = 0, False
        for suffix in COMMENT_RELS:
            for part in ed.rel_parts(main, suffix):
                if suffix == "/comments":
                    n += sum(1 for _ in ed.xml(part).iter(q("w:comment")))
                ed.drop_part(part)
                dropped = True
        for p in parts:
            remove_comment_marks(ed.xml(p))
        if dropped:
            actions.append("removed " + (plural(n, "comment") if n else "the comments") + " and the reviewer list")

    styles_part = next(iter(ed.rel_parts(main, "/styles")), None)
    styles = wd.Styles(ed.xml(styles_part) if styles_part else None)
    hidden = redacted = 0
    for part in parts:
        root = ed.xml(part)
        bg = wd.page_background(root) if part == main else wd.WHITE
        for run, kind, text in list(wd.classify_runs(root, styles, bg)):
            if is_blocked_out(text):
                continue
            if kind == "redaction":
                if "fake-redaction" in keep:
                    continue
                for t in run.iter(q("w:t")):
                    t.text = BLOCK * len(t.text or "")
                redacted += 1
            elif "hidden-text" not in keep:
                hidden += 1
                if unhide:
                    if kind == "vanish":
                        _set_rpr(run, "vanish", "0")
                    elif kind == "tiny":
                        _set_rpr(run, "sz", "22")
                    else:
                        _set_rpr(run, "color", "auto")
                elif run.getparent() is not None:
                    _remove(run)
        ed.touch(part)
    if redacted:
        actions.append("replaced %s of blacked-out text with solid blocks" % plural(redacted, "run"))
    if hidden:
        actions.append(("made %s of hidden text visible" if unhide else "deleted %s of hidden text")
                       % plural(hidden, "run"))

    if settings is not None:
        if "document-variables" not in keep:
            for el in settings.findall(q("w:docVars")):
                _remove(el)
                ed.touch(settings_part)
                actions.append("removed the document variables")
        if "file-paths" not in keep:
            for el in settings.findall(q("w:attachedTemplate")):
                rid = el.get(q("r:id"))
                _remove(el)
                ed.touch(settings_part)
                if rid:
                    ed.drop_rel(settings_part, rid)
                actions.append("removed the attached template path")


# ---- Excel --------------------------------------------------------------------------------

VML_X = "urn:schemas-microsoft-com:office:excel"
VML_V = "urn:schemas-microsoft-com:vml"


def _drop_note_shapes(ed, sheet_part, ws):
    """Remove comment boxes from a sheet's VML drawing; drop the drawing if nothing else is in it."""
    ld = ws.find(q("x:legacyDrawing"))
    if ld is None:
        return
    rid = ld.get(q("r:id"))
    rel = next((r for r in ed.rels(sheet_part) if r.id == rid), None)
    if rel is None or rel.external or not ed.has(rel.part):
        return
    try:
        vml = etree.fromstring(ed.pkg.read(rel.part), etree.XMLParser(resolve_entities=False, no_network=True,
                                                                       recover=True))
    except etree.XMLSyntaxError:
        return
    shapes = list(vml.iter("{%s}shape" % VML_V))
    notes = [s for s in shapes if any(cd.get("ObjectType") == "Note" for cd in s.iter("{%s}ClientData" % VML_X))]
    if notes and len(notes) == len(shapes):
        ed.drop_part(rel.part)
        ws.remove(ld)
        ed.touch(sheet_part)
    elif notes:
        for s in notes:
            s.getparent().remove(s)
        ed.pkg._xml[rel.part] = vml
        ed.touch(rel.part)


def clean_xlsx(ed, keep, unhide, actions):
    wb = xl.Workbook(ed.pkg)
    book = wb.root

    if "comments" not in keep:
        n = 0
        for _, _, part, _ in wb.sheets:
            if not part or not ed.has(part):
                continue
            for suffix in ("/comments", "/threadedComment"):
                for cp in ed.rel_parts(part, suffix):
                    if suffix == "/comments":
                        n += sum(1 for _ in ed.xml(cp).iter(q("x:comment")))
                    ed.drop_part(cp)
            _drop_note_shapes(ed, part, ed.xml(part))
        for pp in ed.rel_parts(wb.part, "/person"):
            ed.drop_part(pp)
        if n:
            actions.append("removed %s and the reviewer list" % plural(n, "comment"))

    if unhide and "hidden-sheets" not in keep:
        shown = []
        for name, state, _, sh in wb.sheets:
            if state in ("hidden", "veryHidden"):
                del sh.attrib["state"]
                shown.append(name)
        if shown:
            ed.touch(wb.part)
            actions.append("unhid " + plural(len(shown), "sheet") + ": " + ", ".join(shown))

    redacted, used_by_redaction = 0, set()
    rows_shown = 0
    for name, _, part, _ in wb.sheets:
        if not part or not ed.has(part):
            continue
        ws = ed.xml(part)
        if unhide and "hidden-rows-columns" not in keep:
            for el in list(ws.iter(q("x:row"))) + list(ws.iter(q("x:col"))):
                if el.get("hidden") in ("1", "true"):
                    del el.attrib["hidden"]
                    rows_shown += 1
                    ed.touch(part)
        if "fake-redaction" in keep:
            continue
        for c in ws.iter(q("x:c")):
            val = wb.value(c)
            if val == "" or is_blocked_out(val) or wb.cell_style(c) != "redaction":
                continue
            if c.get("t") == "s":
                v = c.find(q("x:v"))
                used_by_redaction.add(v.text)
            for child in list(c):
                c.remove(child)
            c.set("t", "inlineStr")
            is_ = etree.SubElement(c, q("x:is"))
            etree.SubElement(is_, q("x:t")).text = BLOCK * len(val)
            redacted += 1
            ed.touch(part)
    if rows_shown:
        actions.append("unhid %s" % plural(rows_shown, "row or column", "rows and columns"))
    if redacted:
        _blank_unused_strings(ed, wb, used_by_redaction)
        for cc in ed.rel_parts(wb.part, "/calcChain"):
            ed.drop_part(cc)  # Excel rebuilds it; stale entries would point at the removed formulas
        actions.append("replaced %s with solid blocks (value and formula removed)"
                       % plural(redacted, "blacked-out cell"))

    if "pivot-cache" not in keep:
        n = 0
        for cd in ed.rel_parts(wb.part, "/pivotCacheDefinition"):
            records = ed.rel_parts(cd, "/pivotCacheRecords")
            if not records:
                continue
            for r in records:
                ed.drop_part(r)
            root = ed.xml(cd)
            root.attrib.pop(q("r:id"), None)
            root.set("saveData", "0")
            ed.touch(cd)
            n += 1
        if n:
            actions.append("dropped the cached source rows of " + plural(n, "pivot table"))

    if "file-paths" not in keep:
        found = False
        for el in list(book.iter(q("x15ac:absPath"))):
            found = True
            box = el.getparent()
            while box is not None and box.tag != q("mc:AlternateContent"):
                box = box.getparent()
            _remove(box if box is not None else el)
        if found:
            ed.touch(wb.part)
            actions.append("removed the folder the workbook was last saved in")


def _blank_unused_strings(ed, wb, indexes):
    ss = next(iter(ed.rel_parts(wb.part, "/sharedStrings")), None)
    if not ss or not indexes:
        return
    still_used = set()
    for _, _, part, _ in wb.sheets:
        if part and ed.has(part):
            for c in ed.xml(part).iter(q("x:c")):
                if c.get("t") == "s":
                    v = c.find(q("x:v"))
                    if v is not None:
                        still_used.add(v.text)
    root = ed.xml(ss)
    items = list(root.iter(q("x:si")))
    for idx in indexes - still_used:
        try:
            si = items[int(idx)]
        except (TypeError, ValueError, IndexError):
            continue
        for child in list(si):
            si.remove(child)
        etree.SubElement(si, q("x:t"))
    ed.touch(ss)


# ---- PowerPoint ---------------------------------------------------------------------------

def clean_pptx(ed, keep, unhide, actions):
    from .pptx import slides

    pres = ed.pkg.main_part()
    deck = slides(ed.pkg)
    if "speaker-notes" not in keep:
        n = 0
        for _, part in deck:
            for np_ in ed.rel_parts(part, "/notesSlide"):
                ed.drop_part(np_)
                n += 1
        if n:
            actions.append("removed speaker notes from " + plural(n, "slide"))
    if "comments" not in keep:
        n = 0
        for _, part in deck:
            for cp in ed.rel_parts(part, "/comments"):
                root = ed.xml(cp)
                n += sum(1 for _ in root.iter(q("p:cm"))) + sum(1 for _ in root.iter(q("p188:cm")))
                ed.drop_part(cp)
        for suffix in ("/commentAuthors", "/authors"):
            for ap in ed.rel_parts(pres, suffix):
                ed.drop_part(ap)
        if n:
            actions.append("removed %s and the reviewer list" % plural(n, "comment"))
    if unhide and "hidden-slides" not in keep:
        shown = []
        for num, part in deck:
            root = ed.xml(part)
            if root.get("show") in ("0", "false"):
                del root.attrib["show"]
                ed.touch(part)
                shown.append(str(num))
        if shown:
            actions.append("unhid " + plural(len(shown), "slide") + " (" + ", ".join(shown) + ")")


# ---- PDF ----------------------------------------------------------------------------------

def _dark_box(annot):
    st = str(annot.get("/Subtype", ""))
    if st not in pd.BOX_TYPES:
        return False
    fill = pd.annot_rgb(annot.get("/C")) if st == "/Highlight" else pd.annot_rgb(annot.get("/IC"))
    return fill is not None and is_dark(fill)


def _pages_to_flatten(pdf, data, keep):
    """{page index: [boxes in user space to paint black]} for pages whose blacked-out text must go."""
    marks, dark, dark_dev = {}, {}, {}
    for i, page in enumerate(pdf.pages):
        ctm = pd.page_ctm(page.obj)
        for annot in page.obj.get("/Annots", []):
            if not isinstance(annot, pd.pikepdf.Dictionary):
                continue
            if str(annot.get("/Subtype", "")) == "/Redact":
                marks.setdefault(i, []).extend(pd.quad_boxes(annot))
            elif _dark_box(annot):
                dark.setdefault(i, []).extend(pd.quad_boxes(annot))
                dark_dev.setdefault(i, []).extend(pd.apply(ctm, b) for b in pd.quad_boxes(annot))
    out = {}
    if "unapplied-redaction" not in keep:
        for i, boxes in marks.items():
            out.setdefault(i, []).extend(boxes)
    if "fake-redaction" not in keep:
        for i, items in enumerate(pd.layouts(data)):
            fake, _ = pd.covered(items, dark_dev.get(i, ()))
            if fake:
                out.setdefault(i, []).extend(dark.get(i, ()))
    return out


def _is_script(action):
    return isinstance(action, pd.pikepdf.Dictionary) and pd._action_kind(action) is not None


def clean_pdf(src, dst, keep, unhide):
    import pikepdf

    with open(src, "rb") as fh:
        data = fh.read()
    try:
        pdf = pd.open_pdf(data)
    except pikepdf.PasswordError:
        raise ValueError("password-protected PDF; save an unprotected copy first")
    actions = []
    with pdf:
        signed = pd.is_signed(pdf)
        try:
            flatten = _pages_to_flatten(pdf, data, keep)
        except Exception as e:  # pdfminer could not read the text; do not pretend the boxes are fine
            raise ValueError("could not read the page text to find fake redactions (%s)" % type(e).__name__)
        root = pdf.Root

        if "author-metadata" not in keep or "custom-properties" not in keep:
            if "/Info" in pdf.trailer:
                info = pdf.trailer.Info
                gone = []
                for key in list(info.keys()):
                    k = str(key)
                    personal = k in pd.PERSONAL_INFO
                    if k in pd.STANDARD_INFO:
                        continue
                    if (personal and "author-metadata" not in keep) or \
                            (not personal and "custom-properties" not in keep):
                        if pd._str(info[key]):
                            gone.append(k[1:])
                        del info[key]
                if gone:
                    actions.append("cleared document info fields (%s)" % ", ".join(gone))
            if "author-metadata" not in keep and "/Metadata" in root:
                del root["/Metadata"]
                actions.append("removed the XMP metadata packet")

        removed = {"comments": 0, "attachments": 0, "scripts": 0}
        for i, page in enumerate(pdf.pages):
            obj = page.obj
            if "/AA" in obj and "javascript" not in keep:
                del obj["/AA"]
                removed["scripts"] += 1
            if "/PieceInfo" in obj and "private-app-data" not in keep:
                del obj["/PieceInfo"]
            if "/Annots" not in obj:
                continue
            kept = []
            for annot in obj.Annots:
                if not isinstance(annot, pikepdf.Dictionary):
                    kept.append(annot)
                    continue
                st = str(annot.get("/Subtype", ""))
                if st == "/Popup" and "comments" not in keep:
                    continue
                if st in pd.COMMENT_TYPES and not _dark_box(annot) and "comments" not in keep:
                    removed["comments"] += 1
                    continue
                if st == "/FileAttachment" and "embedded-files" not in keep:
                    removed["attachments"] += 1
                    continue
                if "javascript" not in keep:
                    if _is_script(annot.get("/A")):
                        del annot["/A"]
                        removed["scripts"] += 1
                    if "/AA" in annot:
                        del annot["/AA"]
                        removed["scripts"] += 1
                kept.append(annot)
            obj.Annots = pikepdf.Array(kept)
        if removed["comments"]:
            actions.append("removed " + plural(removed["comments"], "comment"))

        names = root.get("/Names")
        if "embedded-files" not in keep:
            n = 0
            if names is not None and "/EmbeddedFiles" in names:
                n = len(pdf.attachments)
                del names["/EmbeddedFiles"]
            if "/Collection" in root:
                del root["/Collection"]
            n += removed["attachments"]
            if n:
                actions.append("removed " + plural(n, "attached file"))
        if "javascript" not in keep:
            n = removed["scripts"]
            if names is not None and "/JavaScript" in names:
                del names["/JavaScript"]
                n += 1
            if _is_script(root.get("/OpenAction")):
                del root["/OpenAction"]
                n += 1
            if "/AA" in root:
                del root["/AA"]
                n += 1
            if "/AcroForm" in root and "/Fields" in root.AcroForm:
                stack = list(root.AcroForm.Fields)
                while stack:
                    f = stack.pop()
                    if isinstance(f, pikepdf.Dictionary):
                        if "/AA" in f:
                            del f["/AA"]
                            n += 1
                        stack.extend(f.get("/Kids", []))
            if n:
                actions.append("removed " + plural(n, "script or launch action"))
        if "private-app-data" not in keep and "/PieceInfo" in root:
            del root["/PieceInfo"]
            actions.append("removed private application data")
        if unhide and "hidden-layers" not in keep:
            oc = root.get("/OCProperties")
            if oc is not None and "/D" in oc and len(oc.D.get("/OFF", [])):
                n = len(oc.D.OFF)
                on = list(oc.D.get("/ON", [])) + list(oc.D.OFF)
                oc.D.ON = pikepdf.Array(on)
                oc.D.OFF = pikepdf.Array([])
                actions.append("switched on " + plural(n, "hidden layer"))

        if flatten:
            _flatten(pdf, flatten)
            pages = ", ".join(str(i + 1) for i in sorted(flatten))
            actions.append("replaced page %s with %s image%s: the black boxes are burned in and the text "
                           "under them is gone (that page can no longer be searched or copied from)"
                           % (pages, "a %d dpi" % DPI, "s" if len(flatten) > 1 else ""))
        pdf.save(dst, deterministic_id=True)
        revised = len(pd.revisions(data, pdf.is_linearized)) > 1
    if revised or signed:
        actions.append("rewrote the file, which drops the earlier versions saved inside it"
                       + ("; the digital signature is no longer valid" if signed else ""))
    return actions


def _display_map(page):
    """User-space point -> top-left-origin display point, matching how pdfium renders the page."""
    box = pd.inherited(page, "/CropBox") or pd.inherited(page, "/MediaBox")
    x0, y0, x1, y1 = [float(v) for v in box] if box is not None else (0, 0, 612, 792)
    x0, x1 = min(x0, x1), max(x0, x1)
    y0, y1 = min(y0, y1), max(y0, y1)
    w, h = x1 - x0, y1 - y0
    rot = int(pd.inherited(page, "/Rotate") or 0) % 360

    def to_display(x, y):
        u, v = x - x0, y - y0
        if rot == 90:
            return v, u
        if rot == 180:
            return w - u, v
        if rot == 270:
            return h - v, w - u
        return u, h - v
    return to_display


def _flatten(pdf, flatten):
    """Swap each listed page for a picture of itself with the given boxes painted black."""
    import pikepdf
    import pypdfium2
    from PIL import ImageDraw

    buf = io.BytesIO()
    pdf.save(buf)
    doc = pypdfium2.PdfDocument(buf.getvalue())
    scale = DPI / 72.0
    try:
        for i, boxes in sorted(flatten.items()):
            image = doc[i].render(scale=scale, draw_annots=True).to_pil().convert("RGB")
            page = pdf.pages[i].obj
            to_display = _display_map(page)
            draw = ImageDraw.Draw(image)
            for b in boxes:
                pts = [to_display(x, y) for x, y in ((b[0], b[1]), (b[2], b[3]))]
                xs = sorted(p[0] * scale for p in pts)
                ys = sorted(p[1] * scale for p in pts)
                draw.rectangle([xs[0], ys[0], xs[1], ys[1]], fill=(0, 0, 0))
            jpeg = io.BytesIO()
            image.save(jpeg, "JPEG", quality=85)
            width, height = image.size[0] / scale, image.size[1] / scale
            xobj = pikepdf.Stream(pdf, jpeg.getvalue())
            xobj.Type, xobj.Subtype = pikepdf.Name.XObject, pikepdf.Name.Image
            xobj.Width, xobj.Height = image.size
            xobj.ColorSpace, xobj.BitsPerComponent = pikepdf.Name.DeviceRGB, 8
            xobj.Filter = pikepdf.Name.DCTDecode
            content = ("q %.4f 0 0 %.4f 0 0 cm /Im0 Do Q" % (width, height)).encode()
            for key in ("/CropBox", "/TrimBox", "/BleedBox", "/ArtBox", "/Annots", "/Thumb", "/PieceInfo",
                        "/StructParents", "/Group", "/AA"):
                if key in page:
                    del page[key]
            page.MediaBox = pikepdf.Array([0, 0, width, height])
            page.Rotate = 0
            page.Resources = pikepdf.Dictionary(XObject=pikepdf.Dictionary(Im0=xobj))
            page.Contents = pikepdf.Stream(pdf, content)
    finally:
        doc.close()
    # the structure tree holds the page's tagged text; the flattened page no longer matches it
    for key in ("/StructTreeRoot", "/MarkInfo"):
        if key in pdf.Root:
            del pdf.Root[key]
