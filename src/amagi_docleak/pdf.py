"""PDF checks."""

import io
import re

import pikepdf
from lxml import etree
from pdfminer.converter import PDFPageAggregator
from pdfminer.high_level import extract_text
from pdfminer.layout import LTChar, LTCurve, LTFigure, LTImage
from pdfminer.pdfinterp import PDFPageInterpreter, PDFResourceManager
from pdfminer.pdfpage import PDFPage

from .model import Finding
from .ooxml import PARSER
from .util import (UNREADABLE, contrast, find_paths, is_blocked_out, is_dark, join_names, luminance, plural,
                   quoted, secret)

COMMENT_TYPES = {"/Text", "/FreeText", "/Highlight", "/Underline", "/StrikeOut", "/Squiggly", "/Caret",
                 "/Ink", "/Stamp", "/Square", "/Circle", "/Line", "/Polygon", "/PolyLine", "/Sound"}
BOX_TYPES = {"/Square", "/Circle", "/Polygon", "/Highlight"}
STANDARD_INFO = {"/Producer", "/Creator", "/CreationDate", "/ModDate", "/Trapped"}
PERSONAL_INFO = {"/Author", "/Title", "/Subject", "/Keywords", "/Company", "/Manager", "/LastModifiedBy"}
XMP_PERSONAL = {"creator", "title", "description", "subject", "Keywords", "Author", "Company", "Manager",
                "AuthorsPosition", "CaptionWriter", "LastModifiedBy"}
PDFX_NS = "http://ns.adobe.com/pdfx/1.3/"
MAX_REVISION_BYTES = 50 * 1024 * 1024
MAX_REVISIONS = 5
COVER = 0.6  # share of a character's area a box must cover to count as covering it
WHITE = (1.0, 1.0, 1.0)
BLACK = (0.0, 0.0, 0.0)


# ---- geometry -----------------------------------------------------------------------------

def inherited(page, key):
    obj = page
    for _ in range(64):
        if key in obj:
            return obj[key]
        if "/Parent" not in obj:
            return None
        obj = obj.Parent
    return None


def page_ctm(page):
    """The same device transform pdfminer applies, so annotation and content boxes line up."""
    box = inherited(page, "/MediaBox")
    x0, y0, x1, y1 = [float(v) for v in box] if box is not None else (0, 0, 612, 792)
    x0, x1 = min(x0, x1), max(x0, x1)
    y0, y1 = min(y0, y1), max(y0, y1)
    rot = int(inherited(page, "/Rotate") or 0) % 360
    if rot == 90:
        return (0, -1, 1, 0, -y0, x1)
    if rot == 180:
        return (-1, 0, 0, -1, x1, y1)
    if rot == 270:
        return (0, 1, -1, 0, y1, -x0)
    return (1, 0, 0, 1, -x0, -y0)


def apply(m, rect):
    a, b, c, d, e, f = m
    pts = [(x * a + y * c + e, x * b + y * d + f) for x, y in
           ((rect[0], rect[1]), (rect[0], rect[3]), (rect[2], rect[1]), (rect[2], rect[3]))]
    xs, ys = [p[0] for p in pts], [p[1] for p in pts]
    return (min(xs), min(ys), max(xs), max(ys))


def area(b):
    return max(0.0, b[2] - b[0]) * max(0.0, b[3] - b[1])


def overlap(a, b):
    return area((max(a[0], b[0]), max(a[1], b[1]), min(a[2], b[2]), min(a[3], b[3])))


def quad_boxes(annot):
    qp = annot.get("/QuadPoints")
    if qp is not None and len(qp) >= 8:
        v = [float(x) for x in qp]
        out = []
        for i in range(0, len(v) - 7, 8):
            xs, ys = v[i:i + 8:2], v[i + 1:i + 8:2]
            out.append((min(xs), min(ys), max(xs), max(ys)))
        return out
    r = annot.get("/Rect")
    if r is None or len(r) != 4:
        return []
    r = [float(x) for x in r]
    return [(min(r[0], r[2]), min(r[1], r[3]), max(r[0], r[2]), max(r[1], r[3]))]


# ---- colors -------------------------------------------------------------------------------

def to_rgb(value, default=BLACK):
    """pdfminer color value -> rgb, None when it is a pattern or an indexed color we cannot map."""
    if value is None:
        return default
    if isinstance(value, (int, float)):
        value = (value,)
    if not isinstance(value, (tuple, list)) or not all(isinstance(v, (int, float)) for v in value):
        return None
    v = [min(1.0, max(0.0, float(x))) for x in value]
    if len(v) == 1:
        return (v[0], v[0], v[0])
    if len(v) == 3:
        return tuple(v)
    if len(v) == 4:
        c, m, y, k = v
        return ((1 - c) * (1 - k), (1 - m) * (1 - k), (1 - y) * (1 - k))
    return None


def annot_rgb(arr):
    if arr is None:
        return None
    return to_rgb(tuple(float(x) for x in arr), None) if len(arr) else None


# ---- content stream walk ------------------------------------------------------------------

class _Device(PDFPageAggregator):
    """Keeps the text render mode on each character so invisible OCR text can be skipped."""

    _tr = 0

    def render_string(self, textstate, seq, ncs, graphicstate):
        self._tr = textstate.render
        return super().render_string(textstate, seq, ncs, graphicstate)

    def render_char(self, *args, **kwargs):
        adv = super().render_char(*args, **kwargs)
        self.cur_item._objs[-1].render = self._tr
        return adv


def _walk(item, out):
    for obj in item:
        if isinstance(obj, LTChar):
            if getattr(obj, "render", 0) in (3, 7) or not obj.get_text().strip():
                continue
            gs = obj.graphicstate
            out.append(("char", obj.bbox, to_rgb(gs.ncolor if gs is not None else None), obj.get_text()))
        elif isinstance(obj, LTImage):
            out.append(("image", obj.bbox, None, ""))
        elif isinstance(obj, LTFigure):
            _walk(obj, out)
        elif isinstance(obj, LTCurve) and obj.fill:
            out.append(("box", obj.bbox, to_rgb(obj.non_stroking_color), ""))


def layouts(data):
    """Yield the z-ordered items of every page: (kind, bbox, rgb, text)."""
    rsrc = PDFResourceManager()
    device = _Device(rsrc, laparams=None)
    interp = PDFPageInterpreter(rsrc, device)
    for page in PDFPage.get_pages(io.BytesIO(data), check_extractable=False):
        interp.process_page(page)
        items = []
        _walk(device.get_result(), items)
        yield items


def _join(chars):
    chars = sorted(chars, key=lambda c: (-round(c[1][1] / 3), c[1][0]))
    return "".join(c[3] for c in chars)


class _Grid:
    """Bucket characters by position so box lookups stay fast on dense pages."""

    def __init__(self, chars, size=40.0):
        self.size, self.cells = size, {}
        for c in chars:
            for key in self._keys(c[1]):
                self.cells.setdefault(key, []).append(c)

    def _keys(self, b):
        s = self.size
        for gx in range(int(b[0] // s), int(b[2] // s) + 1):
            for gy in range(int(b[1] // s), int(b[3] // s) + 1):
                yield gx, gy

    def near(self, box):
        if area(box) > 4e6:  # absurd boxes: fall back to all characters
            seen = {id(c): c for cell in self.cells.values() for c in cell}
            return list(seen.values())
        seen = {}
        for key in self._keys(box):
            for c in self.cells.get(key, ()):
                seen[id(c)] = c
        return list(seen.values())


def covered(items, extra_boxes=()):
    """Find text hidden under shapes drawn after it, and text drawn unreadably on a dark box.

    Returns (fake, covered) lists of (box, text) pairs.
    """
    chars = [(i,) + it[1:] for i, it in enumerate(items) if it[0] == "char"]
    grid = _Grid([(c[0], c[1], c[2], c[3]) for c in chars])
    fake, cover = [], []
    for i, (kind, bbox, rgb, _) in enumerate(items):
        if kind == "char" or area(bbox) < 4:
            continue
        under = [c for c in grid.near(bbox) if c[0] < i and area(c[1]) > 0
                 and overlap(c[1], bbox) >= COVER * area(c[1])]
        if not under:
            continue
        text = _join(under)
        if not text.strip() or is_blocked_out(text):
            continue
        if kind == "box" and rgb is not None and is_dark(rgb):
            fake.append((bbox, text))
        elif kind == "image" or rgb is not None:
            cover.append((bbox, text))
    # text drawn on top of a dark box in a color that cannot be read against it
    dark = [(i, it[1], it[2]) for i, it in enumerate(items) if it[0] == "box" and it[2] is not None
            and is_dark(it[2]) and area(it[1]) >= 4]
    if dark:
        hits = {}
        for c in chars:
            if c[2] is None:
                continue
            for i, bbox, rgb in dark:
                if i < c[0] and overlap(c[1], bbox) >= COVER * area(c[1]) and contrast(c[2], rgb) < UNREADABLE:
                    hits.setdefault(bbox, []).append(c)
                    break
        for bbox, cs in hits.items():
            text = _join(cs)
            if text.strip() and not is_blocked_out(text):
                fake.append((bbox, text))
    for bbox in extra_boxes:
        under = [c for c in grid.near(bbox) if area(c[1]) > 0 and overlap(c[1], bbox) >= COVER * area(c[1])]
        text = _join(under)
        if text.strip() and not is_blocked_out(text):
            fake.append((bbox, text))
    return fake, cover


def text_under(items, boxes):
    chars = [it for it in items if it[0] == "char"]
    grid = _Grid(chars)
    seen, out = set(), []
    for b in boxes:
        for c in grid.near(b):
            if id(c) not in seen and area(c[1]) > 0 and overlap(c[1], b) >= COVER * area(c[1]):
                seen.add(id(c))
                out.append(c)
    return _join(out)


# ---- metadata -----------------------------------------------------------------------------

def _str(v):
    try:
        return str(v).strip()
    except Exception:
        return ""


def xmp_values(pdf):
    """(personal, custom, all_text) from the XMP packet."""
    personal, custom, texts = [], [], []
    if "/Metadata" not in pdf.Root:
        return personal, custom, texts
    try:
        root = etree.fromstring(pdf.Root.Metadata.read_bytes(), PARSER)
    except Exception:
        return personal, custom, texts
    for el in root.iter():
        if not isinstance(el.tag, str):
            continue
        texts.extend(v for v in el.attrib.values())
        if el.text and el.text.strip():
            texts.append(el.text.strip())
        local = etree.QName(el).localname
        ns = etree.QName(el).namespace
        if local in XMP_PERSONAL:
            leaves = [t.strip() for t in el.itertext() if t.strip()]
            if leaves:
                personal.append((local, " / ".join(dict.fromkeys(leaves))))
        elif ns == PDFX_NS and el.text and el.text.strip():
            custom.append((local, el.text.strip()))
    return personal, custom, texts


def check_metadata(pdf, findings, paths):
    personal, custom, context = [], [], []
    for key, value in (pdf.docinfo.items() if "/Info" in pdf.trailer else []):
        key = str(key)
        v = _str(value)
        if not v:
            continue
        paths.extend((p, "document info %s" % key) for p in find_paths(v))
        if key in ("/Producer", "/Creator"):
            context.append("%s: %s" % (key[1:].lower(), v))
        elif key in PERSONAL_INFO:
            personal.append((key[1:], v))
        elif key not in STANDARD_INFO:
            custom.append((key[1:], v))
    xp, xc, texts = xmp_values(pdf)
    for p in (p for t in texts for p in find_paths(t)):
        paths.append((p, "XMP metadata"))
    have = {(k.lower(), v) for k, v in personal}
    for k, v in xp:
        name = {"creator": "author", "description": "subject"}.get(k, k)
        if (name.lower(), v) not in have and (k.lower(), v) not in have:
            personal.append(("%s (XMP)" % name, v))
    have_c = {k for k, _ in custom}
    custom.extend((k, v) for k, v in xc if k not in have_c)
    if personal:
        lead = [p for p in personal if p[0].lower().startswith(("author", "company", "manager"))] or personal
        findings.append(Finding("author-metadata", ", ".join("%s %s" % (k.lower(), quoted(v, 40))
                                                              for k, v in lead[:3]),
                                ["%s: %s" % (k, v) for k, v in personal] + context))
    if custom:
        findings.append(Finding("custom-properties", plural(len(custom), "custom property", "custom properties")
                                + " (" + join_names([k for k, _ in custom]) + ")",
                                ["%s: %s" % (k, v) for k, v in custom]))


# ---- actions ------------------------------------------------------------------------------

def _action_kind(action):
    try:
        s = str(action.get("/S", ""))
    except Exception:
        return None
    if s in ("/JavaScript", "/Launch"):
        return s[1:]
    return None


def _filespec_name(fs):
    if isinstance(fs, pikepdf.Dictionary):
        return _str(fs.get("/UF") or fs.get("/F") or "")
    return _str(fs)


def check_actions(pdf, findings, paths):
    where = []
    root = pdf.Root
    names = root.get("/Names")
    if names is not None and "/JavaScript" in names:
        where.append("document-level JavaScript")
    oa = root.get("/OpenAction")
    if isinstance(oa, pikepdf.Dictionary) and _action_kind(oa):
        where.append("%s action when the file opens" % _action_kind(oa))
    if "/AA" in root:
        where.append("document event actions")
    if "/AcroForm" in root and "/Fields" in root.AcroForm:
        stack, n = list(root.AcroForm.Fields), 0
        while stack and n < 5000:
            n += 1
            f = stack.pop()
            if not isinstance(f, pikepdf.Dictionary):
                continue
            if "/AA" in f:
                where.append("form field %s has event actions" % (_str(f.get("/T", "?")) or "?"))
            stack.extend(f.get("/Kids", []))
    for num, page in enumerate(pdf.pages, 1):
        if "/AA" in page.obj:
            where.append("page %d open/close actions" % num)
        for annot in page.obj.get("/Annots", []):
            if not isinstance(annot, pikepdf.Dictionary):
                continue
            a = annot.get("/A")
            if isinstance(a, pikepdf.Dictionary):
                kind = _action_kind(a)
                if kind:
                    where.append("page %d: %s action on a link" % (num, kind))
                if a.get("/S") in ("/Launch", "/GoToR") and "/F" in a:
                    name = _filespec_name(a.F)
                    if name:
                        paths.append((name, "page %d link target" % num))
    if where:
        findings.append(Finding("javascript", plural(len(where), "script or launch action") + ": "
                                + join_names(where, 2), where[:10]))


# ---- prior revisions ----------------------------------------------------------------------

EOF_RE = re.compile(rb"startxref\s+\d+\s+%%EOF")


OBJ_RE = re.compile(rb"(?<![0-9])\d+\s+\d+\s+obj\b")
XREF_STREAM_RE = re.compile(rb"/Type\s*/XRef\b")


def _real_update(section):
    """An appended section counts as a revision only if it changes objects.

    Word saves "hybrid" files whose second section holds nothing but a cross-reference
    table or stream; that is not an earlier version of anything.
    """
    for m in OBJ_RE.finditer(section):
        head = section[m.end():m.end() + 400]
        if not XREF_STREAM_RE.search(head.split(b"stream", 1)[0]):
            return True
    return False


def revisions(data, linearized):
    """Byte offsets where each saved revision ends (the last one is the current file)."""
    ends = [m.end() for m in EOF_RE.finditer(data)]
    if linearized and len(ends) > 1:
        ends = ends[1:]
    kept = ends[:1]
    for prev, end in zip(ends, ends[1:]):
        if _real_update(data[prev:end]):
            kept.append(end)
        elif kept:
            kept[-1] = end
    return kept


def _lines(data):
    try:
        text = extract_text(io.BytesIO(data))
    except Exception:
        return None
    return [" ".join(l.split()) for l in text.splitlines() if l.strip()]


def check_revisions(pdf, data, findings, signed):
    ends = revisions(data, pdf.is_linearized)
    if len(ends) < 2:
        return
    earlier = ends[:-1]
    details = []
    if signed:
        details.append("the file is digitally signed; signing adds a revision, and rewriting the file "
                       "breaks the signature")
    if len(data) <= MAX_REVISION_BYTES:
        current = _lines(data)
        current_set = set(current or [])
        for n, end in enumerate(earlier[-MAX_REVISIONS:], max(1, len(earlier) - MAX_REVISIONS + 1)):
            old = _lines(data[:end])
            if old is None or current is None:
                details.append("revision %d: could not read its text" % n)
                continue
            gone = [l for l in dict.fromkeys(old) if l not in current_set]
            if gone:
                details.append("revision %d: %s no longer in the current version, e.g. %s"
                               % (n, plural(len(gone), "line"), secret(gone[0][:120])))
                details.extend("  removed: " + secret(l[:120]) for l in gone[1:4])
            else:
                details.append("revision %d: same text as the current version (other objects changed)" % n)
    else:
        details.append("file is over %d MB; earlier revisions were counted but not compared"
                       % (MAX_REVISION_BYTES // 2 ** 20))
    findings.append(Finding("prior-versions", "%s saved before the current one, still inside the file"
                            % plural(len(earlier), "earlier version"), details, secret=True))


def is_signed(pdf):
    acro = pdf.Root.get("/AcroForm")
    if acro is None:
        return False
    if int(acro.get("/SigFlags", 0)) & 1:
        return True
    for f in acro.get("/Fields", []):
        if isinstance(f, pikepdf.Dictionary) and f.get("/FT") == "/Sig" and "/V" in f:
            return True
    return False


# ---- main ---------------------------------------------------------------------------------

def open_pdf(path_or_bytes):
    src = io.BytesIO(path_or_bytes) if isinstance(path_or_bytes, bytes) else path_or_bytes
    return pikepdf.open(src)


def check(path, findings):
    with open(path, "rb") as fh:
        data = fh.read()
    try:
        pdf = open_pdf(data)
    except pikepdf.PasswordError:
        raise ValueError("password-protected PDF; open it with the password and save an unprotected copy")
    except pikepdf.PdfError as e:
        raise ValueError("not a readable PDF (%s)" % str(e).splitlines()[0][:120])
    with pdf:
        _check(pdf, data, findings)


def _check(pdf, data, findings):
    paths = []
    check_metadata(pdf, findings, paths)

    comments, authors, redact_marks, box_annots, attach = [], [], [], {}, []
    for num, page in enumerate(pdf.pages, 1):
        ctm = page_ctm(page.obj)
        for annot in page.obj.get("/Annots", []):
            if not isinstance(annot, pikepdf.Dictionary):
                continue
            st = str(annot.get("/Subtype", ""))
            if st == "/Redact":
                redact_marks.append((num, [apply(ctm, b) for b in quad_boxes(annot)]))
            elif st == "/FileAttachment":
                attach.append("page %d: %s" % (num, _filespec_name(annot.get("/FS", "")) or "attachment"))
            elif st in COMMENT_TYPES:
                fill = annot_rgb(annot.get("/IC")) if st != "/Highlight" else annot_rgb(annot.get("/C"))
                if st in BOX_TYPES and fill is not None and is_dark(fill):
                    box_annots.setdefault(num, []).extend(apply(ctm, b) for b in quad_boxes(annot))
                    continue
                who, text = _str(annot.get("/T", "")), _str(annot.get("/Contents", ""))
                if st == "/Highlight" and not text and not who:
                    continue
                authors.append(who)
                comments.append("page %d, %s%s: %s" % (num, st[1:].lower(), " by " + who if who else "",
                                                     quoted(text) if text else "(no text)"))
            elif st == "/Link":
                a = annot.get("/A")
                if isinstance(a, pikepdf.Dictionary) and a.get("/S") == "/URI":
                    uri = _str(a.get("/URI", ""))
                    if uri.lower().startswith("file:"):
                        paths.append((uri, "page %d link" % num))

    if comments:
        who = join_names(authors)
        findings.append(Finding("comments", plural(len(comments), "comment") + (" by " + who if who else ""),
                                comments[:10]))

    # content-stream boxes, dark annotation boxes and unapplied redactions all need the page text
    fake, cover, unapplied = [], [], []
    redact_pages = {}
    for n, boxes in redact_marks:
        redact_pages.setdefault(n, []).extend(boxes)
    try:
        for num, items in enumerate(layouts(data), 1):
            f, c = covered(items, box_annots.get(num, ()))
            fake.extend("page %d: black box over %s" % (num, secret(t[:200])) for _, t in f)
            cover.extend("page %d: %s" % (num, secret(t[:200])) for _, t in c)
            if num in redact_pages:
                t = text_under(items, redact_pages[num])
                unapplied.append("page %d: %s marked%s" % (
                    num, plural(len(redact_pages[num]), "area"),
                    ", text still underneath: " + secret(t[:200]) if t.strip() else ""))
    except Exception as e:  # pdfminer cannot parse every PDF that qpdf can open
        findings.append(Finding("unreadable", "page text could not be read for the redaction checks (%s)"
                                % type(e).__name__, severity="warn"))
    if fake:
        findings.append(Finding("fake-redaction", plural(len(fake), "black box", "black boxes")
                                + " drawn over text that can still be selected and copied", fake[:10],
                                secret=True))
    if redact_marks:
        findings.append(Finding("unapplied-redaction", "redaction marks on " + plural(len(redact_pages), "page")
                                + " were never applied", unapplied[:10] or
                                ["page %d" % n for n in sorted(redact_pages)], secret=True))
    if cover:
        findings.append(Finding("covered-text", plural(len(cover), "passage")
                                + " of text drawn over by a white box or an image", cover[:10], secret=True))

    try:
        files = dict(pdf.attachments)
    except Exception:
        files = {}
    embedded = ["%s (%s KB)" % (name, max(1, len(fs.get_file().read_bytes()) // 1024)) for name, fs in files.items()]
    embedded += attach
    # PDF 2.0 associated files can carry a file without listing it in the attachments panel
    for where, holder in [("document", pdf.Root)] + [("page %d" % n, p.obj) for n, p in enumerate(pdf.pages, 1)]:
        for fs in holder.get("/AF", []):
            if isinstance(fs, pikepdf.Dictionary) and "/EF" in fs:
                name = _filespec_name(fs) or "attachment"
                if name not in files:
                    embedded.append("%s (associated file, %s)" % (name, where))
    if embedded:
        findings.append(Finding("embedded-files", plural(len(embedded), "attached file") + ": "
                                + join_names([e.split(" (")[0] for e in embedded]), embedded[:10]))

    check_actions(pdf, findings, paths)

    oc = pdf.Root.get("/OCProperties")
    if oc is not None and "/D" in oc and len(oc.D.get("/OFF", [])):
        off = [_str(g.get("/Name", "?")) for g in oc.D.OFF if isinstance(g, pikepdf.Dictionary)]
        findings.append(Finding("hidden-layers", plural(len(off), "layer") + " switched off: " + join_names(off),
                                off[:10]))

    apps = []
    if "/PieceInfo" in pdf.Root:
        apps.extend(str(k)[1:] for k in pdf.Root.PieceInfo.keys())
    pages_with = [n for n, p in enumerate(pdf.pages, 1) if "/PieceInfo" in p.obj]
    for n in pages_with[:1]:
        apps.extend(str(k)[1:] for k in pdf.pages[n - 1].obj.PieceInfo.keys())
    if apps:
        findings.append(Finding("private-app-data", "private data from " + join_names(apps)
                                + (" on " + plural(len(pages_with), "page") if pages_with else ""),
                                ["Illustrator and similar apps can keep a full editable copy here"]
                                if any("Illustrator" in a for a in apps) else []))

    if paths:
        found = {}
        for p, w in paths:
            found.setdefault(p, w)
        findings.append(Finding("file-paths", plural(len(found), "file path") + " stored in the file, e.g. "
                                + quoted(next(iter(found)), 60), ["%s  (%s)" % kv for kv in found.items()]))

    check_revisions(pdf, data, findings, is_signed(pdf))
