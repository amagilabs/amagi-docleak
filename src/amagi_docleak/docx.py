"""Word (.docx/.docm) checks."""

from .model import Finding
from .ooxml import q, text_of
from .util import (UNREADABLE, contrast, hex_to_rgb, is_blocked_out, is_dark, join_names, plural, quoted,
                   secret)

HIGHLIGHT = {
    "black": "000000", "blue": "0000FF", "cyan": "00FFFF", "green": "00FF00", "magenta": "FF00FF",
    "red": "FF0000", "yellow": "FFFF00", "white": "FFFFFF", "darkBlue": "000080",
    "darkCyan": "008080", "darkGreen": "008000", "darkMagenta": "800080", "darkRed": "800000",
    "darkYellow": "808000", "darkGray": "808080", "lightGray": "C0C0C0",
}
WHITE = (1.0, 1.0, 1.0)
BLACK = (0.0, 0.0, 0.0)

REVISION_TAGS = ["w:ins", "w:del", "w:moveFrom", "w:moveTo"]
FORMAT_CHANGE_TAGS = ["w:rPrChange", "w:pPrChange", "w:sectPrChange", "w:tblPrChange",
                      "w:tblGridChange", "w:trPrChange", "w:tcPrChange", "w:tblPrExChange",
                      "w:numberingChange", "w:cellIns", "w:cellDel", "w:cellMerge"]


def story_parts(pkg, main):
    parts = [main]
    for r in pkg.rels(main):
        kind = r.type.rsplit("/", 1)[-1]
        if not r.external and kind in ("header", "footer", "footnotes", "endnotes") and pkg.has(r.part):
            parts.append(r.part)
    return parts


def _toggle(el):
    if el is None:
        return None
    return el.get(q("w:val"), "true").lower() not in ("0", "false", "off")


def read_rpr(rpr):
    """The run properties this tool cares about, from one w:rPr element."""
    props = {}
    if rpr is None:
        return props
    v = _toggle(rpr.find(q("w:vanish")))
    if v is not None:
        props["vanish"] = v
    el = rpr.find(q("w:color"))
    if el is not None and el.get(q("w:val")):
        props["color"] = el.get(q("w:val"))
    el = rpr.find(q("w:highlight"))
    if el is not None and el.get(q("w:val")):
        props["highlight"] = el.get(q("w:val"))
    el = rpr.find(q("w:shd"))
    if el is not None and el.get(q("w:fill")):
        props["shd"] = el.get(q("w:fill"))
    el = rpr.find(q("w:sz"))
    if el is not None:
        try:
            props["sz"] = int(el.get(q("w:val"), "0"))
        except ValueError:
            pass
    return props


class Styles:
    def __init__(self, root):
        self.by_id = {}
        self.default_para = None
        self.defaults = {}
        if root is None:
            return
        for st in root.iter(q("w:style")):
            sid = st.get(q("w:styleId"))
            self.by_id[sid] = st
            if st.get(q("w:type")) == "paragraph" and st.get(q("w:default")) in ("1", "true"):
                self.default_para = sid
        dflt = root.find("%s/%s/%s" % (q("w:docDefaults"), q("w:rPrDefault"), q("w:rPr")))
        self.defaults = read_rpr(dflt)

    def chain(self, sid, kind):
        """(rPr props, pPr shading) of a style and the styles it is based on, base first."""
        seen, stack = set(), []
        while sid and sid in self.by_id and sid not in seen:
            seen.add(sid)
            st = self.by_id[sid]
            stack.append(st)
            based = st.find(q("w:basedOn"))
            sid = based.get(q("w:val")) if based is not None else None
        props, pshd = {}, None
        for st in reversed(stack):
            props.update(read_rpr(st.find(q("w:rPr"))))
            if kind == "paragraph":
                shd = st.find("%s/%s" % (q("w:pPr"), q("w:shd")))
                if shd is not None and shd.get(q("w:fill")):
                    pshd = shd.get(q("w:fill"))
        return props, pshd


def _fill(value):
    if not value or value.lower() == "auto":
        return None
    return hex_to_rgb(value)


def run_text(r):
    parts = []
    for el in r:
        if el.tag == q("w:t"):
            parts.append(el.text or "")
        elif el.tag == q("w:tab"):
            parts.append("\t")
    return "".join(parts)


def _ancestor(el, tag):
    el = el.getparent()
    while el is not None:
        if el.tag == tag:
            return el
        el = el.getparent()
    return None


def _inside(el, tags):
    el = el.getparent()
    while el is not None:
        if el.tag in tags:
            return True
        el = el.getparent()
    return False


def classify_runs(root, styles, page_bg=WHITE):
    """Yield (run, kind, text) for runs a reader cannot see.

    kind is 'vanish', 'tiny', 'match' (text color matches a light background) or
    'redaction' (dark text on a dark highlight or shading)."""
    deleted = {q("w:del"), q("w:moveFrom")}
    for r in root.iter(q("w:r")):
        text = run_text(r)
        if not text.strip() or _inside(r, deleted):
            continue
        p = _ancestor(r, q("w:p"))
        ppr = p.find(q("w:pPr")) if p is not None else None
        pstyle = None
        if ppr is not None and ppr.find(q("w:pStyle")) is not None:
            pstyle = ppr.find(q("w:pStyle")).get(q("w:val"))
        props = dict(styles.defaults)
        sprops, pshd = styles.chain(pstyle or styles.default_para, "paragraph")
        props.update(sprops)
        rpr = r.find(q("w:rPr"))
        if rpr is not None and rpr.find(q("w:rStyle")) is not None:
            props.update(styles.chain(rpr.find(q("w:rStyle")).get(q("w:val")), "character")[0])
        props.update(read_rpr(rpr))
        if ppr is not None and ppr.find(q("w:shd")) is not None:
            pshd = ppr.find(q("w:shd")).get(q("w:fill")) or pshd
        tc = _ancestor(r, q("w:tc"))
        cshd = None
        if tc is not None:
            shd = tc.find("%s/%s" % (q("w:tcPr"), q("w:shd")))
            if shd is not None:
                cshd = shd.get(q("w:fill"))
        if props.get("vanish"):
            yield r, "vanish", text
            continue
        if 0 < props.get("sz", 99) <= 2:
            yield r, "tiny", text
            continue
        bg = None
        if props.get("highlight") and props["highlight"] != "none":
            bg = hex_to_rgb(HIGHLIGHT.get(props["highlight"], ""))
        bg = bg or _fill(props.get("shd")) or _fill(pshd) or _fill(cshd) or page_bg
        color = props.get("color", "auto")
        if color.lower() == "auto":
            continue  # Word picks black or white automatically, so auto text is always readable
        fg = hex_to_rgb(color)
        if fg is None or contrast(fg, bg) >= UNREADABLE:
            continue
        yield r, ("redaction" if is_dark(bg) else "match"), text


def page_background(root):
    bg = root.find(q("w:background"))
    if bg is not None:
        rgb = _fill(bg.get(q("w:color")))
        if rgb:
            return rgb
    return WHITE


def _group(items):
    """Join consecutive runs of the same kind inside one paragraph into one excerpt."""
    groups = []
    last_p, last_kind = None, None
    for r, kind, text in items:
        p = _ancestor(r, q("w:p"))
        if groups and p is last_p and kind == last_kind:
            groups[-1][1].append(text)
        else:
            groups.append((kind, [text]))
        last_p, last_kind = p, kind
    return [(k, "".join(t)) for k, t in groups]


def dark_shapes(root):
    """Names of drawn shapes with a solid black fill (outside mc:Fallback copies)."""
    names = []
    for sp in root.iter(q("wps:wsp")):
        if _inside(sp, {q("mc:Fallback")}):
            continue
        sppr = sp.find(q("wps:spPr"))
        fill = sppr.find(q("a:solidFill")) if sppr is not None else None
        if fill is None or not _drawing_fill_dark(fill):
            continue
        anchor = _ancestor(sp, q("wp:anchor")) or _ancestor(sp, q("wp:inline"))
        docpr = anchor.find(q("wp:docPr")) if anchor is not None else None
        names.append(docpr.get("name") if docpr is not None else "shape")
    for tag in ("v:rect", "v:shape", "v:roundrect"):
        for sp in root.iter(q(tag)):
            if _inside(sp, {q("mc:Fallback")}):
                continue
            color = (sp.get("fillcolor") or "").lower()
            if sp.get("filled", "t").lower() in ("f", "false"):
                continue
            if color in ("black", "#000000", "#000", "windowtext"):
                names.append(sp.get("id") or "shape")
    return names


def _drawing_fill_dark(fill):
    el = fill.find(q("a:srgbClr"))
    if el is not None:
        rgb = hex_to_rgb(el.get("val"))
        return bool(rgb and is_dark(rgb))
    el = fill.find(q("a:prstClr"))
    if el is not None:
        return el.get("val") == "black"
    el = fill.find(q("a:schemeClr"))
    if el is not None:
        return el.get("val") in ("tx1", "dk1")
    el = fill.find(q("a:sysClr"))
    if el is not None:
        rgb = hex_to_rgb(el.get("lastClr"))
        return bool(rgb and is_dark(rgb))
    return False


def check(pkg, findings):
    main = pkg.main_part()
    parts = story_parts(pkg, main)
    styles_part = next(iter(pkg.rel_parts(main, "/styles")), None)
    styles = Styles(pkg.xml(styles_part) if styles_part else None)
    settings_part = next(iter(pkg.rel_parts(main, "/settings")), None)
    settings = pkg.xml(settings_part) if settings_part else None

    # tracked changes
    counts, authors, deleted, inserted = {}, [], [], []
    for part in parts:
        root = pkg.xml(part)
        for tag in REVISION_TAGS + FORMAT_CHANGE_TAGS:
            for el in root.iter(q(tag)):
                parent = el.getparent().tag
                if tag in ("w:ins", "w:del") and parent in (q("w:rPr"), q("w:trPr")):
                    key = "paragraph-mark" if parent == q("w:rPr") else "table-row"
                else:
                    key = tag.split(":")[1]
                counts[key] = counts.get(key, 0) + 1
                authors.append(el.get(q("w:author"), ""))
                if tag in ("w:del", "w:moveFrom") and parent not in (q("w:rPr"), q("w:trPr")):
                    t = text_of(el, "w:delText") or text_of(el, "w:t")
                    if t.strip():
                        deleted.append(t)
                elif tag == "w:ins" and parent not in (q("w:rPr"), q("w:trPr")):
                    t = text_of(el, "w:t")
                    if t.strip():
                        inserted.append(t)
    if counts:
        total = sum(v for k, v in counts.items() if k not in ("paragraph-mark", "table-row"))
        total = total or sum(counts.values())
        bits = []
        for key, label in (("del", "deletion"), ("ins", "insertion"), ("moveFrom", "move")):
            if counts.get(key):
                bits.append(plural(counts[key], label))
        fmt = sum(counts.get(t.split(":")[1], 0) for t in FORMAT_CHANGE_TAGS)
        if fmt:
            bits.append(plural(fmt, "formatting change"))
        who = join_names(authors)
        summary = plural(total, "tracked change") + (" by " + who if who else "")
        if bits:
            summary += " (" + ", ".join(bits) + ")"
        details = ["deleted: " + quoted(t) for t in deleted[:5]]
        details += ["inserted: " + quoted(t) for t in inserted[:5]]
        if settings is not None and settings.find(q("w:trackRevisions")) is not None:
            details.append("Track Changes is still switched on")
        findings.append(Finding("tracked-changes", summary, details))

    # comments
    comment_parts = pkg.rel_parts(main, "/comments")
    comments = []
    for cp in comment_parts:
        for c in pkg.xml(cp).iter(q("w:comment")):
            comments.append((c.get(q("w:author"), ""), text_of(c, "w:t")))
    if comments:
        who = join_names([a for a, _ in comments])
        findings.append(Finding("comments", plural(len(comments), "comment") + (" by " + who if who else ""),
                                ["%s: %s" % (a or "?", quoted(t)) for a, t in comments[:8]]))

    # hidden text and fake redactions
    hidden, redacted = [], []
    for part in parts:
        root = pkg.xml(part)
        bg = page_background(root) if part == main else WHITE
        for kind, text in _group(list(classify_runs(root, styles, bg))):
            if is_blocked_out(text):
                continue
            (redacted if kind == "redaction" else hidden).append((kind, text))
    if redacted:
        findings.append(Finding("fake-redaction",
                                plural(len(redacted), "passage") + " blacked out with highlight or shading; "
                                "the text is still there and can be copied",
                                ["blacked out: " + secret(t) for _, t in redacted[:10]], secret=True))
    if hidden:
        label = {"vanish": "hidden", "tiny": "1pt", "match": "same color as its background"}
        findings.append(Finding("hidden-text", plural(len(hidden), "passage") + " of text the reader cannot see ("
                                + join_names([label[k] for k, _ in hidden]) + ")",
                                ["%s: %s" % (label[k], quoted(t)) for k, t in hidden[:10]]))

    shapes = []
    for part in parts:
        shapes.extend(dark_shapes(pkg.xml(part)))
    if shapes:
        findings.append(Finding("possible-redaction-box",
                                plural(len(shapes), "black shape") + " drawn in the document; any text under "
                                "them can still be selected", ["shape: " + s for s in shapes[:10]]))

    # document variables
    if settings is not None:
        dv = settings.find(q("w:docVars"))
        if dv is not None and len(dv):
            vars_ = [(v.get(q("w:name"), "?"), v.get(q("w:val"), "")) for v in dv.iter(q("w:docVar"))]
            findings.append(Finding("document-variables", plural(len(vars_), "document variable") + ": "
                                    + join_names([n for n, _ in vars_]),
                                    ["%s = %s" % (n, quoted(v, 60)) for n, v in vars_[:10]]))
