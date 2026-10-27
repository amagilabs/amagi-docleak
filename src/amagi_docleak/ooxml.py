"""Read Office Open XML packages (.docx, .xlsx, .pptx) and run the checks they share."""

import posixpath
import re
import zipfile
from collections import namedtuple
from urllib.parse import unquote

from lxml import etree

from .model import Finding
from .util import find_paths, is_local_target, join_names, plural, quoted

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
    "ep": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
    "cup": "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "x15ac": "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac",
    "p188": "http://schemas.microsoft.com/office/powerpoint/2018/8/main",
    "tc": "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments",
    "v": "urn:schemas-microsoft-com:vml",
}

PARSER = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=True)
MAX_PART = 512 * 1024 * 1024

Rel = namedtuple("Rel", "id type target external part")


def q(tag):
    """'w:p' -> '{namespace}p'."""
    prefix, local = tag.split(":")
    return "{%s}%s" % (NS[prefix], local)


def rels_path(part):
    if not part:
        return "_rels/.rels"
    d, b = posixpath.split(part)
    return posixpath.join(d, "_rels", b + ".rels")


def resolve(part, target):
    target = unquote(target)
    if target.startswith("/"):
        return target[1:]
    return posixpath.normpath(posixpath.join(posixpath.dirname(part), target))


class NotOffice(Exception):
    pass


class Package:
    def __init__(self, path):
        self.path = path
        try:
            self.zf = zipfile.ZipFile(path)
        except zipfile.BadZipFile:
            raise NotOffice("not a zip package (password-protected Office files are not zips)")
        self.infos = {i.filename: i for i in self.zf.infolist()}
        if "[Content_Types].xml" not in self.infos:
            raise NotOffice("no [Content_Types].xml, so this is not an Office file")
        self._xml = {}

    def close(self):
        self.zf.close()

    def has(self, name):
        return name in self.infos

    def names(self):
        return list(self.infos)

    def read(self, name):
        info = self.infos[name]
        if info.file_size > MAX_PART:
            raise NotOffice("part %s is larger than %d MB" % (name, MAX_PART // 2 ** 20))
        return self.zf.read(info)

    def xml(self, name):
        if name not in self._xml:
            self._xml[name] = etree.fromstring(self.read(name), PARSER) if self.has(name) else None
        return self._xml[name]

    def rels(self, part):
        root = self.xml(rels_path(part))
        out = []
        if root is None:
            return out
        for el in root.iter(q("pr:Relationship")):
            target = el.get("Target", "")
            external = el.get("TargetMode") == "External"
            out.append(Rel(el.get("Id"), el.get("Type", ""), target, external,
                           None if external else resolve(part, target)))
        return out

    def rel_parts(self, part, type_suffix):
        return [r.part for r in self.rels(part)
                if not r.external and r.type.endswith(type_suffix) and self.has(r.part)]

    def main_part(self):
        for r in self.rels(""):
            if r.type.endswith("/officeDocument") and not r.external:
                return r.part
        return None


def text_of(el, tag="a:t"):
    return "".join(t.text or "" for t in el.iter(q(tag)))


# ---- shared checks ------------------------------------------------------------------------

CORE_FIELDS = [
    ("dc:creator", "author"), ("cp:lastModifiedBy", "last modified by"),
    ("dc:title", "title"), ("dc:subject", "subject"), ("dc:description", "comments field"),
    ("cp:keywords", "keywords"), ("cp:category", "category"), ("cp:contentStatus", "status"),
    ("cp:lastPrinted", "last printed"), ("cp:revision", "revision"),
]
APP_FIELDS = [("ep:Company", "company"), ("ep:Manager", "manager"),
              ("ep:HyperlinkBase", "hyperlink base"), ("ep:Template", "template")]
DEFAULT_TEMPLATES = {"", "normal", "normal.dotm", "normal.dot", "normal.dotx"}


def property_values(pkg):
    """Personal properties as (label, value) pairs, in display order."""
    out = []
    core = pkg.xml("docProps/core.xml")
    if core is not None:
        for tag, label in CORE_FIELDS:
            el = core.find(q(tag))
            value = (el.text or "").strip() if el is not None else ""
            if not value or (tag == "cp:revision" and value == "1"):
                continue
            out.append((label, value))
    app = pkg.xml("docProps/app.xml")
    if app is not None:
        for tag, label in APP_FIELDS:
            el = app.find(q(tag))
            value = (el.text or "").strip() if el is not None else ""
            if not value or (tag == "ep:Template" and value.lower() in DEFAULT_TEMPLATES):
                continue
            out.append((label, value))
        el = app.find(q("ep:TotalTime"))
        if el is not None and (el.text or "0").strip() not in ("", "0"):
            out.append(("total editing time", "%s minutes" % el.text.strip()))
    return out


def check_properties(pkg, findings):
    values = property_values(pkg)
    if values:
        lead = [p for p in values if p[0] in ("author", "last modified by", "company", "manager")]
        lead = lead or values[:2]
        summary = ", ".join("%s %s" % (k, quoted(v, 40)) for k, v in lead[:3])
        findings.append(Finding("author-metadata", summary,
                                ["%s: %s" % (k, v) for k, v in values]))
    custom = pkg.xml("docProps/custom.xml")
    if custom is not None:
        props = []
        for el in custom.iter(q("cup:property")):
            value = " ".join(t.strip() for t in el.itertext() if t.strip())
            props.append("%s: %s" % (el.get("name", "?"), value))
        if props:
            findings.append(Finding("custom-properties", plural(len(props), "custom property", "custom properties")
                                    + " (" + join_names([p.split(":")[0] for p in props]) + ")", props))
    thumbs = [n for n in pkg.names() if n.startswith("docProps/thumbnail")]
    if thumbs:
        findings.append(Finding("thumbnail", "preview image of the first page saved in the file (%s)" % thumbs[0]))


def check_embedded(pkg, findings):
    embedded = [n for n in pkg.names() if "/embeddings/" in n and not n.endswith("/")]
    if embedded:
        details = ["%s (%s KB)" % (n, max(1, pkg.infos[n].file_size // 1024)) for n in embedded]
        findings.append(Finding("embedded-files", plural(len(embedded), "embedded file") + ": "
                                + join_names([posixpath.basename(n) for n in embedded]), details))
    macros = [n for n in pkg.names() if n.lower().endswith("vbaproject.bin")]
    if macros:
        findings.append(Finding("macros", "VBA macro project inside the file (%s)" % macros[0]))


# parts whose text is what the reader sees; only attributes, field codes and formulas are
# searched for paths there, so a path typed in the body is not reported as hidden
CONTENT_PART = re.compile(r"^(word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml"
                          r"|ppt/slides/slide\d+\.xml|xl/sharedStrings\.xml|xl/worksheets/sheet\d+\.xml)$")
HIDDEN_TEXT_TAGS = {q("w:instrText"), q("x:f")}


def check_paths(pkg, findings, extra=()):
    found = {}
    for name in pkg.names():
        if name.endswith(".rels"):
            root = pkg.xml(name)
            for el in root.iter(q("pr:Relationship")):
                if el.get("TargetMode") == "External" and is_local_target(el.get("Target", "")):
                    kind = el.get("Type", "").rsplit("/", 1)[-1]
                    found.setdefault(unquote(el.get("Target")), "%s link in %s" % (kind, name))
            continue
        if not name.endswith(".xml") or name.startswith("customXml/"):
            continue
        root = pkg.xml(name)
        if root is None:
            continue
        if CONTENT_PART.match(name):
            chunks = []
            for el in root.iter():
                chunks.extend(v for v in el.attrib.values())
                if el.tag in HIDDEN_TEXT_TAGS and el.text:
                    chunks.append(el.text)
        else:
            chunks = [t for t in root.itertext()] + [v for el in root.iter() for v in el.attrib.values()]
        for chunk in chunks:
            for path in find_paths(chunk):
                found.setdefault(path, name)
    for path, where in extra:
        found.setdefault(path, where)
    if found:
        details = ["%s  (%s)" % (p, w) for p, w in found.items()]
        findings.append(Finding("file-paths", plural(len(found), "file path") + " stored in the file, e.g. "
                                + quoted(next(iter(found)), 60), details))


def check_common(pkg, findings):
    check_properties(pkg, findings)
    check_embedded(pkg, findings)
    check_paths(pkg, findings)
