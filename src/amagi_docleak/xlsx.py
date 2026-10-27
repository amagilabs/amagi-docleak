"""Excel (.xlsx/.xlsm) checks."""

import re

from .model import Finding
from .ooxml import q, text_of
from .util import (UNREADABLE, contrast, hex_to_rgb, is_blocked_out, is_dark, join_names, plural,
                   quoted, secret)

WHITE = (1.0, 1.0, 1.0)
BLACK = (0.0, 0.0, 0.0)
INDEXED = ["000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
           "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
           "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080"]
HIDDEN_FORMAT = re.compile(r'^(?:;|"")*;{2,}(?:"")*$')


def col_index(letters):
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def col_letters(n):
    out = ""
    while n:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


def split_ref(ref):
    m = re.match(r"^([A-Z]+)(\d+)$", ref or "")
    return (col_index(m.group(1)), int(m.group(2))) if m else (0, 0)


def ranges(numbers):
    numbers = sorted(set(numbers))
    out, start, prev = [], None, None
    for n in numbers:
        if start is None:
            start = prev = n
        elif n == prev + 1:
            prev = n
        else:
            out.append((start, prev))
            start = prev = n
    if start is not None:
        out.append((start, prev))
    return out


class Workbook:
    """The parts of a workbook the checks and the cleaner need."""

    def __init__(self, pkg):
        self.pkg = pkg
        self.part = pkg.main_part()
        self.root = pkg.xml(self.part)
        rels = {r.id: r for r in pkg.rels(self.part)}
        self.sheets = []  # (name, state, part, sheet element)
        for sh in self.root.iter(q("x:sheet")):
            rel = rels.get(sh.get(q("r:id")))
            part = rel.part if rel is not None and not rel.external else None
            self.sheets.append((sh.get("name", "?"), sh.get("state", "visible"), part, sh))
        self.strings = []
        ss = next(iter(pkg.rel_parts(self.part, "/sharedStrings")), None)
        if ss:
            self.strings = [text_of(si, "x:t") for si in pkg.xml(ss).iter(q("x:si"))]
        self.theme = self._theme()
        self.xfs, self.fonts, self.fills, self.formats = [], [], [], {}
        st = next(iter(pkg.rel_parts(self.part, "/styles")), None)
        if st:
            self._styles(pkg.xml(st))

    def _theme(self):
        part = next(iter(self.pkg.rel_parts(self.part, "/theme")), None)
        colors = []
        root = self.pkg.xml(part) if part else None
        if root is not None:
            scheme = next(root.iter(q("a:clrScheme")), None)
            if scheme is not None:
                for el in scheme:
                    c = el.find(q("a:srgbClr"))
                    val = c.get("val") if c is not None else None
                    if val is None and el.find(q("a:sysClr")) is not None:
                        val = el.find(q("a:sysClr")).get("lastClr")
                    colors.append(val)
        if len(colors) >= 2:  # Excel numbers lt1 as 0 and dk1 as 1
            colors[0], colors[1] = colors[1], colors[0]
        if len(colors) >= 4:
            colors[2], colors[3] = colors[3], colors[2]
        return colors

    def color(self, el):
        if el is None:
            return None
        if el.get("rgb"):
            rgb = hex_to_rgb(el.get("rgb"))
        elif el.get("theme") is not None:
            i = int(el.get("theme"))
            rgb = hex_to_rgb(self.theme[i]) if i < len(self.theme) and self.theme[i] else None
        elif el.get("indexed") is not None:
            i = int(el.get("indexed"))
            rgb = hex_to_rgb(INDEXED[i]) if i < len(INDEXED) else (BLACK if i == 64 else None)
        else:
            return None
        tint = float(el.get("tint", "0") or 0)
        if rgb and tint:
            rgb = tuple(c + (1 - c) * tint if tint > 0 else c * (1 + tint) for c in rgb)
        return rgb

    def _styles(self, root):
        nf = root.find(q("x:numFmts"))
        if nf is not None:
            for el in nf:
                self.formats[el.get("numFmtId")] = el.get("formatCode", "")
        fonts = root.find(q("x:fonts"))
        for f in (fonts if fonts is not None else []):
            c = f.find(q("x:color"))
            self.fonts.append(self.color(c) or BLACK)
        fills = root.find(q("x:fills"))
        for f in (fills if fills is not None else []):
            pf = f.find(q("x:patternFill"))
            if pf is not None and pf.get("patternType") == "solid":
                # Excel draws a solid fill with no fgColor as no fill at all
                self.fills.append(self.color(pf.find(q("x:fgColor"))) or WHITE)
            else:
                self.fills.append(None)
        xfs = root.find(q("x:cellXfs"))
        for xf in (xfs if xfs is not None else []):
            self.xfs.append((xf.get("numFmtId", "0"), int(xf.get("fontId", "0") or 0), int(xf.get("fillId", "0") or 0)))

    def value(self, c):
        t = c.get("t")
        if t == "inlineStr":
            return text_of(c, "x:t")
        v = c.find(q("x:v"))
        if v is None or v.text is None:
            return ""
        if t == "s":
            try:
                return self.strings[int(v.text)]
            except (ValueError, IndexError):
                return ""
        return v.text

    def cell_style(self, c):
        """('hidden-format' | 'redaction' | 'match' | None) for one cell."""
        try:
            fmt, font, fill = self.xfs[int(c.get("s", "0"))]
        except (ValueError, IndexError):
            return None
        if HIDDEN_FORMAT.match(self.formats.get(fmt, "").replace(" ", "")):
            return "hidden-format"
        bg = self.fills[fill] if fill < len(self.fills) else None
        fg = self.fonts[font] if font < len(self.fonts) else BLACK
        bg = bg or WHITE
        if contrast(fg, bg) < UNREADABLE:
            return "redaction" if is_dark(bg) else "match"
        return None


def _cols(ws):
    out = []
    cols = ws.find(q("x:cols"))
    for col in (cols if cols is not None else []):
        if col.get("hidden") in ("1", "true"):
            out.append((int(col.get("min", "0")), int(col.get("max", "0"))))
    return out


def check(pkg, findings):
    wb = Workbook(pkg)

    hidden_sheets = [(n, s) for n, s, _, _ in wb.sheets if s in ("hidden", "veryHidden")]
    if hidden_sheets:
        names = ["%s%s" % (quoted(n, 40), " (very hidden)" if s == "veryHidden" else "") for n, s in hidden_sheets]
        details = []
        for name, state, part, _ in wb.sheets:
            if state in ("hidden", "veryHidden") and part and pkg.has(part):
                cells = [c for c in pkg.xml(part).iter(q("x:c")) if wb.value(c) != ""]
                details.append("%s: %s with values%s" % (name, plural(len(cells), "cell"),
                               "; very hidden sheets do not show up in Excel's Unhide list" if state == "veryHidden" else ""))
        findings.append(Finding("hidden-sheets", plural(len(hidden_sheets), "hidden sheet") + ": " + ", ".join(names), details))

    rowcol, hidden_cells, redacted, comments = [], [], [], []
    persons = {}
    for pp in pkg.rel_parts(wb.part, "/person"):
        for p in pkg.xml(pp).iter(q("tc:person")):
            persons[p.get("id")] = p.get("displayName", "")
    for name, state, part, _ in wb.sheets:
        if not part or not pkg.has(part):
            continue
        ws = pkg.xml(part)
        hidden_cols = _cols(ws)
        rows_with_data, cols_with_data, n_cells = set(), set(), 0
        for row in ws.iter(q("x:row")):
            row_hidden = row.get("hidden") in ("1", "true")
            for c in row.iter(q("x:c")):
                val = wb.value(c)
                if val == "":
                    continue
                col, rown = split_ref(c.get("r"))
                in_hidden_col = any(lo <= col <= hi for lo, hi in hidden_cols)
                if row_hidden:
                    rows_with_data.add(int(row.get("r", rown) or rown))
                if in_hidden_col:
                    cols_with_data.add(col)
                if row_hidden or in_hidden_col:
                    n_cells += 1
                kind = wb.cell_style(c)
                if kind and not is_blocked_out(val):
                    ref = "%s!%s" % (name, c.get("r"))
                    if kind == "redaction":
                        redacted.append("%s: %s" % (ref, secret(val)))
                    else:
                        why = "number format ;;;" if kind == "hidden-format" else "font matches fill"
                        hidden_cells.append("%s (%s): %s" % (ref, why, quoted(val, 50)))
        if rows_with_data or cols_with_data:
            bits = []
            if rows_with_data:
                bits.append("rows " + ", ".join(str(a) if a == b else "%d-%d" % (a, b) for a, b in ranges(rows_with_data)))
            if cols_with_data:
                bits.append("columns " + ", ".join(col_letters(a) if a == b else "%s-%s" % (col_letters(a), col_letters(b))
                                                   for a, b in ranges(cols_with_data)))
            rowcol.append("%s: %s (%s)" % (name, " and ".join(bits), plural(n_cells, "cell") + " with values"))

        threaded_refs = set()
        for tp in pkg.rel_parts(part, "/threadedComment"):
            for tcm in pkg.xml(tp).iter(q("tc:threadedComment")):
                threaded_refs.add(tcm.get("ref"))
                who = persons.get(tcm.get("personId"), "")
                comments.append((who, "%s!%s" % (name, tcm.get("ref")), text_of(tcm, "tc:text")))
        for cp in pkg.rel_parts(part, "/comments"):
            root = pkg.xml(cp)
            authors = [a.text or "" for a in root.iter(q("x:author"))]
            for cm in root.iter(q("x:comment")):
                if cm.get("ref") in threaded_refs:
                    continue
                try:
                    who = authors[int(cm.get("authorId", "0"))]
                except (ValueError, IndexError):
                    who = ""
                comments.append((who, "%s!%s" % (name, cm.get("ref")), text_of(cm, "x:t")))

    if rowcol:
        findings.append(Finding("hidden-rows-columns", "hidden rows or columns hold data on "
                                + plural(len(rowcol), "sheet"), rowcol))
    if redacted:
        findings.append(Finding("fake-redaction", plural(len(redacted), "cell") + " blacked out with a black fill "
                                "and black text; the values are still in the file", redacted[:10], secret=True))
    if hidden_cells:
        findings.append(Finding("hidden-cells", plural(len(hidden_cells), "cell") + " with values formatted to be invisible",
                                hidden_cells[:10]))
    if comments:
        who = join_names([w for w, _, _ in comments])
        findings.append(Finding("comments", plural(len(comments), "comment") + (" by " + who if who else ""),
                                ["%s %s: %s" % (ref, w or "?", quoted(t)) for w, ref, t in comments[:8]]))

    # pivot caches keep a copy of their source rows
    sheet_names = {n for n, _, _, _ in wb.sheets}
    pivots = []
    for cd in pkg.rel_parts(wb.part, "/pivotCacheDefinition"):
        root = pkg.xml(cd)
        records = pkg.rel_parts(cd, "/pivotCacheRecords")
        if not records:
            continue
        rec = pkg.xml(records[0])
        count = rec.get("count") or str(sum(1 for _ in rec.iter(q("x:r"))))
        src = next(root.iter(q("x:worksheetSource")), None)
        where = ""
        if src is not None:
            sheet = src.get("sheet") or src.get("name") or ""
            where = " from %s%s" % (quoted(sheet, 40), " %s" % src.get("ref") if src.get("ref") else "")
            if src.get("sheet") and src.get("sheet") not in sheet_names and not src.get(q("r:id")):
                where += " (that sheet is no longer in the workbook, but its rows are)"
        pivots.append("%s source rows cached%s" % (count, where))
    if pivots:
        findings.append(Finding("pivot-cache", plural(len(pivots), "pivot cache") + " keep a copy of their source data",
                                pivots))

    conns = []
    for cp in pkg.rel_parts(wb.part, "/connections"):
        for c in pkg.xml(cp).iter(q("x:connection")):
            target = ""
            for child in c:
                target = child.get("connection") or child.get("url") or child.get("sourceFile") or target
            target = re.sub(r"(?i)(password|pwd)=[^;\"]*", r"\1=***", target)
            conns.append("%s: %s" % (c.get("name", "?"), quoted(target, 90)))
    if conns:
        findings.append(Finding("data-connections", plural(len(conns), "data connection"), conns))
