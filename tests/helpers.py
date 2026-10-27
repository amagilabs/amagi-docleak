"""Small builders for test files, so each test states exactly what it plants."""

import zipfile
from pathlib import Path

import pikepdf

EXAMPLES = Path(__file__).resolve().parent.parent / "examples"


def check_ids(report):
    return {f.check for f in report.findings}


def zip_text(path):
    """Every part of an Office file, decoded, joined into one string."""
    with zipfile.ZipFile(path) as z:
        return "\n".join(z.read(n).decode("utf-8", "replace") for n in z.namelist())


def pdf_text(path):
    """Every string and every decodable stream in a PDF, as one string (what a determined reader could pull out)."""
    out = []
    with pikepdf.open(path) as pdf:
        for obj in pdf.objects:
            if isinstance(obj, pikepdf.Stream):
                try:
                    out.append(obj.read_bytes().decode("latin-1"))
                except pikepdf.PdfError:
                    out.append(obj.read_raw_bytes().decode("latin-1"))
            elif isinstance(obj, pikepdf.Dictionary):
                out.append(str(obj))
        if "/Info" in pdf.trailer:
            out.append(str(pdf.trailer.Info))
    return "\n".join(out)


def make_pdf(path, content, annots=(), info=None):
    """One-page Letter PDF with Helvetica as /F1 and the given content stream."""
    pdf = pikepdf.new()
    font = pikepdf.Dictionary(Type=pikepdf.Name.Font, Subtype=pikepdf.Name.Type1,
                              BaseFont=pikepdf.Name.Helvetica)
    page = pikepdf.Dictionary(
        Type=pikepdf.Name.Page,
        MediaBox=pikepdf.Array([0, 0, 612, 792]),
        Resources=pikepdf.Dictionary(Font=pikepdf.Dictionary(F1=font)),
        Contents=pdf.make_stream(content.encode("latin-1")),
    )
    pdf.pages.append(pikepdf.Page(page))
    if annots:
        pdf.pages[0].obj.Annots = pikepdf.Array([pdf.make_indirect(pikepdf.Dictionary(**a)) for a in annots])
    if info:
        for k, v in info.items():
            pdf.docinfo[k] = v
    pdf.save(path)
    return path


def append_update(path, content):
    """Save new page content the way editors do on a quick save: appended after the old file, old bytes kept."""
    import re
    base = Path(path).read_bytes()
    with pikepdf.open(path) as pdf:
        num, gen = pdf.pages[0].obj.Contents.objgen
        root = pdf.trailer.Root.objgen
        size = int(pdf.trailer.Size)
    prev = int(re.findall(rb"startxref\s+(\d+)", base)[-1])
    out = bytearray(base if base.endswith(b"\n") else base + b"\n")
    data = content.encode("latin-1")
    offset = len(out)
    out += b"%d %d obj\n<< /Length %d >>\nstream\n" % (num, gen, len(data)) + data + b"\nendstream\nendobj\n"
    xref = len(out)
    out += b"xref\n0 1\n0000000000 65535 f \n%d 1\n%010d %05d n \n" % (num, offset, gen)
    out += b"trailer\n<< /Size %d /Root %d %d R /Prev %d >>\nstartxref\n%d\n%%%%EOF\n" % (
        size, root[0], root[1], prev, xref)
    Path(path).write_bytes(bytes(out))
    return path


def text_line(x, y, text, color="0 g", mode=0):
    return "%s BT /F1 12 Tf %d Tr %d %d Td (%s) Tj ET\n" % (color, mode, x, y, text)


def box(x, y, w, h, color="0 g"):
    return "%s %d %d %d %d re f\n" % (color, x, y, w, h)


def patch_zip(path, add):
    """Rewrite an Office zip with extra or replaced parts: add = {name: bytes or str}."""
    src = Path(path).read_bytes()
    import io
    with zipfile.ZipFile(io.BytesIO(src)) as zin, zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename not in add:
                zout.writestr(item, zin.read(item.filename))
        for name, data in add.items():
            zout.writestr(name, data.encode("utf-8") if isinstance(data, str) else data)
