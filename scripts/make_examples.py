"""Build the example files in examples/.

Every name, firm, case and number in them is made up. Document properties are set explicitly
so nothing from the machine that runs this script (Office user name, company) ends up inside.

    pip install -e ".[dev]" reportlab
    python scripts/make_examples.py            # Word/PDF export via Word when available (Windows)
    python scripts/make_examples.py --no-word  # reportlab only
"""

import argparse
import datetime as dt
import io
import os
import re
import shutil
import sys
import tempfile
import zipfile

import docx
import openpyxl
import pikepdf
import pptx
from docx.enum.text import WD_COLOR_INDEX
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Pt, RGBColor
from openpyxl.comments import Comment
from openpyxl.styles import Font, PatternFill
from pptx.dml.color import RGBColor as PRGB
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt as PPt
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(os.path.dirname(HERE), "examples")
WHEN = dt.datetime(2026, 9, 14, 16, 5, 0)
FIRM = "Harlow & Pike LLP"
AUTHOR = "Jordan Ellery"
EDITOR = "Sam Okafor"
CASE = "Brightwater Supply v. Kestrel Freight"


def patch_zip(path, replace=None, add=None, content_types=(), rels=None):
    """Rewrite an OOXML zip: replace or add parts, register content types and relationships.

    rels: {rels_part: [xml_snippet, ...]} appended inside <Relationships>.
    """
    replace, add, rels = dict(replace or {}), dict(add or {}), dict(rels or {})
    with zipfile.ZipFile(path) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    parts.update(replace)
    parts.update(add)
    ct = parts["[Content_Types].xml"].decode("utf-8")
    for name, ctype in content_types:
        ct = ct.replace("</Types>", '<Override PartName="/%s" ContentType="%s"/></Types>' % (name, ctype))
    parts["[Content_Types].xml"] = ct.encode("utf-8")
    for rp, snippets in rels.items():
        if rp in parts:
            x = parts[rp].decode("utf-8")
        else:
            x = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>')
        x = x.replace("</Relationships>", "".join(snippets) + "</Relationships>")
        parts[rp] = x.encode("utf-8")
    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        order = ["[Content_Types].xml"] + [n for n in parts if n != "[Content_Types].xml"]
        for n in order:
            info = zipfile.ZipInfo(n, (2026, 9, 14, 16, 5, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, parts[n])
    os.replace(tmp, path)


def app_xml(application, extra):
    body = "".join("<%s>%s</%s>" % (k, v, k) for k, v in extra)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
            'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            "<Application>%s</Application>%s</Properties>" % (application, body)).encode("utf-8")


# The standard FMTID for user-defined document properties (split so secret scanners skip it).
FMTID = "{D5CDD505-2E9C-101B" "-9397-08002B2CF9AE}"
CUSTOM_XML = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
              'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
              '<property fmtid="%s" pid="2" name="ClientMatter">'
              "<vt:lpwstr>BRW-0412</vt:lpwstr></property>"
              '<property fmtid="%s" pid="3" name="DMSDocID">'
              "<vt:lpwstr>HPDMS-88213 v7</vt:lpwstr></property></Properties>" % (FMTID, FMTID)).encode("utf-8")
CUSTOM_REL = ('<Relationship Id="rIdCustom" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
              'relationships/custom-properties" Target="docProps/custom.xml"/>')
CUSTOM_CT = ("docProps/custom.xml", "application/vnd.openxmlformats-officedocument.custom-properties+xml")


# ---- Word ---------------------------------------------------------------------------------

def _w(xml):
    return parse_xml(xml.replace(" ", " %s " % nsdecls("w"), 1))


def make_docx(path):
    d = docx.Document()
    cp = d.core_properties
    cp.author, cp.last_modified_by = AUTHOR, EDITOR
    cp.title = CASE + " - settlement terms"
    cp.subject = cp.keywords = cp.category = ""
    cp.comments = ""
    cp.revision = 7
    cp.created = cp.modified = WHEN
    cp.last_printed = WHEN

    d.add_heading(CASE, 1)
    d.add_paragraph("Settlement terms (draft for discussion)")

    p = d.add_paragraph("1. Kestrel Freight will pay Brightwater Supply ")
    p._p.append(_w('<w:del w:id="101" w:author="%s" w:date="2026-09-12T10:02:00Z"><w:r>'
                   '<w:delText xml:space="preserve">$410,000</w:delText></w:r></w:del>' % EDITOR))
    p._p.append(_w('<w:ins w:id="102" w:author="%s" w:date="2026-09-12T10:02:00Z"><w:r>'
                   '<w:t>$385,000</w:t></w:r></w:ins>' % EDITOR))
    p.add_run(" within 30 days of signing.")

    p = d.add_paragraph("2. Each side bears its own fees. ")
    p._p.append(_w('<w:del w:id="103" w:author="%s" w:date="2026-09-12T10:05:00Z"><w:r>'
                   '<w:delText xml:space="preserve">Our client will not go below $350,000.</w:delText>'
                   "</w:r></w:del>" % AUTHOR))

    p = d.add_paragraph("3. Kestrel confirms its carrier, ")
    r = p.add_run("Halvard Mutual, policy KF-2291")
    r.font.highlight_color = WD_COLOR_INDEX.BLACK
    r.font.color.rgb = RGBColor(0, 0, 0)
    p.add_run(", will fund the payment.")

    p = d.add_paragraph("4. Mutual release of all claims arising from the 2025 delivery contract.")
    r = p.add_run(" Internal: Sam thinks the arbitrator leans our way.")
    r.font.hidden = True

    p = d.add_paragraph("5. Confidentiality applies to these terms.")
    r = p.add_run(" Draft 7 - do not circulate outside the firm.")
    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    r.font.size = Pt(11)

    p = d.add_paragraph("See the ")
    rid = d.part.relate_to(r"D:\Work\jellery\Brightwater\call-notes-0911.docx", RT.HYPERLINK, is_external=True)
    p._p.append(parse_xml('<w:hyperlink %s r:id="%s"><w:r><w:t>call notes</w:t></w:r></w:hyperlink>'
                          % (nsdecls("w", "r"), rid)))
    p.add_run(" for the figures.")

    target = d.paragraphs[4].runs[0]
    d.add_comment(target, text="Check the carrier name with Jordan before this goes out.", author=EDITOR,
                  initials="SO")

    settings = d.part.part_related_by(RT.SETTINGS)
    st = settings.element
    rid = settings.relate_to(r"file:///\\fs01\templates\HarlowPike-Settlement.dotx",
                             "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
                             "attachedTemplate", is_external=True)
    proof = st.find("{%s}proofState" % st.nsmap["w"])
    proof.addnext(parse_xml('<w:attachedTemplate %s r:id="%s"/>' % (nsdecls("w", "r"), rid)))
    st.find("{%s}attachedTemplate" % st.nsmap["w"]).addnext(parse_xml("<w:trackRevisions %s/>" % nsdecls("w")))
    st.find("{%s}compat" % st.nsmap["w"]).addnext(parse_xml(
        '<w:docVars %s><w:docVar w:name="MatterNo" w:val="BRW-0412"/>'
        '<w:docVar w:name="ReviewedBy" w:val="%s"/></w:docVars>' % (nsdecls("w"), EDITOR)))
    d.save(path)
    patch_zip(path,
              replace={"docProps/app.xml": app_xml("Microsoft Office Word", [
                  ("Template", "HarlowPike-Settlement.dotx"), ("TotalTime", "412"), ("Pages", "1"),
                  ("Company", FIRM.replace("&", "&amp;")), ("Manager", "R. Pike"), ("AppVersion", "16.0000")])},
              add={"docProps/custom.xml": CUSTOM_XML}, content_types=[CUSTOM_CT],
              rels={"_rels/.rels": [CUSTOM_REL]})


# ---- Excel --------------------------------------------------------------------------------

def make_xlsx(path):
    wb = openpyxl.Workbook()
    pr = wb.properties
    pr.creator, pr.lastModifiedBy, pr.title = AUTHOR, EDITOR, "Brightwater damages model"
    pr.created = pr.modified = WHEN
    ws = wb.active
    ws.title = "Summary"
    rows = [("Item", "Amount", "Note", "Internal margin"),
            ("Lost shipments", 212000, "per delivery log", 0.31),
            ("Spoiled stock", 96500, "invoices 114-131", 0.22),
            ("Expedited freight", 41800, "", 0.18),
            ("Total claimed", "=SUM(B2:B4)", "", ""),
            ("Insurer reserve", "Halvard reserve $1.2M", "", ""),
            ("Walk-away", 0.82, "share of claim", "")]
    for r in rows:
        ws.append(r)
    ws.column_dimensions["D"].hidden = True
    ws["B6"].fill = PatternFill("solid", fgColor="FF000000")
    ws["B6"].font = Font(color="FF000000")
    ws["B7"].number_format = ";;;"
    ws.append(())
    ws.append(("Old demand", 450000, "superseded 2026-08-30", ""))
    ws.row_dimensions[9].hidden = True
    ws["B2"].comment = Comment("Padded about 15% for negotiation. Keep this between us.", AUTHOR)

    hid = wb.create_sheet("Fee split")
    hid.append(("Partner", "Share"))
    hid.append(("Harlow", 0.6))
    hid.append(("Pike", 0.4))
    hid.sheet_state = "hidden"
    vh = wb.create_sheet("Notes")
    vh.append(("Kestrel CFO is leaving in Q1; push for a quick close.",))
    vh.sheet_state = "veryHidden"
    wb.save(path)

    with zipfile.ZipFile(path) as z:
        book = z.read("xl/workbook.xml").decode("utf-8")
    alt = ('<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
           '<mc:Choice Requires="x15" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">'
           '<x15ac:absPath xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac" '
           'url="D:\\Work\\jellery\\Brightwater\\"/></mc:Choice></mc:AlternateContent>')
    book = book.replace("<workbookPr/>", "<workbookPr/>" + alt, 1)
    conn = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<connection id="1" name="Billing" type="1" refreshedVersion="6" background="1" saveData="1">'
            '<dbPr connection="DSN=HPBilling;UID=reports;PWD=Winter2026;DATABASE=billing" '
            'command="SELECT * FROM matters"/></connection></connections>')
    patch_zip(path,
              replace={"xl/workbook.xml": book.encode("utf-8"),
                       "docProps/app.xml": app_xml("Microsoft Excel", [
                           ("Company", FIRM.replace("&", "&amp;")), ("AppVersion", "16.0300")])},
              add={"xl/connections.xml": conn.encode("utf-8"), "docProps/custom.xml": CUSTOM_XML},
              content_types=[("xl/connections.xml",
                              "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml"),
                             CUSTOM_CT],
              rels={"xl/_rels/workbook.xml.rels": [
                  '<Relationship Id="rIdConn" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                  'relationships/connections" Target="connections.xml"/>'],
                  "_rels/.rels": [CUSTOM_REL]})


# ---- PowerPoint ---------------------------------------------------------------------------

def make_pptx(path):
    prs = pptx.Presentation()
    cp = prs.core_properties
    cp.author, cp.last_modified_by = AUTHOR, EDITOR
    cp.title, cp.comments, cp.revision = "Kestrel settlement strategy", "", 4
    cp.created = cp.modified = WHEN

    s1 = prs.slides.add_slide(prs.slide_layouts[1])
    s1.shapes.title.text = "Kestrel Freight: where we stand"
    s1.placeholders[1].text = "Liability is strong on the delivery log\nDamages are documented\nMediation set for November"
    s1.notes_slide.notes_text_frame.text = ("Do not mention the $350k floor. Sam heard their CFO is leaving in Q1, "
                                            "so they want this closed.")

    s2 = prs.slides.add_slide(prs.slide_layouts[1])
    s2.shapes.title.text = "Fallback: walk-away numbers"
    s2.placeholders[1].text = "Floor $350,000\nFees split 60/40 Harlow/Pike"
    s2._element.set("show", "0")

    s3 = prs.slides.add_slide(prs.slide_layouts[5])
    s3.shapes.title.text = "Witnesses"
    tb = s3.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(0.6))
    tb.text_frame.text = "Dispatcher Dana Reyes (left Kestrel in May)"
    tb.text_frame.paragraphs[0].runs[0].font.size = PPt(24)
    box = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(1.95), Inches(8), Inches(0.7))
    box.fill.solid()
    box.fill.fore_color.rgb = PRGB(0, 0, 0)
    box.line.fill.background()
    tb2 = s3.shapes.add_textbox(Inches(1), Inches(3.2), Inches(8), Inches(0.6))
    tb2.text_frame.text = "Warehouse manager (deposition set)"
    off = s3.shapes.add_textbox(prs.slide_width + Inches(0.5), Inches(1), Inches(4), Inches(1))
    off.text_frame.text = "Old figure: $410,000 demand"
    s3.notes_slide.notes_text_frame.text = "Reyes has not agreed to testify yet."
    prs.save(path)

    authors = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
               '<p:cmAuthor id="0" name="%s" initials="SO" lastIdx="1" clrIdx="0"/></p:cmAuthorLst>' % EDITOR)
    comment = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
               '<p:cm authorId="0" dt="2026-09-13T09:30:00.000" idx="1"><p:pos x="10" y="10"/>'
               "<p:text>Take the CFO point out before we share this with the client.</p:text></p:cm></p:cmLst>")
    with zipfile.ZipFile(path) as z:
        pres_rels = z.read("ppt/_rels/presentation.xml.rels").decode()
        slide_rels = {n: z.read(n).decode() for n in z.namelist() if n.startswith("ppt/slides/_rels/")}
    first = next(n for n in slide_rels if n.endswith("slide1.xml.rels"))
    patch_zip(path,
              replace={"docProps/app.xml": app_xml("Microsoft Office PowerPoint", [
                  ("TotalTime", "95"), ("Company", FIRM.replace("&", "&amp;")), ("AppVersion", "16.0000")])},
              add={"ppt/commentAuthors.xml": authors.encode(), "ppt/comments/comment1.xml": comment.encode()},
              content_types=[("ppt/commentAuthors.xml",
                              "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml"),
                             ("ppt/comments/comment1.xml",
                              "application/vnd.openxmlformats-officedocument.presentationml.comments+xml")],
              rels={"ppt/_rels/presentation.xml.rels": [
                  '<Relationship Id="rIdCmA" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                  'relationships/commentAuthors" Target="commentAuthors.xml"/>'],
                  first: ['<Relationship Id="rIdCm1" Type="http://schemas.openxmlformats.org/officeDocument/'
                          '2006/relationships/comments" Target="../comments/comment1.xml"/>']})
    assert "commentAuthors" not in pres_rels


# ---- PDF ----------------------------------------------------------------------------------

LOG = [
    "Exhibit B - Delivery log extract, Brightwater Supply",
    "",
    "2025-03-04  Load 4471  Atlanta to Savannah  delivered late (31 h)",
    "2025-03-11  Load 4488  Atlanta to Macon     refused, temperature",
    "Driver: Dana Reyes, cell 555-0142",
    "2025-03-19  Load 4502  Atlanta to Augusta   delivered late (19 h)",
    "Reserve per claims memo: $1.2M",
    "2025-04-02  Load 4517  Atlanta to Savannah  lost in transit",
    "Kestrel carrier: Halvard Mutual",
    "Settlement floor discussed: $350,000",
    "Account 4471-0093 (Brightwater operating)",
]


def _draw_log(c, lines, boxes=True):
    y = 720
    pos = {}
    for i, line in enumerate(lines):
        c.setFont("Helvetica-Bold" if i == 0 else "Helvetica", 14 if i == 0 else 11)
        c.setFillColorRGB(0, 0, 0)
        c.drawString(72, y, line)
        pos[line] = y
        y -= 22
    if boxes:
        # black box drawn over the driver line, after the text: a fake redaction
        y0 = pos[LOG[4]]
        c.setFillColorRGB(0, 0, 0)
        c.rect(70, y0 - 4, 200, 16, stroke=0, fill=1)
        # white box over the reserve line: covered text
        y0 = pos[LOG[6]]
        c.setFillColorRGB(1, 1, 1)
        c.rect(70, y0 - 4, 200, 16, stroke=0, fill=1)
    return pos


def _rl_pdf(lines):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=letter, invariant=1, pageCompression=0)
    c.setAuthor(AUTHOR)
    c.setTitle("Exhibit B - delivery log")
    c.setSubject("")
    c.setCreator("Microsoft Word")
    pos = _draw_log(c, lines)
    c.showPage()
    c.setFont("Helvetica", 11)
    c.drawString(72, 720, "Page 2: summary of late deliveries (see page 1).")
    c.showPage()
    c.save()
    return buf.getvalue(), pos


def make_exhibit_pdf(path):
    data, pos = _rl_pdf(LOG)
    pdf = pikepdf.open(io.BytesIO(data))
    page = pdf.pages[0]
    annots = pikepdf.Array()
    annots.append(pdf.make_indirect(pikepdf.Dictionary(
        Type=pikepdf.Name.Annot, Subtype=pikepdf.Name.Text, Rect=[400, 690, 420, 710],
        T=pikepdf.String(EDITOR), Contents=pikepdf.String("Check with Jordan before we produce this page."),
        Name=pikepdf.Name.Comment)))
    y = pos[LOG[8]]
    annots.append(pdf.make_indirect(pikepdf.Dictionary(
        Type=pikepdf.Name.Annot, Subtype=pikepdf.Name.Square, Rect=[70, y - 4, 260, y + 12],
        C=[0, 0, 0], IC=[0, 0, 0], T=pikepdf.String(AUTHOR))))
    y = pos[LOG[10]]
    annots.append(pdf.make_indirect(pikepdf.Dictionary(
        Type=pikepdf.Name.Annot, Subtype=pikepdf.Name.Redact, Rect=[70, y - 4, 300, y + 12],
        QuadPoints=[70, y + 12, 300, y + 12, 70, y - 4, 300, y - 4], IC=[0, 0, 0],
        T=pikepdf.String(EDITOR))))
    page.obj.Annots = annots
    pdf.attachments["call-notes.txt"] = pikepdf.AttachedFileSpec(
        pdf, b"Call with Kestrel counsel 2026-09-11: they will move to $385k if we drop the late fees.\n",
        filename="call-notes.txt")
    pdf.Root.OpenAction = pdf.make_indirect(pikepdf.Dictionary(
        S=pikepdf.Name.JavaScript, JS=pikepdf.String("console.println('opened');")))
    with pdf.open_metadata(set_pikepdf_as_editor=False, update_docinfo=False) as meta:
        meta["dc:creator"] = [AUTHOR]
        meta["pdf:Producer"] = "Microsoft Word for Microsoft 365"
    pdf.docinfo["/Company"] = FIRM
    pdf.docinfo["/SourceModified"] = "D:20260912100500"
    pdf.docinfo["/Producer"] = "Microsoft Word for Microsoft 365"
    buf = io.BytesIO()
    pdf.save(buf, object_stream_mode=pikepdf.ObjectStreamMode.disable, deterministic_id=True,
             fix_metadata_version=False)
    base = buf.getvalue()

    # second save, appended as an incremental update: the settlement-floor line is taken out
    new_lines = ["" if l.startswith("Settlement floor") else l for l in LOG]
    new_data, _ = _rl_pdf(new_lines)
    with pikepdf.open(io.BytesIO(new_data)) as newer:
        new_stream = newer.pages[0].obj.Contents.read_bytes()
    pdf = pikepdf.open(io.BytesIO(base))
    page = pdf.pages[0].obj
    stream = pdf.make_stream(new_stream)
    page.Contents = stream
    num = stream.objgen[0]
    pnum, pgen = page.objgen
    prev = int(re.findall(rb"startxref\s+(\d+)", base)[-1])
    out = bytearray(base)
    if not out.endswith(b"\n"):
        out += b"\n"
    offsets = {}
    offsets[num] = len(out)
    out += b"%d 0 obj\n<< /Length %d >>\nstream\n" % (num, len(new_stream)) + new_stream + b"\nendstream\nendobj\n"
    offsets[pnum] = len(out)
    out += b"%d %d obj\n" % (pnum, pgen) + page.unparse(resolved=True) + b"\nendobj\n"
    xref = len(out)
    out += b"xref\n0 1\n0000000000 65535 f \n"
    for n in sorted(offsets):
        out += b"%d 1\n%010d %05d n \n" % (n, offsets[n], pgen if n == pnum else 0)
    root = pdf.trailer.Root.objgen
    info = pdf.trailer.Info.objgen
    out += (b"trailer\n<< /Size %d /Root %d %d R /Info %d %d R /Prev %d >>\nstartxref\n%d\n%%%%EOF\n"
            % (max(num, pnum) + 1, root[0], root[1], info[0], info[1], prev, xref))
    with open(path, "wb") as fh:
        fh.write(bytes(out))
    check = pikepdf.open(path)
    assert len(check.pages) == 2
    check.close()


def make_letter_pdf(path, use_word=True):
    """A letter blacked out with Word's highlighter, then saved as PDF. Returns how it was made."""
    body = ["Dear Ms. Alvarez,",
            "We write on behalf of Brightwater Supply regarding the March 2025 deliveries.",
            "Our client's records show that the carrier's insurer, ",
            "The parties met on September 11 and agreed to continue talks."]
    secret_text = "Halvard Mutual (claim HM-55120)"
    if use_word and sys.platform == "win32":
        try:
            return _letter_via_word(path, body, secret_text)
        except Exception as e:  # Word missing or COM blocked: fall back to drawing the same thing
            print("Word export failed (%s), using reportlab" % e)
    c = canvas.Canvas(path, pagesize=letter, invariant=1)
    c.setAuthor(AUTHOR)
    c.setTitle("Letter to Kestrel counsel")
    c.setSubject("")
    c.setCreator("Microsoft Word")
    c.setFont("Helvetica", 11)
    y = 700
    for i, line in enumerate(body):
        c.drawString(72, y, line)
        if i == 2:
            w = c.stringWidth(line, "Helvetica", 11)
            sw = c.stringWidth(secret_text, "Helvetica", 11)
            c.setFillColorRGB(0, 0, 0)
            c.rect(72 + w, y - 3, sw, 14, stroke=0, fill=1)
            c.drawString(72 + w, y, secret_text)
            c.drawString(72 + w + sw, y, ", accepted coverage.")
        y -= 20
    c.showPage()
    c.save()
    return "reportlab"


def _letter_via_word(path, body, secret_text):
    import win32com.client

    tmp = tempfile.mkdtemp()
    src = os.path.join(tmp, "letter.docx")
    d = docx.Document()
    cp = d.core_properties
    cp.author, cp.last_modified_by, cp.title = AUTHOR, AUTHOR, "Letter to Kestrel counsel"
    cp.comments = cp.subject = cp.keywords = cp.category = ""
    cp.created = cp.modified = WHEN
    for i, line in enumerate(body):
        p = d.add_paragraph(line)
        if i == 2:
            r = p.add_run(secret_text)
            r.font.highlight_color = WD_COLOR_INDEX.BLACK
            r.font.color.rgb = RGBColor(0, 0, 0)
            p.add_run(", accepted coverage.")
    d.save(src)
    patch_zip(src, replace={"docProps/app.xml": app_xml("Microsoft Office Word", [
        ("Company", FIRM.replace("&", "&amp;")), ("AppVersion", "16.0000")])})
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Open(src, ReadOnly=True, AddToRecentFiles=False)
        doc.ExportAsFixedFormat(os.path.abspath(path), 17)  # wdExportFormatPDF
        doc.Close(0)
    finally:
        word.Quit()
        shutil.rmtree(tmp, ignore_errors=True)
    return "Microsoft Word"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--no-word", action="store_true")
    args = ap.parse_args()
    os.makedirs(OUT, exist_ok=True)
    make_docx(os.path.join(OUT, "settlement-draft.docx"))
    make_xlsx(os.path.join(OUT, "damages-model.xlsx"))
    make_pptx(os.path.join(OUT, "strategy-deck.pptx"))
    make_exhibit_pdf(os.path.join(OUT, "exhibit-b.pdf"))
    how = make_letter_pdf(os.path.join(OUT, "letter-to-counsel.pdf"), not args.no_word)
    print("wrote examples/ (letter PDF made with %s)" % how)


if __name__ == "__main__":
    main()
