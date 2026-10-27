"""Small files built in the test, one planted problem each, to pin down what counts and what does not."""

import zipfile

import pikepdf
import pytest
from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import RGBColor
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

from amagi_docleak.clean import clean_file
from amagi_docleak.scan import check_file

from helpers import append_update, box, check_ids, make_pdf, patch_zip, pdf_text, text_line, zip_text


# ---- PDF ----------------------------------------------------------------------------------

def test_pdf_clean_file_has_no_findings(tmp_path):
    p = make_pdf(tmp_path / "plain.pdf", text_line(72, 700, "Nothing to see here"))
    assert check_file(str(p)).findings == []


def test_pdf_black_box_after_text_is_fake(tmp_path):
    p = make_pdf(tmp_path / "fake.pdf", text_line(72, 700, "Account 4471") + box(70, 695, 120, 18))
    r = check_file(str(p))
    assert check_ids(r) == {"fake-redaction"}
    assert "4471" in r.findings[0].details[0]  # kept in the finding, masked at print time


def test_pdf_box_drawn_before_text_is_not_fake(tmp_path):
    # white text on a black label: readable, nothing hidden
    p = make_pdf(tmp_path / "label.pdf", box(70, 695, 120, 18) + text_line(72, 700, "LABEL", color="1 g"))
    assert check_file(str(p)).findings == []


def test_pdf_black_text_on_black_box_is_fake(tmp_path):
    p = make_pdf(tmp_path / "dark.pdf", box(70, 695, 120, 18) + text_line(72, 700, "Hidden figure"))
    assert check_ids(check_file(str(p))) == {"fake-redaction"}


def test_pdf_invisible_ocr_text_is_ignored(tmp_path):
    # render mode 3 = invisible text, as OCR layers use; a box over it hides nothing
    p = make_pdf(tmp_path / "ocr.pdf", text_line(72, 700, "scanned words", mode=3) + box(70, 695, 120, 18))
    assert check_file(str(p)).findings == []


def test_pdf_white_box_is_covered_text(tmp_path):
    p = make_pdf(tmp_path / "white.pdf", text_line(72, 700, "Old clause") + box(70, 695, 120, 18, "1 g"))
    assert check_ids(check_file(str(p))) == {"covered-text"}


def test_pdf_block_characters_are_real_redaction(tmp_path):
    p = make_pdf(tmp_path / "x.pdf", text_line(72, 700, "XXXX") + box(70, 695, 120, 18))
    assert check_ids(check_file(str(p))) == {"fake-redaction"}
    # an already blacked-out run of block-like characters is not text worth reporting
    p2 = make_pdf(tmp_path / "blank.pdf", text_line(72, 700, "    ") + box(70, 695, 120, 18))
    assert check_file(str(p2)).findings == []


def test_pdf_metadata_and_custom_keys(tmp_path):
    p = make_pdf(tmp_path / "meta.pdf", text_line(72, 700, "hi"),
                 info={"/Author": "Pat Example", "/Producer": "Writer 1.0", "/MatterNo": "M-17"})
    r = check_file(str(p))
    assert check_ids(r) == {"author-metadata", "custom-properties"}
    dst = tmp_path / "meta.clean.pdf"
    clean_file(str(p), str(dst))
    assert check_file(str(dst)).findings == []
    assert "Pat Example" not in pdf_text(dst) and "M-17" not in pdf_text(dst)


def test_pdf_dark_square_annotation_over_text(tmp_path):
    annot = dict(Type=pikepdf.Name.Annot, Subtype=pikepdf.Name.Square, Rect=[70, 695, 190, 713],
                 IC=[0, 0, 0], C=[0, 0, 0])
    p = make_pdf(tmp_path / "sq.pdf", text_line(72, 700, "Client SSN"), annots=[annot])
    assert check_ids(check_file(str(p))) == {"fake-redaction"}
    dst = tmp_path / "sq.clean.pdf"
    clean_file(str(p), str(dst))
    assert check_file(str(dst)).findings == []
    assert "Client SSN" not in pdf_text(dst)


def test_pdf_incremental_update_keeps_old_text(tmp_path):
    p = make_pdf(tmp_path / "rev.pdf", text_line(72, 700, "Offer is 9137"))
    append_update(p, text_line(72, 700, "Offer is 5000"))
    r = check_file(str(p))
    assert check_ids(r) == {"prior-versions"}
    pv = r.findings[0]
    assert any("Offer is 9137" in d for d in pv.details)
    dst = tmp_path / "rev.clean.pdf"
    clean_file(str(p), str(dst))
    assert check_file(str(dst)).findings == []
    assert "9137" not in dst.read_bytes().decode("latin-1")


def test_pdf_not_a_pdf(tmp_path):
    p = tmp_path / "broken.pdf"
    p.write_bytes(b"this is not a pdf")
    r = check_file(str(p))
    assert check_ids(r) == {"unreadable"}
    assert r.failed("fail")


# ---- Word ---------------------------------------------------------------------------------

def _docx(path, build):
    doc = Document()
    cp = doc.core_properties
    cp.author = cp.last_modified_by = ""
    cp.comments = cp.title = cp.subject = cp.keywords = cp.category = ""
    build(doc)
    doc.save(str(path))
    return path


def test_docx_plain_file_passes(tmp_path):
    p = _docx(tmp_path / "plain.docx", lambda d: d.add_paragraph("Hello"))
    r = check_file(str(p))
    assert not r.failed("fail"), [(f.check, f.summary) for f in r.findings]
    assert "tracked-changes" not in check_ids(r)


def test_docx_white_text_and_highlight_box(tmp_path):
    def build(d):
        run = d.add_paragraph().add_run("secret note")
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        black = d.add_paragraph().add_run("Account 12")
        black.font.highlight_color = WD_COLOR_INDEX.BLACK
        black.font.color.rgb = RGBColor(0, 0, 0)  # automatic color would turn white on black and stay readable
    p = _docx(tmp_path / "w.docx", build)
    ids = check_ids(check_file(str(p)))
    assert {"hidden-text", "fake-redaction"} <= ids
    dst = tmp_path / "w.clean.docx"
    clean_file(str(p), str(dst))
    left = check_ids(check_file(str(dst)))
    assert "hidden-text" not in left and "fake-redaction" not in left
    text = zip_text(dst)
    assert "secret note" not in text and "Account 12" not in text


def test_docx_tracked_insertion_and_deletion_accepted(tmp_path):
    def build(d):
        para = d.add_paragraph()
        para.add_run("Pay ")
        ins = OxmlElement("w:ins")
        for k, v in (("w:id", "1"), ("w:author", "Rev One"), ("w:date", "2026-01-01T00:00:00Z")):
            ins.set(qn(k), v)
        r = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.text = "5000"
        r.append(t)
        ins.append(r)
        para._p.append(ins)
        dl = OxmlElement("w:del")
        for k, v in (("w:id", "2"), ("w:author", "Rev One"), ("w:date", "2026-01-01T00:00:00Z")):
            dl.set(qn(k), v)
        r = OxmlElement("w:r")
        t = OxmlElement("w:delText")
        t.text = "9137"
        r.append(t)
        dl.append(r)
        para._p.append(dl)
    p = _docx(tmp_path / "t.docx", build)
    r = check_file(str(p))
    tc = next(f for f in r.findings if f.check == "tracked-changes")
    assert "2 tracked changes" in tc.summary and "Rev One" in tc.summary
    dst = tmp_path / "t.clean.docx"
    clean_file(str(p), str(dst))
    assert "tracked-changes" not in check_ids(check_file(str(dst)))
    assert Document(str(dst)).paragraphs[-1].text == "Pay 5000"
    assert "9137" not in zip_text(dst)


# ---- Excel --------------------------------------------------------------------------------

def _xlsx(path, build):
    wb = Workbook()
    wb.properties.creator = ""
    wb.properties.lastModifiedBy = ""
    build(wb)
    wb.save(str(path))
    return path


def test_xlsx_plain_file_passes(tmp_path):
    p = _xlsx(tmp_path / "plain.xlsx", lambda wb: wb.active.append(["a", 1]))
    r = check_file(str(p))
    assert not r.failed("fail"), [(f.check, f.summary) for f in r.findings]


def test_xlsx_solid_fill_without_color_is_not_black(tmp_path):
    # Excel shows a solid fill with no fgColor as no fill, so black text on it is readable
    def build(wb):
        ws = wb.active
        ws["A1"] = "visible"
        ws["A1"].fill = PatternFill("solid")
        ws["A1"].font = Font(color="FF000000")
    p = _xlsx(tmp_path / "nofg.xlsx", build)
    assert "fake-redaction" not in check_ids(check_file(str(p)))


def test_xlsx_black_on_black_and_white_on_white(tmp_path):
    def build(wb):
        ws = wb.active
        ws["A1"] = "reserve 2M"
        ws["A1"].fill = PatternFill("solid", fgColor="FF000000")
        ws["A1"].font = Font(color="FF000000")
        ws["A2"] = "white note"
        ws["A2"].font = Font(color="FFFFFFFF")
    p = _xlsx(tmp_path / "bb.xlsx", build)
    ids = check_ids(check_file(str(p)))
    assert "fake-redaction" in ids and "hidden-cells" in ids


def test_xlsx_very_hidden_sheet(tmp_path):
    def build(wb):
        wb.create_sheet("Internal").sheet_state = "veryHidden"
    p = _xlsx(tmp_path / "vh.xlsx", build)
    r = check_file(str(p))
    hs = next(f for f in r.findings if f.check == "hidden-sheets")
    assert "Internal" in " ".join([hs.summary] + hs.details)
    dst = tmp_path / "vh.clean.xlsx"
    clean_file(str(p), str(dst), unhide=True)
    assert load_workbook(str(dst))["Internal"].sheet_state == "visible"


PIVOT_DEF = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
             'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" '
             'refreshOnLoad="1" recordCount="2"><cacheSource type="worksheet">'
             '<worksheetSource ref="A1:B3" sheet="Raw salaries"/></cacheSource>'
             '<cacheFields count="2"><cacheField name="Name" numFmtId="0"><sharedItems/></cacheField>'
             '<cacheField name="Pay" numFmtId="0"><sharedItems containsNumber="1"/></cacheField></cacheFields>'
             '</pivotCacheDefinition>')
PIVOT_REC = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2">'
             '<r><s v="Lee Sample"/><n v="91000"/></r><r><s v="Kim Sample"/><n v="87000"/></r>'
             '</pivotCacheRecords>')
PIVOT_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
              '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
              'pivotCacheRecords" Target="pivotCacheRecords1.xml"/></Relationships>')


def _add_pivot_cache(path):
    with zipfile.ZipFile(path) as z:
        rels = z.read("xl/_rels/workbook.xml.rels").decode()
        types = z.read("[Content_Types].xml").decode()
    rels = rels.replace("</Relationships>",
                        '<Relationship Id="rIdPv1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                        'relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/>'
                        "</Relationships>")
    ct = "application/vnd.openxmlformats-officedocument.spreadsheetml."
    types = types.replace("</Types>",
                          '<Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="%s'
                          'pivotCacheDefinition+xml"/><Override PartName="/xl/pivotCache/pivotCacheRecords1.xml" '
                          'ContentType="%spivotCacheRecords+xml"/></Types>' % (ct, ct))
    patch_zip(path, {
        "xl/_rels/workbook.xml.rels": rels,
        "[Content_Types].xml": types,
        "xl/pivotCache/pivotCacheDefinition1.xml": PIVOT_DEF,
        "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels": PIVOT_RELS,
        "xl/pivotCache/pivotCacheRecords1.xml": PIVOT_REC,
    })


def test_xlsx_pivot_cache_found_and_dropped(tmp_path):
    p = _xlsx(tmp_path / "pv.xlsx", lambda wb: wb.active.append(["Total", 2]))
    _add_pivot_cache(p)
    r = check_file(str(p))
    pv = next(f for f in r.findings if f.check == "pivot-cache")
    assert "2 source rows cached" in pv.details[0]
    assert "no longer in the workbook" in pv.details[0]
    dst = tmp_path / "pv.clean.xlsx"
    clean_file(str(p), str(dst))
    assert "pivot-cache" not in check_ids(check_file(str(dst)))
    with zipfile.ZipFile(dst) as z:
        names = z.namelist()
        definition = z.read("xl/pivotCache/pivotCacheDefinition1.xml").decode()
    assert "xl/pivotCache/pivotCacheRecords1.xml" not in names
    assert 'saveData="0"' in definition
    assert "Lee Sample" not in zip_text(dst)


# ---- PowerPoint ---------------------------------------------------------------------------

def test_pptx_speaker_notes_removed(tmp_path):
    from pptx import Presentation
    prs = Presentation()
    prs.core_properties.author = prs.core_properties.last_modified_by = ""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.notes_slide.notes_text_frame.text = "say the number is soft"
    p = tmp_path / "n.pptx"
    prs.save(str(p))
    r = check_file(str(p))
    assert "speaker-notes" in check_ids(r)
    dst = tmp_path / "n.clean.pptx"
    clean_file(str(p), str(dst))
    assert "speaker-notes" not in check_ids(check_file(str(dst)))
    assert "number is soft" not in zip_text(dst)
    assert len(Presentation(str(dst)).slides) == 1


def test_unsupported_type_raises(tmp_path):
    p = tmp_path / "a.txt"
    p.write_text("x")
    with pytest.raises(ValueError):
        clean_file(str(p), str(tmp_path / "b.txt"))
