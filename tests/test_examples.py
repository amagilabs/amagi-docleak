"""The bundled example files: every planted leak is found, and cleaning removes it."""

import pytest
from docx import Document
from openpyxl import load_workbook
from pptx import Presentation

from amagi_docleak.clean import clean_file
from amagi_docleak.scan import check_file

from helpers import EXAMPLES, check_ids, pdf_text, zip_text

EXPECTED = {
    "settlement-draft.docx": {"tracked-changes", "comments", "fake-redaction", "hidden-text", "document-variables",
                              "author-metadata", "custom-properties", "file-paths", "thumbnail"},
    "damages-model.xlsx": {"hidden-sheets", "fake-redaction", "comments", "hidden-rows-columns", "hidden-cells",
                           "data-connections", "author-metadata", "custom-properties", "file-paths"},
    "strategy-deck.pptx": {"speaker-notes", "comments", "hidden-slides", "off-slide-content",
                           "possible-redaction-box", "author-metadata", "thumbnail"},
    "exhibit-b.pdf": {"comments", "fake-redaction", "unapplied-redaction", "covered-text", "author-metadata",
                      "custom-properties", "embedded-files", "javascript", "prior-versions"},
    "letter-to-counsel.pdf": {"fake-redaction"},
}


@pytest.mark.parametrize("name", sorted(EXPECTED))
def test_example_findings(name):
    report = check_file(str(EXAMPLES / name))
    assert report.error is None
    assert check_ids(report) == EXPECTED[name]
    assert report.failed("fail")


def test_docx_details():
    report = check_file(str(EXAMPLES / "settlement-draft.docx"))
    by = {f.check: f for f in report.findings}
    assert "3 tracked changes" in by["tracked-changes"].summary
    assert any("$350,000" in d for d in by["tracked-changes"].details)
    assert "Track Changes is still switched on" in by["tracked-changes"].details
    assert by["fake-redaction"].secret
    assert any("call-notes-0911.docx" in d for d in by["file-paths"].details)
    assert any("HarlowPike-Settlement.dotx" in d for d in by["file-paths"].details)


def test_pdf_details():
    report = check_file(str(EXAMPLES / "exhibit-b.pdf"))
    by = {f.check: f for f in report.findings}
    assert "2 black boxes" in by["fake-redaction"].summary
    assert any("Settlement floor discussed" in d for d in by["prior-versions"].details)
    assert any("text still underneath" in d for d in by["unapplied-redaction"].details)
    assert "call-notes.txt" in by["embedded-files"].summary
    assert "Sam Okafor" in by["comments"].summary


def _clean(tmp_path, name, **kw):
    dst = tmp_path / name.replace(".", ".clean.")
    actions = clean_file(str(EXAMPLES / name), str(dst), **kw)
    return dst, actions


def test_clean_docx(tmp_path):
    dst, actions = _clean(tmp_path, "settlement-draft.docx")
    assert any("accepted 3 tracked changes" in a for a in actions)
    left = check_ids(check_file(str(dst)))
    assert left == {"file-paths"}  # the hyperlink to another file stays; it is part of the text
    xml = zip_text(dst)
    for gone in ("$410,000", "will not go below", "KF-2291", "arbitrator", "Sam Okafor", "Jordan Ellery",
                 "HarlowPike", "BRW-0412", "w:del ", "w:ins "):
        assert gone not in xml, gone
    doc = Document(str(dst))
    body = "\n".join(p.text for p in doc.paragraphs)
    assert "$385,000" in body
    assert "█" * 10 in body


def test_clean_docx_unhide_and_keep(tmp_path):
    dst, actions = _clean(tmp_path, "settlement-draft.docx", keep={"author-metadata"}, unhide=True)
    assert any("made 2 runs of hidden text visible" in a for a in actions)
    report = check_file(str(dst))
    assert "author-metadata" in check_ids(report)
    assert "hidden-text" not in check_ids(report)
    body = "\n".join(p.text for p in Document(str(dst)).paragraphs)
    assert "arbitrator leans our way" in body


def test_clean_xlsx(tmp_path):
    dst, actions = _clean(tmp_path, "damages-model.xlsx")
    left = check_ids(check_file(str(dst)))
    assert left == {"hidden-sheets", "hidden-rows-columns", "hidden-cells", "data-connections"}
    xml = zip_text(dst)
    for gone in ("Halvard reserve", "Padded about 15%", "Jordan Ellery", "Harlow &amp; Pike", "jellery", "BRW-0412"):
        assert gone not in xml, gone
    wb = load_workbook(str(dst))
    assert wb["Summary"]["B6"].value == "█" * len("Halvard reserve $1.2M")


def test_clean_xlsx_unhide(tmp_path):
    dst, actions = _clean(tmp_path, "damages-model.xlsx", unhide=True)
    left = check_ids(check_file(str(dst)))
    assert "hidden-sheets" not in left and "hidden-rows-columns" not in left
    wb = load_workbook(str(dst))
    assert all(ws.sheet_state == "visible" for ws in wb.worksheets)


def test_clean_pptx(tmp_path):
    dst, actions = _clean(tmp_path, "strategy-deck.pptx")
    left = check_ids(check_file(str(dst)))
    assert left == {"hidden-slides", "off-slide-content", "possible-redaction-box"}
    xml = zip_text(dst)
    for gone in ("CFO is leaving", "has not agreed to testify", "Sam Okafor", "Take the CFO point out"):
        assert gone not in xml, gone
    prs = Presentation(str(dst))
    assert len(prs.slides) == 3
    assert not any(s.has_notes_slide for s in prs.slides)


def test_clean_pdf_burns_in_redactions(tmp_path):
    dst, actions = _clean(tmp_path, "exhibit-b.pdf")
    assert any("200 dpi image" in a for a in actions)
    assert any("earlier versions" in a for a in actions)
    report = check_file(str(dst))
    assert not report.findings
    raw = pdf_text(dst)
    for gone in ("Dana Reyes", "Halvard", "4471-0093", "350,000", "Okafor", "Jordan Ellery", "call-notes",
                 "JavaScript"):
        assert gone not in raw, gone


def test_clean_word_pdf(tmp_path):
    dst, actions = _clean(tmp_path, "letter-to-counsel.pdf")
    assert not check_file(str(dst)).findings
    assert "Halvard" not in pdf_text(dst)
