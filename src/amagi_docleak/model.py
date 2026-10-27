"""Findings, per-file reports and the check registry."""

from dataclasses import dataclass, field
from typing import List, Optional

RANK = {"info": 1, "warn": 2, "fail": 3}


@dataclass(frozen=True)
class Check:
    id: str
    severity: str
    formats: str
    description: str
    clean: str


CHECKS = {c.id: c for c in [
    Check("tracked-changes", "fail", "docx",
          "Tracked insertions, deletions, moves and formatting changes",
          "accepts every change"),
    Check("comments", "fail", "docx xlsx pptx pdf",
          "Comments, threaded replies, reviewer names and PDF sticky notes",
          "removes them with their author lists"),
    Check("hidden-text", "fail", "docx",
          "Text formatted as hidden, colored to match its background, or 1pt",
          "deletes the text (--unhide makes it visible instead)"),
    Check("fake-redaction", "fail", "docx xlsx pdf",
          "Black boxes or black highlight over text that can still be selected and copied",
          "Word/Excel: replaces the text with blocks; PDF: flattens the page to an image"),
    Check("unapplied-redaction", "fail", "pdf",
          "PDF redaction marks that were drawn but never applied",
          "burns the box in and flattens the page to an image"),
    Check("speaker-notes", "fail", "pptx",
          "Speaker notes",
          "removes the notes pages"),
    Check("hidden-sheets", "fail", "xlsx",
          "Hidden and very hidden worksheets",
          "left in place (--unhide makes them visible)"),
    Check("hidden-rows-columns", "warn", "xlsx",
          "Hidden or filtered-out rows and columns that hold data",
          "left in place (--unhide makes them visible)"),
    Check("hidden-cells", "warn", "xlsx",
          "Cell values hidden by a ;;; number format or a font that matches the fill",
          "left in place"),
    Check("hidden-slides", "warn", "pptx",
          "Slides marked hidden",
          "left in place (--unhide makes them visible)"),
    Check("off-slide-content", "warn", "pptx",
          "Shapes parked outside the visible slide area",
          "left in place"),
    Check("possible-redaction-box", "warn", "docx pptx",
          "Black shapes that may sit on top of text (the text underneath is still there)",
          "left in place"),
    Check("covered-text", "warn", "pdf",
          "Text drawn over by a white box or an image",
          "left in place"),
    Check("author-metadata", "warn", "docx xlsx pptx pdf",
          "Author, last editor, company, manager, title, keywords and editing time",
          "clears the properties"),
    Check("custom-properties", "warn", "docx xlsx pptx pdf",
          "Custom document properties (document IDs, client-matter numbers, DMS fields)",
          "removes them"),
    Check("file-paths", "warn", "docx xlsx pptx pdf",
          "Local and network file paths, template paths and links to other files",
          "removes template and save-location paths; other links left in place"),
    Check("embedded-files", "warn", "docx xlsx pptx pdf",
          "Embedded files and objects (a pasted chart can carry its whole workbook)",
          "PDF: removes attachments; Office: left in place"),
    Check("pivot-cache", "warn", "xlsx",
          "Pivot table source data cached in the workbook",
          "drops the cached records (the pivot keeps its current values)"),
    Check("data-connections", "warn", "xlsx",
          "External data connections and connection strings",
          "left in place"),
    Check("document-variables", "warn", "docx",
          "Word document variables",
          "removes them"),
    Check("macros", "warn", "docx xlsx pptx",
          "VBA macro projects",
          "left in place"),
    Check("prior-versions", "warn", "pdf",
          "Incremental saves: earlier versions of the PDF are still inside the file",
          "rewrites the file without them"),
    Check("javascript", "warn", "pdf",
          "JavaScript and launch actions",
          "removes them"),
    Check("hidden-layers", "warn", "pdf",
          "Optional content layers that are switched off",
          "left in place (--unhide switches them on)"),
    Check("private-app-data", "warn", "pdf",
          "Application private data (PieceInfo), which can hold an editable original",
          "removes it"),
    Check("thumbnail", "info", "docx xlsx pptx",
          "Saved preview thumbnail of the first page",
          "removes it"),
    Check("unsupported-format", "fail", "doc xls ppt",
          "Legacy binary Office files this tool cannot read",
          "not cleaned; save as .docx/.xlsx/.pptx first"),
    Check("unreadable", "fail", "all",
          "Files that could not be opened (password-protected, corrupt, wrong extension)",
          "not cleaned"),
]}


@dataclass
class Finding:
    check: str
    summary: str
    details: List[str] = field(default_factory=list)
    secret: bool = False  # details quote text the author meant to hide; masked unless --reveal
    severity: str = ""

    def __post_init__(self):
        if not self.severity:
            self.severity = CHECKS[self.check].severity


@dataclass
class FileReport:
    path: str
    fmt: str
    findings: List[Finding] = field(default_factory=list)
    error: Optional[str] = None

    def worst(self):
        return max((RANK[f.severity] for f in self.findings), default=0)

    def failed(self, fail_on):
        if fail_on == "never":
            return False
        return self.worst() >= RANK[fail_on]

    def counts(self):
        out = {"fail": 0, "warn": 0, "info": 0}
        for f in self.findings:
            out[f.severity] += 1
        return out
