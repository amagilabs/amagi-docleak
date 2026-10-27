"""Pick the right checker for each file and collect the findings."""

import os

from . import docx, pdf, pptx, xlsx
from .model import FileReport, Finding
from .ooxml import NotOffice, Package, check_common

OOXML = {".docx": "docx", ".docm": "docx", ".xlsx": "xlsx", ".xlsm": "xlsx", ".pptx": "pptx", ".pptm": "pptx"}
LEGACY = {".doc": "doc", ".xls": "xls", ".ppt": "ppt", ".rtf": "rtf"}
SUPPORTED = set(OOXML) | {".pdf"}
CHECKERS = {"docx": docx.check, "xlsx": xlsx.check, "pptx": pptx.check}
MAIN_TYPES = {"docx": "word/", "xlsx": "xl/", "pptx": "ppt/"}


def fmt_of(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in OOXML:
        return OOXML[ext]
    if ext == ".pdf":
        return "pdf"
    return LEGACY.get(ext)


def iter_files(paths):
    """Expand directories; skip Office lock files and formats we do not read."""
    for p in paths:
        if os.path.isdir(p):
            for root, dirs, files in os.walk(p):
                dirs[:] = sorted(d for d in dirs if not d.startswith("."))
                for name in sorted(files):
                    if name.startswith("~$"):
                        continue
                    if fmt_of(name):
                        yield os.path.join(root, name)
        else:
            yield p


def check_file(path):
    fmt = fmt_of(path)
    report = FileReport(path, fmt or "unknown")
    if not os.path.isfile(path):
        report.error = "file not found"
        report.findings.append(Finding("unreadable", "file not found"))
        return report
    if fmt in LEGACY.values():
        report.findings.append(Finding("unsupported-format", "legacy .%s file; save it as .%sx (or PDF) and check "
                                       "that copy" % (fmt, fmt if fmt != "rtf" else "doc")))
        return report
    if fmt is None:
        report.error = "unsupported file type"
        report.findings.append(Finding("unsupported-format", "not a .docx, .xlsx, .pptx or .pdf file"))
        return report
    try:
        if fmt == "pdf":
            pdf.check(path, report.findings)
        else:
            pkg = Package(path)
            try:
                main = pkg.main_part()
                if not main or not main.startswith(MAIN_TYPES[fmt]):
                    raise NotOffice("the extension says .%s but the contents are not" % fmt)
                CHECKERS[fmt](pkg, report.findings)
                check_common(pkg, report.findings)
            finally:
                pkg.close()
    except (NotOffice, ValueError) as e:
        report.error = str(e)
        report.findings.append(Finding("unreadable", str(e)))
    except Exception as e:  # a malformed file must not stop a folder scan
        report.error = "%s: %s" % (type(e).__name__, e)
        report.findings.append(Finding("unreadable", "could not read the file (%s)" % report.error[:160]))
    order = {"fail": 0, "warn": 1, "info": 2}
    report.findings.sort(key=lambda f: order[f.severity])
    return report
