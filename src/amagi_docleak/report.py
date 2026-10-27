"""Text and JSON reports."""

import json
import os
import sys

from . import __version__
from .model import CHECKS
from .util import plural, render

COLORS = {"FAIL": "31", "PASS": "32", "fail": "31", "warn": "33", "info": "36"}


def _paint(word, on):
    return "\x1b[%sm%s\x1b[0m" % (COLORS[word], word) if on else word


def use_color(stream):
    return hasattr(stream, "isatty") and stream.isatty() and not os.environ.get("NO_COLOR")


def verdict(report, fail_on):
    return "FAIL" if report.failed(fail_on) else "PASS"


def text(reports, fail_on="fail", reveal=False, detail=3, color=False, out=sys.stdout):
    width = max([len(f.check) for r in reports for f in r.findings] + [14])
    for r in reports:
        v = verdict(r, fail_on)
        counts = r.counts()
        tail = ", ".join(plural(counts[k], k if k != "info" else "note") for k in ("fail", "warn", "info") if counts[k])
        out.write("%s  %s%s\n" % (_paint(v, color), r.path, "  (%s)" % tail if tail else "  (nothing found)"))
        if detail < 0:
            continue
        for f in r.findings:
            out.write("  %s  %-*s  %s\n" % (_paint(f.severity, color), width, f.check, render(f.summary, reveal)))
            shown = f.details if detail == 0 else f.details[:detail]
            for d in shown:
                out.write("  %s  %s\n" % (" " * (4 + width), render(d, reveal)))
            if len(f.details) > len(shown):
                out.write("  %s  ... %d more (use --verbose)\n" % (" " * (4 + width), len(f.details) - len(shown)))
        if r.findings:
            out.write("\n")
    failed = sum(1 for r in reports if r.failed(fail_on))
    if len(reports) > 1 or detail < 0:
        out.write("%s checked: %d failed, %d passed\n" % (plural(len(reports), "file"), failed, len(reports) - failed))
    if failed and detail >= 0 and any(f.check in CHECKS and f.secret for r in reports for f in r.findings) \
            and not reveal:
        out.write("Hidden text is masked. Add --reveal to print it.\n")


def as_json(reports, fail_on="fail", reveal=False):
    return json.dumps({
        "tool": "amagi-docleak",
        "version": __version__,
        "fail_on": fail_on,
        "files": [{
            "path": r.path,
            "format": r.fmt,
            "verdict": verdict(r, fail_on),
            "error": r.error,
            "counts": r.counts(),
            "findings": [{"check": f.check, "severity": f.severity, "summary": render(f.summary, reveal),
                          "details": [render(d, reveal) for d in f.details]} for f in r.findings],
        } for r in reports],
        "failed": sum(1 for r in reports if r.failed(fail_on)),
    }, indent=2, ensure_ascii=False)
