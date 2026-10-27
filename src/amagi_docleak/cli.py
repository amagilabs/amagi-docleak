"""Command line: amagi-docleak [check|clean|checks] ..."""

import argparse
import os
import sys

from . import __version__
from .model import CHECKS
from .report import as_json, text, use_color
from .scan import SUPPORTED, check_file, iter_files

COMMANDS = {"check", "clean", "checks"}


def _check_list(value):
    items = [v.strip() for v in value.split(",") if v.strip()]
    bad = [v for v in items if v not in CHECKS]
    if bad:
        raise argparse.ArgumentTypeError("unknown check %s (run 'amagi-docleak checks' for the list)"
                                         % ", ".join(bad))
    return set(items)


def _common(p):
    p.add_argument("paths", nargs="+", metavar="PATH", help="files or folders")
    p.add_argument("--json", action="store_true", help="print a JSON report")
    p.add_argument("--fail-on", choices=["fail", "warn", "info", "never"], default="fail",
                   help="lowest severity that makes a file FAIL and the exit code 1 (default: fail)")
    p.add_argument("--skip", type=_check_list, default=set(), metavar="CHECK,...",
                   help="ignore these checks")
    p.add_argument("--reveal", action="store_true", help="print hidden text instead of masking it")
    g = p.add_mutually_exclusive_group()
    g.add_argument("-v", "--verbose", action="store_true", help="show every detail line")
    g.add_argument("-q", "--quiet", action="store_true", help="one line per file")


def build_parser():
    parser = argparse.ArgumentParser(
        prog="amagi-docleak",
        description="Check Word, Excel, PowerPoint and PDF files for tracked changes, comments, hidden "
                    "content, personal metadata and fake redactions before you send them.",
        epilog="A bare path runs 'check': amagi-docleak contract.docx")
    parser.add_argument("--version", action="version", version="amagi-docleak " + __version__)
    sub = parser.add_subparsers(dest="command", metavar="COMMAND")
    p = sub.add_parser("check", help="report what each file leaks (default)")
    _common(p)
    p = sub.add_parser("clean", help="write a cleaned copy, then check the copy")
    _common(p)
    out = p.add_mutually_exclusive_group()
    out.add_argument("-o", "--output", metavar="FILE", help="output file (one input only)")
    out.add_argument("--out-dir", metavar="DIR", help="write cleaned copies into this folder")
    p.add_argument("--force", action="store_true", help="overwrite an existing output file")
    p.add_argument("--keep", type=_check_list, default=set(), metavar="CHECK,...",
                   help="leave these alone, e.g. --keep author-metadata,comments")
    p.add_argument("--unhide", action="store_true",
                   help="make hidden text, sheets, rows, slides and layers visible instead of deleting or "
                        "keeping them hidden")
    sub.add_parser("checks", help="list every check with its severity")
    return parser


def _filter(reports, skip):
    for r in reports:
        r.findings = [f for f in r.findings if f.check not in skip]
    return reports


def _emit(reports, args):
    if args.json:
        print(as_json(reports, args.fail_on, args.reveal))
    else:
        detail = -1 if args.quiet else (0 if args.verbose else 3)
        text(reports, args.fail_on, args.reveal, detail, use_color(sys.stdout), sys.stdout)
    return 1 if any(r.failed(args.fail_on) for r in reports) else 0


def cmd_checks():
    width = max(len(c) for c in CHECKS)
    for c in CHECKS.values():
        print("%-*s  %-4s  %-19s  %s" % (width, c.id, c.severity, c.formats, c.description))
    print("\nclean: what 'amagi-docleak clean' does for each check")
    for c in CHECKS.values():
        print("%-*s  %s" % (width, c.id, c.clean))
    return 0


def output_path(src, args):
    stem, ext = os.path.splitext(os.path.basename(src))
    name = stem + ".clean" + ext
    if args.output:
        return args.output
    if args.out_dir:
        return os.path.join(args.out_dir, name)
    return os.path.join(os.path.dirname(src), name)


def cmd_clean(args, files):
    from .clean import clean_file

    if args.output and len(files) != 1:
        print("amagi-docleak: -o works with one input file; use --out-dir for several", file=sys.stderr)
        return 2
    if args.out_dir:
        os.makedirs(args.out_dir, exist_ok=True)
    reports, logs = [], []
    for src in files:
        if os.path.splitext(src)[1].lower() not in SUPPORTED:
            reports.append(check_file(src))
            continue
        if ".clean." in os.path.basename(src) and not args.output and not args.out_dir:
            continue  # do not clean our own output again when a folder is cleaned twice
        dst = output_path(src, args)
        if os.path.abspath(dst) == os.path.abspath(src):
            print("amagi-docleak: refusing to overwrite the input %s" % src, file=sys.stderr)
            return 2
        if os.path.exists(dst) and not args.force:
            print("amagi-docleak: %s exists (add --force to overwrite)" % dst, file=sys.stderr)
            return 2
        try:
            actions = clean_file(src, dst, keep=args.keep, unhide=args.unhide)
        except Exception as e:
            r = check_file(src)
            r.error = r.error or "clean failed: %s" % e
            from .model import Finding
            r.findings.insert(0, Finding("unreadable", "could not write a cleaned copy (%s)" % str(e)[:160]))
            reports.append(r)
            continue
        logs.append((src, dst, actions))
        reports.append(check_file(dst))
    _filter(reports, args.skip)
    if not args.json:
        for src, dst, actions in logs:
            print("cleaned %s -> %s" % (src, dst))
            for a in actions:
                print("  - " + a)
        if logs:
            print("\nchecking the cleaned copies:")
    return _emit(reports, args)


def main(argv=None):
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except (AttributeError, ValueError):
            pass
    argv = list(sys.argv[1:] if argv is None else argv)
    if argv and argv[0] not in COMMANDS and not argv[0].startswith("-"):
        argv.insert(0, "check")
    elif argv and argv[0].startswith("-") and argv[0] not in ("-h", "--help", "--version"):
        argv.insert(0, "check")
    parser = build_parser()
    args = parser.parse_args(argv)
    if args.command is None:
        parser.print_help()
        return 2
    if args.command == "checks":
        return cmd_checks()
    files = list(iter_files(args.paths))
    if not files:
        print("amagi-docleak: no .docx, .xlsx, .pptx or .pdf files found", file=sys.stderr)
        return 2
    if args.command == "clean":
        return cmd_clean(args, files)
    return _emit(_filter([check_file(f) for f in files], args.skip), args)


if __name__ == "__main__":
    sys.exit(main())
