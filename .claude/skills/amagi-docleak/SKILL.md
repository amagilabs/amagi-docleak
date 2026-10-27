---
name: amagi-docleak
description: Check a Word, Excel, PowerPoint or PDF file before it goes to opposing counsel, a client or a bidder - finds tracked changes, comments, hidden text, hidden sheets and rows, speaker notes, author and company metadata, file paths, cached pivot data, attachments, earlier PDF versions and fake black-box redactions - and write a cleaned copy. Runs locally. Use when asked to scrub metadata, check whether a redaction is real, make a document safe to send, or add a pre-send document check to a script or CI.
---

# amagi-docleak

A CLI that opens .docx, .xlsx, .pptx and .pdf files directly (zip + XML for Office, qpdf and
pdfminer for PDF) and reports what a recipient could find that the sender cannot see on screen.
Each finding has a severity: `fail` means content the author almost certainly did not mean to
send, `warn` means data worth a look, `info` is harmless but removable. `clean` writes a copy with
the fixable items removed and checks that copy. No network, no Office install needed.

## When to use this skill

- The user is about to send a document outside the firm or company and wants to know what it
  leaks (tracked changes, comments, reviewer names, hidden sheets, speaker notes, metadata).
- The user blacked something out and wants to know if the redaction is real, or wants a PDF whose
  black boxes cannot be copied around.
- The user wants a cleaned copy of a file, or a pass/fail gate in a script (`exit 1` on fail).

## How to use it

Install once: `pipx install amagi-docleak` (or `pip install -e .` from this repo in development).

```bash
amagi-docleak contract.docx                  # check one file (same as: amagi-docleak check contract.docx)
amagi-docleak outbox/                        # check every supported file in a folder
amagi-docleak check -v exhibit.pdf           # every detail line
amagi-docleak check --json outbox/ > report.json
amagi-docleak check --fail-on warn deck.pptx # warnings fail too
amagi-docleak check --skip thumbnail,author-metadata model.xlsx
amagi-docleak clean contract.docx            # writes contract.clean.docx, then checks it
amagi-docleak clean --out-dir to-send/ outbox/
amagi-docleak clean --keep comments draft.docx
amagi-docleak clean --unhide model.xlsx      # show hidden sheets/rows/text instead of keeping them hidden
amagi-docleak checks                         # every check, its severity and what clean does
```

Exit codes: 0 every file passed, 1 at least one file failed (at `--fail-on`, default `fail`),
2 usage error (bad option, nothing to check, output exists without `--force`).

## Privacy rules

- Text that was blacked out is masked in the report (first letter and length). `--reveal` prints
  it: use it only when the user asks, and never in a shared channel or a ticket.
- Other findings (deleted text, comments, hidden text, notes) are quoted as they are. Treat the
  report as being as confidential as the document.
- `clean` never touches the original. It refuses to write over the input and will not overwrite
  an existing output without `--force`.

## Working with the results

1. Run `check` and read the `fail` lines first. Tracked changes and comments are often meant to
   go: confirm with the user before cleaning a document they are still negotiating.
2. Run `clean` and read the actions list and the second report. Items `clean` leaves on purpose:
   hidden sheets, rows and slides (unless `--unhide`), data connections, macros, Office embedded
   objects, black shapes in PowerPoint, hyperlinks to other files. Fix those in the app.
3. A PDF with a fake or unapplied redaction is flattened page by page to a 200 dpi image. Those
   pages lose selectable text and accessibility tags. Say so to the user, and OCR the copy if they
   need searchable text.
4. A signed PDF is rewritten by `clean`, which breaks the signature. Clean before signing.
5. Open the cleaned copy in the app before sending it. Do not report a file as safe because the
   tool passed it; report what was checked and what was left.

## Limits

- Legacy .doc/.xls/.ppt files are not read (reported as `unsupported-format`). Save as the x
  format first.
- Password-protected files are reported as unreadable.
- Black shapes over text in PowerPoint and Word drawings are flagged (`possible-redaction-box`),
  but not removed: the tool cannot tell a deliberate dark design element from a cover-up.
- Text inside images is not read, so a scanned page with a sticker over a name looks clean.
