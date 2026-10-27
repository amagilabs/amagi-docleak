# amagi-docleak

Check a Word, Excel, PowerPoint or PDF file before it leaves the office. amagi-docleak lists what the recipient could dig out that you can't see on screen: tracked changes, comments, hidden sheets and rows, speaker notes, author and company metadata, file paths, and black boxes that were drawn over text instead of removing it. It gives each file a PASS or FAIL, and `clean` writes a copy with that material removed.

![amagi-docleak checking a settlement draft, then cleaning a PDF exhibit with fake redactions](https://raw.githubusercontent.com/amagilabs/amagi-docleak/master/assets/demo.gif)

Office's Inspect Document only works one file at a time, inside Office, and it never looks at PDFs. Commercial metadata scrubbers do this job well, but they are paid add-ins, usually set up through Outlook or a document management system. exiftool and mat2 are good at metadata, but they won't tell you that the settlement figure under a black box can still be copied. amagi-docleak is a free command: point it at a file or a folder, read the report, and get exit code 1 when something should not go out.

It is built for people who send documents to the other side: law firms, consultants, contractors bidding on work, and anyone who has redacted a PDF by drawing a black rectangle.

## Install

```
pipx install amagi-docleak
```

Python 3.9 or newer. It runs offline and does not need Office installed. It reads the files directly with lxml, qpdf (through pikepdf) and pdfminer.six.

Not on PyPI yet. Until the first release, install from a clone with `pip install .` in the repo folder.

## Usage

### Check a file

The [examples](examples) folder holds five made-up files with the usual leaks planted in them. Here is the settlement draft:

```
$ amagi-docleak examples/settlement-draft.docx
FAIL  examples/settlement-draft.docx  (4 fails, 4 warns, 1 note)
  fail  tracked-changes     3 tracked changes by Sam Okafor, Jordan Ellery (2 deletions, 1 insertion)
                          deleted: "$410,000"
                          deleted: "Our client will not go below $350,000."
                          inserted: "$385,000"
                          ... 1 more (use --verbose)
  fail  comments            1 comment by Sam Okafor
                          Sam Okafor: "Check the carrier name with Jordan before this goes out."
  fail  fake-redaction      1 passage blacked out with highlight or shading; the text is still there and can be copied
                          blacked out: "H***********" (30 chars)
  fail  hidden-text         2 passages of text the reader cannot see (hidden, same color as its background)
                          hidden: "Internal: Sam thinks the arbitrator leans our way."
                          same color as its background: "Draft 7 - do not circulate outside the firm."
  warn  document-variables  2 document variables: MatterNo, ReviewedBy
                          MatterNo = "BRW-0412"
                          ReviewedBy = "Sam Okafor"
  warn  author-metadata     author "Jordan Ellery", last modified by "Sam Okafor", company "Harlow & Pike LLP"
                          author: Jordan Ellery
                          last modified by: Sam Okafor
                          title: Brightwater Supply v. Kestrel Freight - settlement terms
                          ... 6 more (use --verbose)
  warn  custom-properties   2 custom properties (ClientMatter, DMSDocID)
                          ClientMatter: BRW-0412
                          DMSDocID: HPDMS-88213 v7
  warn  file-paths          2 file paths stored in the file, e.g. "D:\Work\jellery\Brightwater\call-notes-0911.docx"
                          D:\Work\jellery\Brightwater\call-notes-0911.docx  (hyperlink link in word/_rels/document.xml.rels)
                          file:///\\fs01\templates\HarlowPike-Settlement.dotx  (attachedTemplate link in word/_rels/settings.xml.rels)
  info  thumbnail           preview image of the first page saved in the file (docProps/thumbnail.jpeg)

Blacked-out and covered text is masked. Add --reveal to print it.
```

`fail` means something the author almost certainly did not mean to send. `warn` is worth a look before it goes out. `info` is harmless but removable. Text that someone blacked out or covered is masked to its first letter and length, so the report doesn't repeat what you meant to hide. Add `--reveal` to see it. Everything else (deleted text, comments, hidden text, notes) is quoted as it is, so keep the report as private as the document.

Give it a folder to check everything inside (`.docx`, `.xlsx`, `.pptx`, `.pdf`):

```
$ amagi-docleak examples -q
FAIL  examples/damages-model.xlsx  (3 fails, 6 warns)
FAIL  examples/exhibit-b.pdf  (3 fails, 6 warns)
FAIL  examples/letter-to-counsel.pdf  (1 fail)
FAIL  examples/settlement-draft.docx  (4 fails, 4 warns, 1 note)
FAIL  examples/strategy-deck.pptx  (2 fails, 4 warns, 1 note)
5 files checked: 5 failed, 0 passed
```

`-v` prints every detail line, `--json` prints a JSON report, `--fail-on warn` makes warnings fail too, and `--skip thumbnail,author-metadata` ignores the checks you don't care about.

### Fake redactions in PDFs

`letter-to-counsel.pdf` was exported from Word with a black highlight over an insurer's name and claim number. It looks redacted, but anyone can select the line and copy it. `exhibit-b.pdf` has a bit of everything: black rectangles drawn over a driver's name and phone number and over the insurer's name, a Redact mark over an account number that was never applied, a white box over the claims reserve, a note to a colleague, an attached file, a script that runs on open, and an earlier save in which a line about the settlement floor is still readable.

```
$ amagi-docleak examples/exhibit-b.pdf
FAIL  examples/exhibit-b.pdf  (3 fails, 6 warns)
  fail  comments             1 comment by Sam Okafor
                           page 1, text by Sam Okafor: "Check with Jordan before we produce this page."
  fail  fake-redaction       2 black boxes drawn over text that can still be selected and copied
                           page 1: black box over "D***********" (29 chars)
                           page 1: black box over "K***********" (28 chars)
  fail  unapplied-redaction  redaction marks on 1 page were never applied
                           page 1: 1 area marked, text still underneath: "A***********" (38 chars)
  warn  author-metadata      author "Jordan Ellery", company "Harlow & Pike LLP"
                           Author: Jordan Ellery
                           Company: Harlow & Pike LLP
                           Title: Exhibit B - delivery log
                           ... 2 more (use --verbose)
  warn  custom-properties    1 custom property (SourceModified)
                           SourceModified: D:20260912100500
  warn  covered-text         1 passage of text drawn over by a white box or an image
                           page 1: "R***********" (26 chars)
  warn  embedded-files       1 attached file: call-notes.txt
                           call-notes.txt (1 KB)
  warn  javascript           1 script or launch action: JavaScript action when the file opens
                           JavaScript action when the file opens
  warn  prior-versions       1 earlier version saved before the current one, still inside the file
                           revision 1: 1 line no longer in the current version, e.g. "S***********" (36 chars)

Blacked-out and covered text is masked. Add --reveal to print it.
```

To find these, it reads each page in drawing order and compares every filled shape with the characters drawn before it. Black text on a black box counts. White text on a black label does not, and neither does the invisible text layer of a scanned page.

### Write a cleaned copy

```
$ amagi-docleak clean examples/exhibit-b.pdf -o exhibit-b.clean.pdf
cleaned examples/exhibit-b.pdf -> exhibit-b.clean.pdf
  - cleared document info fields (Author, Company, SourceModified, Title)
  - removed the XMP metadata packet
  - removed 1 comment
  - removed 1 attached file
  - removed 1 script or launch action
  - replaced page 1 with a 200 dpi image: the black boxes are burned in and the text under them is gone (that page can no longer be searched or copied from)
  - rewrote the file, which drops the earlier versions saved inside it

checking the cleaned copies:
PASS  exhibit-b.clean.pdf  (nothing found)
```

`clean` never changes the original. It writes `name.clean.ext` next to it (or `-o FILE`, or `--out-dir DIR` for a folder), refuses to overwrite anything without `--force`, and then checks the copy so you can see what is left. For Word files it accepts every tracked change, removes comments, document variables and the template path, turns blacked-out runs into solid blocks, and deletes hidden text. In Excel it removes comments, cached pivot data and the folder the workbook was last saved in, and replaces blacked-out cells with blocks. In PowerPoint it removes speaker notes and comments. All three Office formats lose their author and company properties, custom properties and the preview thumbnail.

Some things are left for you to decide. Hidden sheets, rows and slides stay hidden unless you add `--unhide`, which makes them visible. Hidden text in Word is deleted, or made visible with `--unhide`. `--keep comments,author-metadata` leaves those checks alone. Run `amagi-docleak checks` to see every check, its severity, and what `clean` does about it.

Open the cleaned copy in Word, Excel or Acrobat before you send it.

### Exit codes

`0` every file passed, `1` at least one file failed at the `--fail-on` level (default `fail`), `2` usage error (unknown option or check name, nothing to check, output file exists). That makes it usable as a last step in a script that builds a production set or a bid package.

## What it checks

| Format | Checks |
|---|---|
| Word `.docx` | tracked changes (and Track Changes left on), comments, text that is hidden, 1 pt or the same color as its background, black highlight or shading over text, black shapes over text, document variables, template and file paths, metadata, custom properties, embedded objects, macros, thumbnail |
| Excel `.xlsx` | hidden and very hidden sheets, hidden rows and columns that hold data, `;;;` and font-matches-fill cells, black-on-black cells, comments and threaded comments, pivot caches holding source rows, data connections (passwords masked), last save folder, metadata, custom properties, embedded objects, macros, thumbnail |
| PowerPoint `.pptx` | speaker notes, comments, hidden slides, shapes parked off the slide, black boxes over text, file paths, metadata, custom properties, embedded objects, macros, thumbnail |
| PDF | black boxes and black annotations over live text, Redact marks never applied, text under white boxes or images, comments, attachments, JavaScript and launch actions, earlier versions kept by incremental saves, layers switched off, application private data, Info and XMP metadata, file paths |

Legacy `.doc`, `.xls` and `.ppt` files fail with `unsupported-format`; save them in the newer format and check that copy.

## Use it from Claude Code

[`.claude/skills/amagi-docleak/SKILL.md`](.claude/skills/amagi-docleak/SKILL.md) teaches Claude Code when to run the check, how to read the report, and to keep `--reveal` output out of shared places. Copy the folder into your project's `.claude/skills/` or into `~/.claude/skills/`.

## Limits

- Flattening a PDF page turns it into an image. That page loses selectable text and its accessibility tags, so run OCR on the clean copy if it has to stay searchable. Pages without a bad redaction are left as they were.
- `clean` rewrites a PDF, which breaks any digital signature on it. Clean first, then sign.
- Black shapes over text in Word and PowerPoint are reported but not removed, because a dark design element and a cover-up look the same to a parser. Data connections, macros, embedded Office objects, hidden cells and hyperlinks to other files are reported and left in place too.
- Text inside images is not read. A scanned page with a sticker over a name will pass.
- Password-protected files are reported as unreadable. Remove the password, check, and add it back.
- It looks for the leaks listed above and nothing else. A PASS does not certify that a file is safe to publish.

## If this saved you time

A star helps other people find it. If it missed a leak or flagged something that wasn't one, open an [issue](https://github.com/amagilabs/amagi-docleak/issues) with a made-up sample file that shows it.

## More from AmagiLabs

[github.com/amagilabs](https://github.com/amagilabs) has the other small open source tools from the same lab.

## License

MIT, see [LICENSE](LICENSE). Copyright (c) 2026 Amagisolutionsgroup, Inc. (AmagiLabs)
