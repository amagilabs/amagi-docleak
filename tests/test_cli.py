import json
import shutil

import pytest

from amagi_docleak.cli import main

from helpers import EXAMPLES, box, make_pdf, text_line


def run(capsys, *argv):
    code = main(list(argv))
    out, err = capsys.readouterr()
    return code, out, err


def test_bare_path_runs_check_and_fails(capsys):
    code, out, _ = run(capsys, str(EXAMPLES / "settlement-draft.docx"))
    assert code == 1
    assert "FAIL" in out and "tracked-changes" in out


def test_blacked_out_text_is_masked_unless_revealed(capsys):
    target = str(EXAMPLES / "settlement-draft.docx")
    _, out, _ = run(capsys, "check", "-v", target)
    assert "KF-2291" not in out
    assert '"H***********" (30 chars)' in out
    _, out, _ = run(capsys, "check", "-v", "--reveal", target)
    assert "Halvard Mutual, policy KF-2291" in out


def test_clean_file_passes(tmp_path, capsys):
    p = make_pdf(tmp_path / "ok.pdf", text_line(72, 700, "fine"))
    code, out, _ = run(capsys, str(p))
    assert code == 0 and "PASS" in out


def test_fail_on_and_skip(tmp_path, capsys):
    p = make_pdf(tmp_path / "f.pdf", text_line(72, 700, "secret") + box(70, 695, 100, 18))
    assert run(capsys, str(p))[0] == 1
    assert run(capsys, "--fail-on", "never", str(p))[0] == 0
    assert run(capsys, "--skip", "fake-redaction", str(p))[0] == 0


def test_json_report(capsys):
    code, out, _ = run(capsys, "check", "--json", str(EXAMPLES / "exhibit-b.pdf"))
    data = json.loads(out)
    assert code == 1 and data["failed"] == 1
    f = data["files"][0]
    assert f["verdict"] == "FAIL" and f["format"] == "pdf"
    checks = {x["check"] for x in f["findings"]}
    assert "unapplied-redaction" in checks
    assert "Dana Reyes" not in out  # masked in JSON too


def test_unknown_check_name_is_usage_error(capsys):
    with pytest.raises(SystemExit) as e:
        main(["check", "--skip", "nope", str(EXAMPLES)])
    assert e.value.code == 2


def test_checks_lists_every_check(capsys):
    code, out, _ = run(capsys, "checks")
    assert code == 0
    for cid in ("tracked-changes", "fake-redaction", "speaker-notes", "prior-versions", "unreadable"):
        assert cid in out


def test_old_binary_formats_are_reported(tmp_path, capsys):
    p = tmp_path / "old.doc"
    p.write_bytes(b"\xd0\xcf\x11\xe0 not really")
    code, out, _ = run(capsys, str(p))
    assert code == 1 and "unsupported-format" in out


def test_clean_command_writes_copy_and_checks_it(tmp_path, capsys):
    src = tmp_path / "deck.pptx"
    shutil.copy(EXAMPLES / "strategy-deck.pptx", src)
    code, out, _ = run(capsys, "clean", str(src))
    assert (tmp_path / "deck.clean.pptx").exists()
    assert "removed speaker notes" in out
    assert code == 0  # what is left on the deck is warnings only
    # second run refuses to overwrite without --force
    code, _, err = run(capsys, "clean", str(src))
    assert code == 2 and "--force" in err
    assert run(capsys, "clean", "--force", str(src))[0] == 0


def test_clean_refuses_to_overwrite_input(tmp_path, capsys):
    src = tmp_path / "a.pdf"
    shutil.copy(EXAMPLES / "letter-to-counsel.pdf", src)
    code, _, err = run(capsys, "clean", "-o", str(src), str(src))
    assert code == 2 and "refusing" in err


def test_clean_output_needs_single_input(tmp_path, capsys):
    code, _, err = run(capsys, "clean", "-o", str(tmp_path / "x.pdf"),
                       str(EXAMPLES / "exhibit-b.pdf"), str(EXAMPLES / "letter-to-counsel.pdf"))
    assert code == 2 and "--out-dir" in err


def test_clean_folder_to_out_dir(tmp_path, capsys):
    out_dir = tmp_path / "out"
    code, out, _ = run(capsys, "clean", "--unhide", "--out-dir", str(out_dir), str(EXAMPLES))
    names = sorted(p.name for p in out_dir.iterdir())
    assert names == ["damages-model.clean.xlsx", "exhibit-b.clean.pdf", "letter-to-counsel.clean.pdf",
                     "settlement-draft.clean.docx", "strategy-deck.clean.pptx"]
    assert code == 0


def test_missing_path_fails(tmp_path, capsys):
    code, out, _ = run(capsys, str(tmp_path / "nope.docx"))
    assert code == 1 and "file not found" in out
