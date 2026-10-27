from amagi_docleak.util import contrast, find_paths, hex_to_rgb, is_blocked_out, mask, render, secret

BS = "\\"


def test_find_paths_windows_unc_and_unix():
    text = ("saved to C:" + BS + "Work" + BS + "a.docx and " + BS * 2 + "fs01" + BS + "share" + BS + "x.docx"
            " and /Users/jdoe/Desktop/x.txt")
    assert find_paths(text) == ["C:" + BS + "Work" + BS + "a.docx", BS * 2 + "fs01" + BS + "share" + BS + "x.docx",
                                "/Users/jdoe/Desktop/x.txt"]


def test_find_paths_escaped_backslashes_collapse():
    text = '{"path": "C:' + BS * 2 + "Work" + BS * 2 + 'a.docx"}'
    assert find_paths(text) == ["C:" + BS + "Work" + BS + "a.docx"]


def test_find_paths_file_url_wins_over_inner_path():
    assert find_paths("see file:///C:/Cases/brief.docx today") == ["file:///C:/Cases/brief.docx"]


def test_find_paths_ignores_web_urls():
    assert find_paths("https://example.com/a/b.docx") == []


def test_mask_and_reveal():
    s = "found " + secret("Halvard reserve") + " here"
    assert render(s) == 'found "H***********" (15 chars) here'
    assert render(s, reveal=True) == 'found "Halvard reserve" here'
    assert mask("") == '""'


def test_contrast_and_colors():
    black, white = hex_to_rgb("000000"), hex_to_rgb("FFFFFFFF")
    assert black == (0.0, 0.0, 0.0) and white == (1.0, 1.0, 1.0)
    assert round(contrast(black, white)) == 21
    assert contrast(black, black) == 1
    assert hex_to_rgb("zz") is None


def test_blocked_out():
    assert is_blocked_out("\u2588\u2588 \u2588")
    assert not is_blocked_out("\u2588a")
