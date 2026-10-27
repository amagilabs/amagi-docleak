"""Small helpers shared by the checkers: colors, text excerpts, path detection."""

import re

BLOCKS = set("\u2588\u2587\u25a0\u25ae ")


def is_blocked_out(text):
    """True when text is already redacted with block characters (or is blank)."""
    return all(ch in BLOCKS or ch.isspace() for ch in text)


def excerpt(text, limit=70):
    text = " ".join(text.split())
    if len(text) > limit:
        text = text[: limit - 1].rstrip() + "\u2026"
    return text


def quoted(text, limit=70):
    return '"' + excerpt(text, limit) + '"'


def mask(text):
    """Keep the first character and the length so the reader can find the spot."""
    text = " ".join(text.split())
    if not text:
        return '""'
    shown = text[0] + "*" * min(len(text) - 1, 11)
    return '"%s" (%d chars)' % (shown, len(text))


_SECRET = re.compile("(.*?)", re.S)


def secret(text):
    """Mark text the author meant to hide; the report masks it unless --reveal is given."""
    return "%s" % " ".join(text.split())


def render(s, reveal=False):
    return _SECRET.sub(lambda m: quoted(m.group(1), 200) if reveal else mask(m.group(1)), s)


def plural(n, word, many=None):
    return "%d %s" % (n, word if n == 1 else (many or word + "s"))


def join_names(names, limit=3):
    names = [n for n in dict.fromkeys(names) if n]
    if not names:
        return ""
    if len(names) > limit:
        return ", ".join(names[:limit]) + " and %d more" % (len(names) - limit)
    return ", ".join(names)


# ---- colors -------------------------------------------------------------------------------

def hex_to_rgb(value):
    if not value:
        return None
    value = value.strip().lstrip("#")
    if len(value) == 8:  # ARGB in spreadsheets
        value = value[2:]
    if len(value) != 6:
        return None
    try:
        return tuple(int(value[i:i + 2], 16) / 255.0 for i in (0, 2, 4))
    except ValueError:
        return None


def luminance(rgb):
    def lin(c):
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = rgb
    return 0.2126 * lin(r) + 0.7152 * lin(g) + 0.0722 * lin(b)


def contrast(rgb1, rgb2):
    a, b = luminance(rgb1), luminance(rgb2)
    hi, lo = max(a, b), min(a, b)
    return (hi + 0.05) / (lo + 0.05)


def is_dark(rgb):
    return luminance(rgb) < 0.05


UNREADABLE = 1.6  # contrast ratio below this: the text cannot be read on its background


# ---- file paths ---------------------------------------------------------------------------

_PATH_RES = [
    re.compile(r"(?<![A-Za-z0-9])[A-Za-z]:(?:\\{1,2}|/)(?:[^<>\"|?*\r\n\t\\/:]+(?:\\{1,2}|/))+[^<>\"|?*\r\n\t\\/:]*"),
    re.compile(r"\\{2,4}[A-Za-z0-9._$-]+(?:\\{1,2}[^<>\"|?*\r\n\t\\:]+)+"),
    re.compile(r"file:/{2,3}[^\s<>\"']+", re.IGNORECASE),
    re.compile(r"(?<![A-Za-z0-9.])/(?:Users|home)/[A-Za-z0-9._-]+(?:/[^\s<>\"'/]+)*"),
]


def find_paths(text):
    out = []
    for rx in _PATH_RES:
        for m in rx.finditer(text):
            value = m.group(0).rstrip(" .,;)")
            value = value.replace("\\\\", "\\") if not value.startswith("\\\\\\") else value
            if len(value) > 4:
                out.append(value)
    # a match from one pattern can contain another (file:///C:/...), keep the longest
    out.sort(key=len, reverse=True)
    kept = []
    for v in out:
        if not any(v in k for k in kept):
            kept.append(v)
    return kept


def is_local_target(target):
    """A relationship target that points at a file rather than a web page."""
    t = target.strip()
    low = t.lower()
    if low.startswith(("http://", "https://", "mailto:", "#", "tel:", "ftp://")):
        return False
    return bool(low.startswith("file:") or re.match(r"^[A-Za-z]:[\\/]", t) or t.startswith("\\\\")
                or t.startswith("/") or ".." in t or re.search(r"\.(docx?|xlsx?|xlsm|pptx?|pdf|dotx?|dotm|csv|txt)$", low))
