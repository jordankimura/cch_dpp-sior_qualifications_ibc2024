"""
Shared by the harness and by anyone writing evidence entries, so both sides normalise text
the same way. A quote "is in" a source when the normalised quote is a substring of the
normalised source.

Normalisation only removes differences that come from extraction, never from wording:
  - typographic quotes, dashes and non-breaking spaces become plain ASCII
  - letters split across a line break by a hyphen are tried both joined ("per-formed" and
    "performed"), because a PDF cannot tell a soft hyphen from a real one
  - all runs of whitespace become one space
"""
import json, os, re, unicodedata

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))   # verification/
SOURCES = os.path.join(HERE, "sources")

_TRANS = str.maketrans({"‘": "'", "’": "'", "“": '"', "”": '"', "–": "-", "—": "-",
                        "−": "-", " ": " ", "‑": "-", "�": "-"})


def norm(s, join_hyphens=False):
    s = unicodedata.normalize("NFKC", s or "").translate(_TRANS)
    s = re.sub(r"-\s*\n\s*", "" if join_hyphens else "-", s)
    return re.sub(r"\s+", " ", s).strip()


_cache = {}


def manifest():
    if "manifest" not in _cache:
        _cache["manifest"] = json.load(open(os.path.join(SOURCES, "manifest.json"), encoding="utf8"))
    return _cache["manifest"]


def source_entry(sid):
    for s in manifest()["sources"]:
        if s["id"] == sid: return s
    raise KeyError("unknown source id: " + sid)


def source_text(sid, joined=False):
    key = (sid, joined)
    if key not in _cache:
        e = source_entry(sid)
        if not e.get("file"): raise KeyError("source %s is not held (no file)" % sid)
        raw = open(os.path.join(SOURCES, e["file"]), encoding="utf8", errors="replace").read()
        raw = raw.split("\n\n", 1)[1] if raw.startswith("# ") else raw     # drop the provenance header
        _cache[key] = norm(raw, join_hyphens=joined)
    return _cache[key]


HEADING = re.compile(r"^(?:\[[A-Z]+\]\s*)?(1[0-9]{3}(?:\.\d+)*)\s+\S", re.M)


def section_span(sid, section):
    """The text from a section's heading to the next heading, for ICC chapter files where each
    heading sits on its own line. Returns normalised text, or None if the heading is absent."""
    e = source_entry(sid)
    raw = open(os.path.join(SOURCES, e["file"]), encoding="utf8", errors="replace").read()
    heads = [(m.start(), m.group(1)) for m in HEADING.finditer(raw)]
    for i, (pos, num) in enumerate(heads):
        if num == section:
            end = heads[i + 1][0] if i + 1 < len(heads) else len(raw)
            return norm(raw[pos:end])
    return None


def headings(sid):
    """{section number: heading line} for an ICC chapter file."""
    e = source_entry(sid)
    raw = open(os.path.join(SOURCES, e["file"]), encoding="utf8", errors="replace").read()
    out = {}
    for line in raw.splitlines():
        m = re.match(r"^(?:\[[A-Z]+\]\s*)?(1[0-9]{3}(?:\.\d+)*)\s+(.+?)\s*$", line)
        if m and m.group(1) not in out: out[m.group(1)] = norm(m.group(2))
    return out


def contains(sid, quote, within=None, match="exact"):
    """True if the quote occurs in the source (or within one section's span).
    match="nospace" ignores ALL whitespace on both sides; only allowed for sources whose
    catalog entry is 'partial' (a garbled scan), and the harness enforces that."""
    if match == "nospace":
        return re.sub(r"\s", "", norm(quote)) in re.sub(r"\s", "", source_text(sid))
    q1, q2 = norm(quote), norm(quote, join_hyphens=True)
    if within:
        span = section_span(sid, within)
        return bool(span) and (q1 in span or q2 in norm(span, True))
    return q1 in source_text(sid) or q2 in source_text(sid, joined=True)


def find(sid, pattern, width=0):
    """Authoring helper: return the exact normalised passage matching a regex, so a quote is
    copied from the source rather than typed from memory."""
    t = source_text(sid)
    m = re.search(pattern, t, re.S)
    if not m:
        t = source_text(sid, joined=True); m = re.search(pattern, t, re.S)
    if not m: raise LookupError("%s: pattern not found: %s" % (sid, pattern))
    return t[m.start():m.end() + width].strip()
