#!/usr/bin/env python3
"""
Writes the GENERATED parts of the workbook from the files that own them, so they cannot drift:

  README tab                       <- README_ROWS below
  Evidence tab                     <- verification/evidence.json
  Sources tab                      <- verification/sources/manifest.json
  "Evidence" column on Matrix 2024 (R), Flags, Decisions, Credential check, Change Log
                                   <- the used_by lists in evidence.json
  verification/workbook.txt        <- a plain-text dump of every tab, so git shows cell-level diffs

    python3 tools/sync_workbook.py WORKBOOK            # rewrite the generated parts
    python3 tools/sync_workbook.py WORKBOOK --check    # exit 1 if anything is out of date

Check S-01 in verify_si.py runs the same comparison.
"""
import argparse, json, os, re, sys
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
from xlsx import Workbook, table_sheet, read_all, col_number, split_ref, STYLE_HEADER, STYLE_BODY

VER = os.path.dirname(HERE)
EVIDENCE = os.path.join(VER, "evidence.json")
MANIFEST = os.path.join(VER, "sources", "manifest.json")
DUMP = os.path.join(VER, "workbook.txt")

# tab -> (column holding the row key, column for evidence IDs)
KEYED = {"Flags": ("A", "H"), "Decisions": ("A", "I"), "Credential check": ("A", "H"), "Change Log": ("A", "E"), "Matrix 2024": ("B", "R")}

README_ROWS = [
 ("Status", "DRAFT, not adopted. Honolulu adopts the Hawaii State Building Code of 20 April 2021, which adopts the 2018 IBC (ROH 16-1.1). "
            "This workbook updates DPP's special inspector qualification matrix to the 2024 IBC, ahead of any adoption."),
 ("Start here", "Matrix 2024 is the deliverable: one row per special inspection classification. Use the filter on the Flag column to see "
                "what needs attention, and the Decisions tab for what only DPP can settle."),
 ("Matrix 2024", "The deliverable. Columns H to K: who may perform the inspection. Columns L to P: minimum certification, copied unchanged "
                 "from the 2018 sheet. Q: where the credential requirement comes from. R: Evidence IDs."),
 ("Flags", "The note behind every flagged or checked row, keyed by Ref."),
 ("Decisions", "Questions only DPP can answer. Columns F to H are for DPP's answer, initials and date."),
 ("Credential check", "Each certification checked against its issuer or the referenced standard."),
 ("Evidence", "What backs each claim: the document, the section, and a short quote. Every quote is checked automatically against the full "
              "text of the document."),
 ("SI FORM dropdown", "What changed in the SI FORM pick-list, and what was held back."),
 ("SI FORM, Drop-down, Footer, blurb", "DPP's original tabs. SI FORM's text is untouched; only its dropdown ranges were extended. "
                                       "Drop-down column C maps each pick-list entry to matrix rows."),
 ("SI Qual 2018 (as received)", "DPP's 2018 sheet, cell-for-cell as received (renamed only because Excel limits tab names to 31 characters)."),
 ("Change Log", "What changed from the 2018 sheet, and every repair or correction since."),
 ("Sources", "Every document relied on, with its edition and completeness, plus the documents that could not be obtained and what "
             "each leaves open."),
 ("Flags, in one line", "COLLISION = same section number used by City and base code. CONFLICT = the row permits something a code, standard or "
                        "ordinance appears to restrict. GAP = 2024 section with no row. check = needs DPP confirmation. policy = department "
                        "policy. Blank = checked, nothing found."),
 ("Evidence IDs", "E-### points to a row on the Evidence tab. Types: supports (the quote is in the document); absent (the terms do not "
                  "occur in a complete document or section); internal (the quote is a cell of DPP's 2018 sheet); open (the document is "
                  "not held, see Sources)."),
 ("How this is checked", "The workbook is kept in a private git repository with an automated check script (verification/verify_si.py) "
                         "that runs on every change: the 2018 data and SI FORM are untouched, every citation and title matches the code "
                         "text, every flag and credential basis has evidence, every quote is found in its source, and the file opens "
                         "in Excel's format rules."),
]


def load():
    return json.load(open(EVIDENCE, encoding="utf8")), json.load(open(MANIFEST, encoding="utf8"))


def evidence_rows(ev):
    out = []
    for e in ev["evidence"]:
        q = e.get("quote") or ("absent: " + ", ".join(e.get("terms", [])) + ((" (within " + e["within"] + ")") if e.get("within") else "")
                              if e["type"] == "absent" else "")
        out.append([e["id"], e["claim"], e["type"], e["source"], e.get("where", ""), q, "; ".join(e["used_by"])])
    return out


def source_rows(man):
    rows = []
    for s in man["sources"]:
        if s["kind"] == "not-held": continue
        rows.append([s["id"], s["title"], s["issuer"], s["edition"], s["completeness"], s["use"],
                     s.get("limitation") or s.get("context_reason") or ""])
    notes = ["NOT HELD — each of these was needed and could not be obtained:"]
    for s in man["sources"]:
        if s["kind"] == "not-held":
            notes.append("%s — %s (%s). Leaves open: %s" % (s["id"], s["title"], s["reason"], s["blocks"]))
    return rows, notes


def ids_by_key(ev):
    """{(tab, key): [E-ids]}"""
    m = {}
    for e in ev["evidence"]:
        for u in e["used_by"]:
            parts = u.split(":")
            m.setdefault((parts[0], parts[1]), [])
            if e["id"] not in m[(parts[0], parts[1])]: m[(parts[0], parts[1])].append(e["id"])
    return m


def build(path, write=True):
    ev, man = load()
    w = Workbook(path)
    names = [s[0] for s in w.sheets()]
    readme = table_sheet("README — how to read this workbook", ["Topic", "What it is"], [list(r) for r in README_ROWS], [30, 140], freeze_row=0)
    if "README" in names: w.replace_sheet("README", readme)
    else: w.add_sheet("README", readme, 0)
    evx = table_sheet("EVIDENCE — what backs each claim. Quotes are checked against the full text of each source (verification/sources).",
                      ["ID", "Claim", "Type", "Source", "Where", "Quote or terms", "Used by"], evidence_rows(ev), [7, 48, 9, 20, 18, 70, 40], autofilter=True)
    if "Evidence" in names: w.replace_sheet("Evidence", evx)
    else: w.add_sheet("Evidence", evx, len([s for s in w.sheets()]))
    srows, snotes = source_rows(man)
    w.replace_sheet("Sources", table_sheet("SOURCES — every document relied on. Generated from verification/sources/manifest.json.",
                    ["ID", "Document", "Issuer", "Edition", "Complete?", "Use", "Limitation"], srows, [22, 60, 16, 22, 11, 9, 60],
                    note_rows=snotes, merge_notes=True))
    cur = read_all_parts(w)
    idx = ids_by_key(ev)
    for tab, (kcol, ecol) in KEYED.items():
        cells = cur[tab]; setc = {}
        rows = sorted({int(split_ref(r)[1]) for r in cells})
        for r in rows:
            if r < 3: continue
            key = cells.get("%s%d" % (kcol, r))
            if not key: continue
            if tab == "Matrix 2024" and not re.match(r"^(SI|NEW)-\d+$", key): continue
            want = ", ".join(idx.get((tab, key), []))
            if (cells.get("%s%d" % (ecol, r)) or "") != want: setc["%s%d" % (ecol, r)] = want
        if tab == "Matrix 2024" and cells.get("R1") != "Evidence":
            setc["R1"] = "Evidence"
        if setc: w.set_cells(tab, setc)
    fix_matrix_column_r(w)
    if write:
        w.save(path)
        open(DUMP, "w", encoding="utf8", newline="\n").write(dump(path))
    return w


def read_all_parts(w):
    import tempfile
    fd, tmp = tempfile.mkstemp(suffix=".xlsx"); os.close(fd)
    try:
        w.save(tmp); return read_all(tmp)
    finally:
        os.remove(tmp)


def fix_matrix_column_r(w):
    """Give Matrix 2024 column R a width, header style, merged header, and include it in the filter."""
    x = w.xml("Matrix 2024")
    if '<col min="18" max="18"' not in x:
        x = x.replace("</cols>", '<col min="18" max="18" width="22" customWidth="1"/></cols>')
    q1 = re.search(r'<c r="Q1" s="(\d+)"', x)
    if q1: x = re.sub(r'<c r="R1" s="\d+"', '<c r="R1" s="%s"' % q1.group(1), x)
    q2 = re.search(r'<c r="Q2" s="(\d+)"', x)
    if q2 and '<c r="R2"' not in x:
        x = re.sub(r'(<row r="2"[^>]*>.*?)(</row>)', lambda m: m.group(1) + '<c r="R2" s="%s"/>' % q2.group(1) + m.group(2), x, count=1, flags=re.S)
    if '<mergeCell ref="R1:R2"/>' not in x:
        x = re.sub(r'<mergeCells count="(\d+)">', lambda m: '<mergeCells count="%d"><mergeCell ref="R1:R2"/>' % (int(m.group(1)) + 1), x)
    x = re.sub(r'<autoFilter ref="A2:Q(\d+)"', r'<autoFilter ref="A2:R\1"', x)
    x = re.sub(r'<dimension ref="A1:Q(\d+)"', r'<dimension ref="A1:R\1"', x)
    for r, s in re.findall(r'<c r="R(\d+)" s="(\d+)"', x):
        pass
    w.set_xml("Matrix 2024", x)


def dump(path):
    d = read_all(path); out = []
    order = [s[0] for s in Workbook(path).sheets()]
    for tab in order:
        out.append("########## %s" % tab)
        for ref in sorted(d[tab], key=lambda r: (split_ref(r)[1], col_number(split_ref(r)[0]))):
            out.append("%s: %s" % (ref, d[tab][ref].replace("\n", " / ")))
    return "\n".join(out) + "\n"


def check(path):
    """Return a list of differences between the workbook and what sync would write."""
    import tempfile
    fd, tmp = tempfile.mkstemp(suffix=".xlsx"); os.close(fd)
    try:
        w = build(path, write=False); w.save(tmp)
        want, have = read_all(tmp), read_all(path)
        diffs = []
        for tab in set(want) | set(have):
            if tab not in have: diffs.append("tab missing: " + tab); continue
            if tab not in want: continue
            for ref in set(want[tab]) | set(have[tab]):
                if want[tab].get(ref) != have[tab].get(ref): diffs.append("%s!%s" % (tab, ref))
        if not os.path.exists(DUMP) or open(DUMP, encoding="utf8").read() != dump(path):
            diffs.append("verification/workbook.txt is out of date")
        return sorted(diffs)
    finally:
        os.remove(tmp)


if __name__ == "__main__":
    ap = argparse.ArgumentParser(); ap.add_argument("workbook"); ap.add_argument("--check", action="store_true"); a = ap.parse_args()
    if a.check:
        d = check(a.workbook)
        print("in sync" if not d else "OUT OF DATE: %d cells, e.g. %s" % (len(d), ", ".join(d[:8]))); sys.exit(1 if d else 0)
    build(a.workbook); print("synced")
