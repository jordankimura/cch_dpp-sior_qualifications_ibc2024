#!/usr/bin/env python3
"""
Renders sources/SOURCES.md from sources/manifest.json.

SOURCES.md is a GENERATED FILE. Do not hand-edit it: edit tools/extract_sources.py (which writes
the manifest), re-run it, then re-run this. Check C-01 in verify_si.py fails if they drift.

    python3 gen_sources_md.py            # rewrite sources/SOURCES.md
    python3 gen_sources_md.py --check    # exit 1 if it is out of date
"""
import argparse, json, os, sys

HERE = os.path.dirname(os.path.abspath(__file__))
MANIFEST = os.path.join(HERE, "sources", "manifest.json")
OUTPUT = os.path.join(HERE, "sources", "SOURCES.md")


def esc(s): return (s or "").replace("|", "\\|").replace("\n", " ")


def render(man):
    S = man["sources"]
    held = [s for s in S if s.get("file")]; missing = [s for s in S if s["kind"] == "not-held"]
    L = []; a = L.append
    a("# Sources")
    a("")
    a("<!-- GENERATED FROM manifest.json BY gen_sources_md.py — DO NOT HAND-EDIT. -->")
    a("")
    a("Compiled %s. **%d documents held as full text in this repository, %d needed but not held.**" % (man["compiled"], len(held), len(missing)))
    a("")
    a("Every quote on the workbook's Evidence tab is checked against the text files listed here. "
      "Each file records the fingerprint (sha256) of the original it was extracted from; the originals "
      "themselves are not committed. Folder `licensed/` holds copyrighted codes and standards: keep this "
      "repository private (see LICENSING.md).")
    a("")
    for kind, heading in [("code", "Building codes"), ("standard", "Referenced standards"), ("ordinance", "City & County ordinances"),
                          ("state-rule", "State of Hawaii"), ("accreditation", "Accreditation criteria"), ("issuer-page", "Issuer pages (dated snapshots)")]:
        rows = [s for s in held if s["kind"] == kind]
        if not rows: continue
        a("## " + heading); a("")
        a("| ID | Document | Edition | Complete | Use | File | Limitation |")
        a("|---|---|---|---|---|---|---|")
        for s in rows:
            a("| `%s` | %s | %s | %s | %s | `%s` | %s |" % (s["id"], esc(s["title"]), esc(s["edition"]), s["completeness"], s["use"],
                                                         s["file"], esc(s.get("limitation") or s.get("context_reason") or "")))
        a("")
    a("## Needed but not held"); a("")
    a("| ID | Document | Why not | What it leaves open |"); a("|---|---|---|---|")
    for s in missing:
        a("| `%s` | %s | %s | %s |" % (s["id"], esc(s["title"]), esc(s["reason"]), esc(s["blocks"])))
    a("")
    a("## Field notes"); a("")
    for k, v in man["field_notes"].items(): a("- **%s**: %s" % (k, v))
    return "\n".join(L) + "\n"


def main():
    ap = argparse.ArgumentParser(); ap.add_argument("--check", action="store_true"); a = ap.parse_args()
    man = json.load(open(MANIFEST, encoding="utf-8")); want = render(man)
    if a.check:
        have = open(OUTPUT, encoding="utf-8").read() if os.path.exists(OUTPUT) else ""
        if have.strip() != want.strip(): print("SOURCES.md is out of date - run gen_sources_md.py"); sys.exit(1)
        print("SOURCES.md is in sync"); return
    open(OUTPUT, "w", encoding="utf-8", newline="\n").write(want); print("wrote", OUTPUT)


if __name__ == "__main__":
    main()
