#!/usr/bin/env python3
"""
Renders SOURCES.md from manifest.json.

SOURCES.md is a GENERATED FILE. Do not hand-edit it — edit manifest.json and re-run
this script. Check M-01 in verify_si.py fails if the two have drifted apart.

    python3 gen_sources_md.py            # rewrite sources/SOURCES.md
    python3 gen_sources_md.py --check    # exit 1 if it is out of date
"""
import argparse, json, os, sys

HERE = os.path.dirname(os.path.abspath(__file__))
MANIFEST = os.path.join(HERE, "sources", "manifest.json")
OUTPUT = os.path.join(HERE, "sources", "SOURCES.md")


def esc(s):
    return (s or "").replace("|", "\\|")


def render(man):
    S = man["sources"]
    t1 = [s for s in S if s["tier"] == 1]
    t2 = [s for s in S if s["tier"] == 2]
    t3 = [s for s in S if s["tier"] == 3]
    L = []
    a = L.append

    a("# Sources — full enumeration")
    a("")
    a("<!-- GENERATED FROM manifest.json BY gen_sources_md.py — DO NOT HAND-EDIT. -->")
    a("")
    a("Compiled %s for `%s`." % (man["compiled"], man["document"]))
    a("")
    a("**%d sources: %d bundled and machine-tested, %d consulted but not frozen, "
      "%d declared unobtainable.**" % (len(S), len(t1), len(t2), len(t3)))
    a("")
    a("Every claim in the workbook traces to an entry below. The manifest is the master "
      "copy; this page is a rendering of it.")
    a("")

    a("## Tier 1 — bundled in this repo, tested by the harness")
    a("")
    a("A byte-identical copy is committed here and hashed. Check M-01 re-hashes each one, "
      "so a silently swapped or edited source fails the build.")
    a("")
    a("| ID | Document | File | Completeness | Used by |")
    a("|---|---|---|---|---|")
    for s in t1:
        a("| `%s` | %s | `%s` | %s | %s |" % (
            s["id"], esc(s["title"]), s["file"], s["completeness"],
            ", ".join(s["supports"]) or "—"))
    a("")
    a("**Limits on the partial sources.** A claim may not be stated more strongly than the "
      "source under it. The harness marks any check resting on a partial source with `*`.")
    a("")
    for s in t1:
        if s.get("limitation"):
            a("- `%s` — %s" % (s["id"], esc(s["limitation"])))
    a("")

    a("## Tier 2 — read, but not frozen")
    a("")
    a("Live issuer pages and documents that cannot be redistributed. Each is re-checkable "
      "by hand; none can be frozen into a test. All consulted %s." % man["compiled"])
    a("")
    a("| ID | Source | Issuer token | What it established |")
    a("|---|---|---|---|")
    for s in t2:
        a("| `%s` | %s | %s | %s |" % (
            s["id"], esc(s["title"]),
            ", ".join("`%s`" % i for i in s["issuers"]) or "—",
            esc(s.get("established", ""))))
    a("")
    part2 = [s for s in t2 if s["completeness"] != "full"]
    if part2:
        a("Partial among these:")
        a("")
        for s in part2:
            a("- `%s` — %s" % (s["id"], esc(s["limitation"])))
        a("")

    a("## Tier 3 — not obtained")
    a("")
    a("Declared so the gap is visible rather than implied by silence. Check M-03 fails if "
      "any entry here stops naming what it leaves open.")
    a("")
    a("| ID | Document | Why not | Leaves open |")
    a("|---|---|---|---|")
    for s in t3:
        a("| `%s` | %s | %s | %s |" % (
            s["id"], esc(s["title"]), esc(s["method"]),
            "; ".join(esc(b) for b in s["blocks"]) or "—"))
    a("")

    a("## Issuer coverage")
    a("")
    a("Every issuer named on the workbook's **Credential check** tab must resolve to at "
      "least one source above. Check M-02 enforces this.")
    a("")
    cov = {}
    for s in S:
        for i in s.get("issuers", []):
            cov.setdefault(i, []).append(s["id"])
    for i in sorted(cov):
        a("- **%s** → %s" % (i, ", ".join("`%s`" % x for x in cov[i])))
    a("")

    a("## Refreshing for the next code cycle")
    a("")
    a("Replace the tier 1 files with the new edition's equivalents, update their `sha256` "
      "and `edition` in `manifest.json`, re-run `gen_sources_md.py`, and re-run the harness. "
      "T-01 and T-04 report exactly which citations and titles no longer match. Tier 2 has "
      "to be re-walked by hand — that is the cost of a credential requirement whose issuer "
      "can renumber it without telling anyone.")
    a("")
    return "\n".join(L)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--check", action="store_true")
    args = ap.parse_args()
    man = json.load(open(MANIFEST, encoding="utf-8"))
    out = render(man)
    if args.check:
        cur = open(OUTPUT, encoding="utf-8").read() if os.path.exists(OUTPUT) else ""
        if cur.strip() != out.strip():
            sys.exit("SOURCES.md is out of date — run: python3 gen_sources_md.py")
        print("SOURCES.md is current")
        return
    open(OUTPUT, "w", encoding="utf-8").write(out + "\n")
    print("wrote %s (%d sources)" % (OUTPUT, len(man["sources"])))


if __name__ == "__main__":
    main()
