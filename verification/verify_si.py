#!/usr/bin/env python3
"""
Verification harness for the DPP Special Inspector Qualification workbook.

Deterministic checks only: every assertion compares files, or searches the full text of a
source document held in verification/sources. No language model is involved in any check.

    python3 verify_si.py --new WORKBOOK --orig DPP_ORIGINAL [--excel]

--excel additionally opens the workbook in Microsoft Excel (Windows only) and reads the SI FORM
dropdowns back. It is skipped everywhere else, including on GitHub.

Exit code 0 if every executed check passes, 1 otherwise.

Check families
  I  Integrity      DPP's 2018 data and form are untouched where they must be
  X  Excel format   the file obeys the rules Excel enforces when opening
  W  Who may perform  the 2018 permission notation was translated faithfully; the City welding exception is flagged everywhere it applies
  D  Dropdown       the SI FORM pick-list is complete and covers every matrix row
  L  Legend         every flag value is defined and every flag has a note
  H  Headings       every citation and title matches the full code text; collisions are computed, not asserted
  E  Evidence       every claim has evidence, every quote is in its source, nothing floats
  C  Catalog        every source file matches its fingerprint and is used
  S  Sync           generated tabs and the text dump match the files that own them
  A  Analysis       internal consistency of the matrix
"""
import argparse, json, os, re, subprocess, sys, zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "tools"))
import hashlib
from xlsx import Workbook, read_all, col_number, split_ref
import evidence_lib as EL

RESULTS = []
def check(cid, desc, fn):
    try:
        ok, detail = fn()
    except Exception as e:
        ok, detail = False, "harness error: %r" % (e,)
    RESULTS.append((cid, desc, ok, detail))

def norm(s): return EL.norm(s or "").lower()
def keyrows(cells, col, start=3):
    """{key: row} for a keyed tab."""
    out = {}
    for ref, v in cells.items():
        c, r = split_ref(ref)
        if c == col and r >= start and v: out.setdefault(v, r)
    return out

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--new", required=True); ap.add_argument("--orig", required=True)
    ap.add_argument("--excel", action="store_true")
    a = ap.parse_args()

    NEW, ORIG = read_all(a.new), read_all(a.orig)
    WBN, WBO = Workbook(a.new), Workbook(a.orig)
    M, O = NEW["Matrix 2024"], ORIG["SI Qualification"]
    hdr = {}
    for ref, v in M.items():
        c, r = split_ref(ref)
        if r in (1, 2): hdr[norm(v)] = c
    def col(name): return hdr[norm(name)]
    C = dict(flag=col("Flag"), ref=col("Ref"), sec=col("Section No."), code=col("Code"), title=col("Section title — as printed"),
             dpor=col("DPOR / EOR (PE or AR)"), emp=col("Employee under DPOR supervision"), third=col("DPP-registered 3rd party, or IAS agency"),
             regsi=col("Registered Special Inspector"), icc=col("ICC"), aci=col("ACI"), nicet=col("NICET"), aws=col("AWS"), other=col("Other"),
             basis=col("Credential basis"))
    ROWS = keyrows(M, C["ref"])
    ROWS = {k: r for k, r in ROWS.items() if re.match(r"^(SI|NEW)-\d+$", k)}
    mv = lambda ref, key: M.get("%s%d" % (C[key], ROWS[ref]), "")
    SI = sorted([k for k in ROWS if k.startswith("SI-")], key=lambda k: ROWS[k])
    CERT_COLS = ["icc", "aci", "nicet", "aws", "other"]
    ORIG_CERT = {"icc": "H", "aci": "I", "nicet": "J", "aws": "K", "other": "L"}
    ORIG_PERM = {"dpor": "D", "emp": "E", "third": "F", "regsi": "G"}
    oval = lambda ref, letter: O.get("%s%d" % (letter, int(ref.split("-")[1])), "")
    ev = json.load(open(os.path.join(HERE, "evidence.json"), encoding="utf8"))["evidence"]
    man = EL.manifest(); SRC = {s["id"]: s for s in man["sources"]}

    # ================= I: integrity =================================================================
    check("I-01", "The 2018 tab is cell-for-cell identical to DPP's original", lambda: (
        NEW["SI Qual 2018 (as received)"] == ORIG["SI Qualification"],
        "%d cells compared" % len(ORIG["SI Qualification"])))

    check("I-02", "Drop-down column B rows 2-25 are unchanged", lambda: (
        all(NEW["Drop-down"].get("B%d" % r) == ORIG["Drop-down"].get("B%d" % r) for r in range(2, 26)), "B2:B25"))

    def si_form():
        # Compared by meaning, not bytes, so that re-saving the file in Excel does not trip it.
        o = WBO.xml("SI FORM"); n = WBN.xml("SI FORM"); probs = []
        if NEW["SI FORM"] != ORIG["SI FORM"]: probs.append("cell text differs")
        merges = lambda x: sorted(re.findall(r'<mergeCell ref="([^"]+)"', x))
        if merges(o) != merges(n): probs.append("merged cells differ")
        if set(dv_cells(o)) != set(dv_cells(n)): probs.append("dropdowns apply to different cells")
        legacy = lambda x: sorted(re.findall(r'<dataValidation [^>]*sqref="([^"]+)"', x))
        if legacy(o) != legacy(n): probs.append("other validations differ")
        return (not probs, "same %d cells, merges and validated cells; only the list ranges differ" % len(ORIG["SI FORM"]) if not probs else "; ".join(probs))

    def dv_cells(x):
        """{cell: list formula} for the classification dropdowns. Excel may merge identical rules
        when it saves, so compare the cells covered, not how the rules are grouped."""
        out = {}
        for f, sq in re.findall(r"<xm:f>(.*?)</xm:f>.*?<xm:sqref>(.*?)</xm:sqref>", x, re.S):
            for part in sq.split():
                a_, b_ = (part.split(":") + [part])[:2]
                (c1, r1), (c2, r2) = split_ref(a_), split_ref(b_)
                for r in range(r1, r2 + 1):
                    for c in range(col_number(c1), col_number(c2) + 1):
                        out["%s%d" % (chr(64 + c), r)] = f
        return out
    check("I-03", "The SI FORM tab matches DPP's original (text, merges, validated cells) apart from its list ranges", si_form)

    def dv_rules():
        n, o = dv_cells(WBN.xml("SI FORM")), dv_cells(WBO.xml("SI FORM"))
        targets = set(n.values())
        ok = set(n) == set(o) and len(targets) == 1 and re.fullmatch(r"'Drop-down'!\$B\$1:\$B\$\d+", next(iter(targets), ""))
        return (bool(ok), "all %d dropdown cells (%s) point at %s" % (len(n), ", ".join(sorted(n, key=lambda c: split_ref(c)[1])), ", ".join(targets))
                if ok else "cells %s -> %s" % (sorted(o), n))
    check("I-04", "Every SI FORM dropdown cell survives and all point at one Drop-down range", dv_rules)

    def creds():
        cl = NEW["Change Log"]; register = " ".join(v for k, v in cl.items() if split_ref(k)[0] in ("B", "C", "D"))
        bad = []
        for ref in SI:
            for k in CERT_COLS:
                if norm(mv(ref, k)) != norm(oval(ref, ORIG_CERT[k])):
                    cells = [r for r, v in cl.items() if split_ref(r)[0] == "B" and v == "Credential"]
                    reg = [cl.get("C%d" % split_ref(r)[1], "") for r in cells]
                    if not any(ref in x for x in reg): bad.append("%s %s" % (ref, k.upper()))
        return (not bad, "every certification cell matches 2018 or is registered on the Change Log (Area 'Credential')" if not bad
                else "changed without a Change Log 'Credential' entry: " + ", ".join(bad[:6]))
    check("I-05", "Certification columns match 2018 unless a change is registered with evidence", creds)
    check("I-06", "Workbook archive is not corrupt", lambda: (zipfile.ZipFile(a.new).testzip() is None, "zip integrity"))

    # ================= X: Excel file-format rules ===================================================
    def names():
        bad = [n for n, *_ in WBN.sheets() if len(n) > 31 or re.search(r"[\[\]:*?/\\]", n)]
        ns = [n for n, *_ in WBN.sheets()]; dup = len(ns) != len(set(n.lower() for n in ns))
        return (not bad and not dup, "%d tab names, all <= 31 characters, no forbidden characters" % len(ns) if not bad and not dup
                else "invalid tab names: %s%s" % (bad, " (duplicates)" if dup else ""))
    check("X-01", "Tab names obey Excel's limits (31 characters, no [ ] : * ? / \\)", names)

    ORDER = ["sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetCalcPr", "sheetProtection", "protectedRanges",
             "scenarios", "autoFilter", "sortState", "dataConsolidate", "customSheetViews", "mergeCells", "phoneticPr", "conditionalFormatting",
             "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks",
             "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF", "picture",
             "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"]
    def structure():
        probs = []
        for name, _sid, _rid, part in WBN.sheets():
            x = WBN.parts[part].decode("utf8")
            rows = [int(r) for r in re.findall(r'<row r="(\d+)"', x)]
            if rows != sorted(set(rows)): probs.append("%s: rows out of order" % name)
            for rm in re.finditer(r'<row r="(\d+)"[^>]*?(?:/>|>(.*?)</row>)', x, re.S):
                body = rm.group(2) or ""
                cols = [col_number(c) for c in re.findall(r'<c r="([A-Z]+)\d+"', body)]
                if cols != sorted(set(cols)): probs.append("%s: row %s cells out of order" % (name, rm.group(1)))
                if any(r != rm.group(1) for r in re.findall(r'<c r="[A-Z]+(\d+)"', body)): probs.append("%s: row %s holds another row's cell" % (name, rm.group(1)))
            top = re.sub(r"<sheetData>.*</sheetData>|<sheetData/>", "<sheetData/>", x, flags=re.S)
            top = re.sub(r"<extLst>.*</extLst>", "<extLst/>", top, flags=re.S)
            seen = [t for t in re.findall(r"<(\w+)[ >/]", re.sub(r"^.*?<worksheet[^>]*>", "", top, flags=re.S)) if t in ORDER]
            depth_ok = [ORDER.index(t) for t in seen]
            if depth_ok != sorted(depth_ok): probs.append("%s: element order %s" % (name, seen))
        return (not probs, "%d tabs: rows and cells in order, elements in schema order" % len(WBN.sheets()) if not probs else "; ".join(probs[:4]))
    check("X-02", "Every tab's rows, cells and elements are in the order Excel requires", structure)

    def excel_open():
        if not a.excel: return (None, "run with --excel on a machine with Microsoft Excel")
        ps = ("$x=New-Object -ComObject Excel.Application;$x.DisplayAlerts=$false;"
              "try{$w=$x.Workbooks.Open('%s',0,$true);$v=$w.Worksheets.Item('SI FORM').Range('A9').Validation.Formula1;"
              "'OK|'+$w.Worksheets.Count+'|'+$v;$w.Close($false)}catch{'FAIL|'+$_.Exception.Message}finally{$x.Quit()}") % os.path.abspath(a.new).replace("'", "''")
        out = subprocess.run(["powershell", "-NoProfile", "-Command", ps], capture_output=True, text=True, timeout=180).stdout.strip()
        return (out.startswith("OK|"), "Excel: " + out)
    check("X-03", "Microsoft Excel opens the file and reads the SI FORM dropdown", excel_open)

    # ================= W: who may perform =========================================================
    FN = {"1": "ICC", "2": "ACI", "3": "NICET", "4": "AWS", "5": "Other"}
    PROF = r"(?<![A-Za-z])(SE|AR|CE|ME|FP|IAS)(?![A-Za-z])"
    def translation():
        bad = []; n = 0
        for ref in SI:
            for k, letter in ORIG_PERM.items():
                n += 1; ov = oval(ref, letter).strip(); nv = mv(ref, k).strip()
                probs = []
                if (ov in ("", "NA")) != (nv in ("", "—")): probs.append("permission")
                if sorted({FN[d] for d in re.findall(r"[1-5]", ov)}) != sorted({x for x in FN.values() if x in nv}): probs.append("credentials")
                if sorted(set(re.findall(PROF, ov))) != sorted(set(re.findall(PROF, nv))): probs.append("professions")
                if (" or " in " %s " % ov) != (" or " in " %s " % nv.replace(", or IAS", " IAS")) and not re.search(PROF, ov): probs.append("and/or")
                if probs: bad.append("%s %s (%s)" % (ref, k, "/".join(probs)))
        return (not bad, "%d permission cells match the 2018 notation" % n if not bad else "; ".join(bad[:5]))
    check("W-01", "Who-may-perform cells are a faithful translation of the 2018 notation", translation)

    def welding_exception():
        dec = NEW["Decisions"]; drow = keyrows(dec, "A")
        exc = [r for k, r in drow.items() if "exception" in norm(dec.get("B%d" % r))]
        affected = " ".join(dec.get("E%d" % r, "") for r in exc)
        bad = []
        for ref in SI:
            work = " ".join([mv(ref, "icc"), mv(ref, "aws"), M.get("F%d" % ROWS[ref], ""), M.get("G%d" % ROWS[ref], "")])
            weld = re.search(r"weld|bolt|\(S1\)|\(S2\)|Welding Inspector", work, re.I)
            permitted = mv(ref, "dpor").startswith("✓") or mv(ref, "emp").startswith("✓")
            if weld and permitted and (mv(ref, "flag") != "CONFLICT" or ref not in affected):
                bad.append(ref)
        return (not bad, "every row permitting the DPOR for welding or bolting work is flagged CONFLICT and named on the exception Decision"
                if not bad else "not flagged or not on the Decision: " + ", ".join(bad))
    check("W-02", "The City welding / bolting exception is flagged on every row it touches", welding_exception)

    # ================= D: dropdown ================================================================
    DD = NEW["Drop-down"]
    rng_end = int(re.findall(r"'Drop-down'!\$B\$1:\$B\$(\d+)", WBN.xml("SI FORM"))[0])
    def dd_fill():
        missing = [r for r in range(2, rng_end + 1) if not DD.get("B%d" % r)]
        beyond = [r for r in range(rng_end + 1, rng_end + 15) if DD.get("B%d" % r)]
        return (not missing and not beyond, "B2:B%d filled, nothing beyond the range" % rng_end if not missing and not beyond
                else "empty inside range: %s; stranded beyond it: %s" % (missing, beyond))
    check("D-01", "Dropdown entries fill the SI FORM range exactly", dd_fill)

    def dd_cover():
        listed = set()
        for r in range(2, rng_end + 1): listed |= set(re.findall(r"(?:SI|NEW)-\d+", DD.get("C%d" % r, "")))
        need = {k for k in SI}
        unknown = sorted(listed - set(ROWS)); uncovered = sorted(need - listed, key=lambda k: ROWS[k])
        return (not unknown and not uncovered, "all %d matrix rows reachable from the SI FORM dropdown" % len(need) if not unknown and not uncovered
                else "no dropdown entry: %s; unknown refs in column C: %s" % (uncovered, unknown))
    check("D-02", "Every matrix row is covered by a dropdown entry (Drop-down column C)", dd_cover)

    # ================= L: legend and flags ========================================================
    def legend():
        leg = " ".join(v for k, v in M.items() if split_ref(k)[1] > max(ROWS.values()))
        used = sorted({mv(k, "flag") for k in ROWS if mv(k, "flag")})
        undefined = [u for u in used if not re.search(re.escape(u) + r"\s*=", leg)]
        return (not undefined, "flag values %s all defined in the legend" % used if not undefined else "undefined in legend: %s" % undefined)
    check("L-01", "Every flag value used in the matrix is defined in its legend", legend)

    def flags_tab():
        F = NEW["Flags"]; fr = keyrows(F, "A"); bad = []
        for ref in ROWS:
            fl = mv(ref, "flag")
            if not fl: continue
            if ref not in fr: bad.append("%s has no Flags row" % ref); continue
            if F.get("D%d" % fr[ref]) != fl: bad.append("%s flag differs (%s vs %s)" % (ref, fl, F.get("D%d" % fr[ref])))
            if norm(F.get("F%d" % fr[ref])) != norm(mv(ref, "basis")): bad.append("%s credential basis differs" % ref)
        for ref in fr:
            if ref not in ROWS: bad.append("Flags row %s is not in the matrix" % ref)
        return (not bad, "%d flagged rows, each with a matching Flags note" % sum(1 for k in ROWS if mv(k, "flag")) if not bad else "; ".join(bad[:5]))
    check("L-02", "Every flagged row has a Flags note with the same flag and basis", flags_tab)

    # ================= H: headings and citations ==================================================
    IBC = EL.headings("ibc-2024-ch17")
    def expand(secs):
        out = []; base = None
        for s in [x.strip() for x in secs.split("/")]:
            if s.startswith(".") and base: s = base + s
            if re.match(r"^\d{4}(\.\d+)*$", s): out.append(s); base = s.split(".")[0]
        return out
    def ibc_titles():
        bad = []; n = 0
        for ref in ROWS:
            if mv(ref, "code") != "IBC": continue
            secs = expand(mv(ref, "sec")); titles = [t.strip() for t in re.split(r"\s*\[", mv(ref, "title"))[0].split(" / ")]
            if len(secs) != len(titles): bad.append("%s: %d sections, %d titles" % (ref, len(secs), len(titles))); continue
            for s, t in zip(secs, titles):
                n += 1
                if s not in IBC: bad.append("%s: %s is not a heading in the 2024 IBC" % (ref, s))
                elif norm(IBC[s]) != norm(t): bad.append("%s: %s is '%s', sheet says '%s'" % (ref, s, IBC[s], t))
        return (not bad, "%d IBC citations are headings in the full 2024 Chapter 17, titles exact" % n if not bad else "; ".join(bad[:4]))
    check("H-01", "Every IBC citation is a real 2024 heading and its title is transcribed exactly", ibc_titles)

    def roh_titles():
        bad = []; n = 0
        roh = EL.headings("roh-16-1-1")
        for ref in ROWS:
            code = mv(ref, "code")
            if code == "ROH 16-1.1":
                n += 1; s = mv(ref, "sec")
                if s not in roh: bad.append("%s: %s not found in ROH 16-1.1" % (ref, s))
                elif not norm(roh[s]).startswith(norm(mv(ref, "title")).split(". ")[0]): bad.append("%s: title differs from ROH %s" % (ref, s))
            elif code == "ROH 19-1":
                n += 1; s = mv(ref, "sec"); t = mv(ref, "title")
                head = t.split(".")[0]
                q = re.search(r'"\.{0,3}(.*?)"', t)
                if not EL.contains("roh-19-1", "%s %s." % (s, head)): bad.append("%s: '%s %s.' not in ROH 19-1" % (ref, s, head))
                if q and not EL.contains("roh-19-1", q.group(1)): bad.append("%s: quoted text not in ROH 19-1" % ref)
        return (not bad, "%d City citations and quoted amendment text found in the ordinances" % n if not bad else "; ".join(bad[:4]))
    check("H-02", "Every City citation exists in its ordinance and its title or quote matches", roh_titles)

    def collisions():
        txt = EL.source_text("roh-16-1-1")
        added = set()
        for m in re.finditer(r"Section 1705 is amended by adding Sections? ([\d.]+(?: and [\d.]+)?) to read", txt):
            added |= set(re.findall(r"1705\.\d+", m.group(1)))
        clash = {s for s in added if s in IBC}
        want = {ref for ref in ROWS if not mv(ref, "flag").startswith("GAP") and set(expand(mv(ref, "sec"))) & clash}
        have = {ref for ref in ROWS if mv(ref, "flag") == "COLLISION"}
        return (want == have, "ROH adds %s; %s also exist in the 2024 IBC; COLLISION rows = %s" % (sorted(added), sorted(clash), sorted(have))
                if want == have else "should be COLLISION: %s; flagged: %s" % (sorted(want), sorted(have)))
    check("H-03", "COLLISION flags are exactly the rows whose number both ROH and the 2024 IBC use", collisions)

    # ================= E: evidence ================================================================
    def resolve(key):
        parts = key.split(":")
        tab = parts[0]
        if tab not in NEW: return False
        if tab == "Matrix 2024":
            return len(parts) == 3 and parts[1] in ROWS and parts[2] in ("Flag", "Credential basis", "Section title")
        return parts[1] in keyrows(NEW[tab], "A")
    def quotes():
        bad = []; n = 0
        for e in ev:
            if e["type"] == "supports":
                n += 1; s = SRC.get(e["source"])
                if not s or not s.get("file"): bad.append("%s: source %s not held" % (e["id"], e["source"])); continue
                if e.get("match") == "nospace" and s["completeness"] == "full": bad.append("%s: nospace matching on a complete source" % e["id"]); continue
                if e["part"] == "code" and e["source"].endswith("commentary"): bad.append("%s: code claim cites commentary" % e["id"])
                if not EL.contains(e["source"], e["quote"], match=e.get("match", "exact")): bad.append("%s: quote not found in %s" % (e["id"], e["source"]))
            elif e["type"] == "internal":
                n += 1; tab, ref = e["where"].split("!")
                if norm(ORIG.get(tab, {}).get(ref)) != norm(e["quote"]): bad.append("%s: %s is not '%s'" % (e["id"], e["where"], e["quote"][:30]))
        return (not bad, "%d quotes found verbatim in their sources" % n if not bad else "; ".join(bad[:4]))
    check("E-01", "Every quoted passage is found in its source (code and commentary kept apart)", quotes)

    def absences():
        bad = []; n = 0
        for e in [x for x in ev if x["type"] == "absent"]:
            n += 1; s = SRC.get(e["source"])
            if not s or s["completeness"] != "full": bad.append("%s: absence claimed against an incomplete source" % e["id"]); continue
            text = EL.section_span(e["source"], e["within"]) if e.get("within") else EL.source_text(e["source"])
            if text is None: bad.append("%s: section %s not found" % (e["id"], e["within"])); continue
            hit = [t for t in e["terms"] if t.lower() in text.lower()]
            if hit: bad.append("%s: %s DOES occur in %s%s" % (e["id"], hit, e["source"], (" " + e["within"]) if e.get("within") else ""))
        return (not bad, "%d absence claims hold against complete sources" % n if not bad else "; ".join(bad[:4]))
    check("E-02", "Every 'does not exist' claim is made against a complete source, and holds", absences)

    def opens():
        bad = [e["id"] for e in ev if e["type"] == "open" and not (SRC.get(e["source"]) and
               (SRC[e["source"]]["kind"] == "not-held" or SRC[e["source"]]["completeness"] != "full"))]
        return (not bad, "%d open items each name a source that is not held or incomplete" % sum(1 for e in ev if e["type"] == "open")
                if not bad else "open items naming a complete source: " + ", ".join(bad))
    check("E-03", "Every open item names a source that is genuinely missing or incomplete", opens)

    def keys():
        bad = sorted({u for e in ev for u in e["used_by"] if not resolve(u)})
        orphans = [e["id"] for e in ev if not e["used_by"]]
        return (not bad and not orphans, "%d references from evidence to workbook cells all resolve; no orphan evidence" % sum(len(e["used_by"]) for e in ev)
                if not bad and not orphans else "unresolved: %s; orphans: %s" % (bad[:5], orphans[:5]))
    check("E-04", "Every evidence reference points at a real row, and every entry is used", keys)

    used = {u for e in ev for u in e["used_by"]}
    def coverage():
        need = []
        for ref in ROWS:
            fl = mv(ref, "flag")
            if fl and not fl.startswith("GAP"): need.append("Matrix 2024:%s:Flag" % ref)
            if ref.startswith("SI-"): need.append("Matrix 2024:%s:Credential basis" % ref)
        need += ["Flags:" + k for k in keyrows(NEW["Flags"], "A")]
        need += ["Decisions:" + k for k in keyrows(NEW["Decisions"], "A")]
        need += ["Credential check:" + k for k in keyrows(NEW["Credential check"], "A")]
        cl = NEW["Change Log"]
        need += ["Change Log:" + k for k, r in keyrows(cl, "A").items() if cl.get("B%d" % r) in ("Citation", "Title", "Credential")]
        floating = [k for k in need if k not in used]
        return (not floating, "%d claims, every one backed by evidence" % len(need) if not floating else "%d floating, e.g. %s" % (len(floating), floating[:6]))
    check("E-05", "Nothing floats: every flag, basis, note, decision, credential finding and citation change has evidence", coverage)

    def basis_rules():
        bad = []
        by = {}
        for e in ev:
            for u in e["used_by"]: by.setdefault(u, []).append(e)
        for ref in SI:
            b = mv(ref, "basis"); es = by.get("Matrix 2024:%s:Credential basis" % ref, [])
            if re.match(r"^(code|standard|ordinance)", b) and not any(e["type"] == "supports" and e["part"] != "commentary" for e in es):
                bad.append("%s: '%s' but no supporting (non-commentary) evidence" % (ref, b.split(":")[0]))
            if "unverified" in b and not any(e["type"] == "open" for e in es): bad.append("%s: unverified but no open item" % ref)
            if b.startswith("DPP policy") and not any(e["type"] in ("absent", "internal", "supports") for e in es): bad.append("%s: DPP policy with no evidence" % ref)
            if b == "no credential listed" and any(mv(ref, k) for k in CERT_COLS): bad.append("%s: says no credential but lists one" % ref)
        return (not bad, "credential basis for %d rows is consistent with its evidence type" % len(SI) if not bad else "; ".join(bad[:4]))
    check("E-06", "Credential basis wording matches its evidence (a requirement never rests on commentary)", basis_rules)

    # ================= C: catalog =================================================================
    def catalog_files():
        probs = []; listed = set()
        for s in man["sources"]:
            if not s.get("file"): continue
            listed.add(s["file"]); p = os.path.join(EL.SOURCES, s["file"])
            if not os.path.exists(p): probs.append("%s missing" % s["file"]); continue
            raw = open(p, "rb").read()
            if hashlib.sha256(raw).hexdigest() != s["sha256"]: probs.append("%s fingerprint mismatch" % s["id"])
        disk = {sub + "/" + f for sub in ("licensed", "public", "snapshots") if os.path.isdir(os.path.join(EL.SOURCES, sub))
                for f in os.listdir(os.path.join(EL.SOURCES, sub)) if not f.startswith(".")}
        probs += ["%s on disk but not in the catalog" % f for f in sorted(disk - listed)]
        from gen_sources_md import render
        md = os.path.join(EL.SOURCES, "SOURCES.md")
        if not os.path.exists(md) or open(md, encoding="utf8").read().strip() != render(man).strip(): probs.append("SOURCES.md out of date")
        return (not probs, "%d source files match their fingerprints; nothing unlisted; SOURCES.md current" % len(listed) if not probs else "; ".join(probs[:4]))
    check("C-01", "Every source file matches its catalog fingerprint, and nothing is unlisted", catalog_files)

    def catalog_use():
        cited = {e["source"] for e in ev}; probs = []
        for s in man["sources"]:
            if s["use"] == "evidence" and s["id"] not in cited: probs.append("%s is marked evidence but nothing cites it" % s["id"])
            if s["use"] == "context" and not s.get("context_reason"): probs.append("%s is context with no reason" % s["id"])
            if s["kind"] == "not-held" and not s.get("blocks"): probs.append("%s not held but names nothing it blocks" % s["id"])
            if s.get("file") and s["completeness"] != "full" and not s.get("limitation"): probs.append("%s partial with no limitation" % s["id"])
        return (not probs, "every held source is cited or has a stated reason; every gap says what it leaves open" if not probs else "; ".join(probs[:4]))
    check("C-02", "Every source is used, and every partial or missing source states its limit", catalog_use)

    def issuers():
        known = {s.get("issuer") for s in man["sources"]}
        cc = NEW["Credential check"]; miss = sorted({cc.get("B%d" % r) for r in keyrows(cc, "A").values() if cc.get("B%d" % r) not in known})
        return (not miss, "every issuer on Credential check has a catalog entry" if not miss else "issuers with no source: %s" % miss)
    check("C-03", "Every issuer on Credential check resolves to a catalog source", issuers)

    # ================= S: sync ====================================================================
    def sync():
        import sync_workbook
        d = sync_workbook.check(a.new)
        return (not d, "README, Evidence and Sources tabs, Evidence columns and workbook.txt all current" if not d
                else "out of date (run tools/sync_workbook.py): %d cells, e.g. %s" % (len(d), d[:5]))
    check("S-01", "Generated tabs, Evidence columns and the text dump match the files that own them", sync)

    # ================= A: analysis ================================================================
    check("A-01", "Every matrix row carries a Ref, a Code and a Section", lambda: (
        all(mv(k, "code") and mv(k, "sec") for k in ROWS), "%d rows" % len(ROWS)))
    check("A-02", "Every Ref is unique", lambda: (
        len([v for k, v in M.items() if split_ref(k)[0] == C["ref"] and re.match(r"^(SI|NEW)-", v)]) == len(ROWS), "%d refs" % len(ROWS)))
    def named(word):
        def fn():
            dec = NEW["Decisions"]; txt = " ".join(dec.get("E%d" % r, "") for r in keyrows(dec, "A").values())
            rows = [k for k in ROWS if mv(k, "flag") == word]; miss = [k for k in rows if k not in txt]
            return (not miss, "%d %s rows, each named on Decisions" % (len(rows), word) if not miss else "not on any Decision: " + ", ".join(miss))
        return fn
    check("A-03", "Every COLLISION row is named in a Decision's Rows affected", named("COLLISION"))
    check("A-04", "Every CONFLICT row is named in a Decision's Rows affected", named("CONFLICT"))
    check("A-05", "GAP rows carry no invented certifications", lambda: (
        not [k for k in ROWS if mv(k, "flag").startswith("GAP") and any(mv(k, c) for c in CERT_COLS)], "gap rows have empty certification cells"))

    # ================= report =====================================================================
    w = max(len(d) for _, d, _, _ in RESULTS)
    run = [x for x in RESULTS if x[2] is not None]; passed = [x for x in run if x[2]]
    print("=" * (w + 30)); print("VERIFICATION REPORT — deterministic checks only"); print("=" * (w + 30))
    for cid, desc, ok, detail in RESULTS:
        tag = "SKIP" if ok is None else "PASS" if ok else "FAIL"
        print("%-5s %s  %-*s  %s" % (cid, tag, w, desc, detail))
    print("-" * (w + 30))
    print("%d/%d executed checks passed, %d skipped" % (len(passed), len(run), len(RESULTS) - len(run)))

    print("\nOPEN — things no check can settle (from evidence.json and the workbook):")
    for e in [x for x in ev if x["type"] == "open"]:
        s = SRC.get(e["source"], {})
        print("  %s  %s  [%s: %s]" % (e["id"], e["claim"], e["source"], s.get("blocks") or s.get("limitation") or ""))
    cc = NEW["Credential check"]
    for k, r in keyrows(cc, "A").items():
        if cc.get("F%d" % r) in ("UNCONFIRMED", "not checked"): print("  %s  %s: %s" % (k, cc.get("B%d" % r), cc.get("G%d" % r, "")))
    dec = NEW["Decisions"]; waiting = [k for k, r in keyrows(dec, "A").items() if not dec.get("F%d" % r)]
    print("  Decisions awaiting a DPP answer: %d of %d" % (len(waiting), len(keyrows(dec, "A"))))
    sys.exit(0 if len(passed) == len(run) else 1)

if __name__ == "__main__":
    main()
