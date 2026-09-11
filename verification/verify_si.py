#!/usr/bin/env python3
"""
Verification harness for the DPP Special Inspector Qualification workbook.

Runs deterministic checks only. No language model is involved in any assertion
below: every check either compares files, or greps a source document that is
passed in on disk. Checks that CANNOT be decided by a program are listed at the
end as NOT MACHINE-VERIFIABLE, with the reason.

Usage:
    python3 verify_si.py --new <rev.xlsx> --orig <2018 original.xlsx> [--sources DIR]

Exit code 0 if every executed check passes, 1 otherwise.
"""
import argparse, re, sys, zipfile, os, json, hashlib
try:
    import openpyxl
except ImportError:
    sys.exit("openpyxl required:  pip install openpyxl")
import warnings; warnings.filterwarnings("ignore")

UNVERIFIABLE="""NOT MACHINE-VERIFIABLE — these require a document I cannot obtain, or a judgement:
  U-01  Whether NICET still offers "Geotechnical Engineering Technology".
        Absent from their current programs page; a 1994 manual is still hosted with
        no withdrawal notice. Needs the issuer.
  U-02  Whether DPP intends the 2021 UPC reference despite ROH 19-1 adopting the
        2018 edition. Intent, not fact.
  U-03  Whether CAWI alongside CWI is acceptable to DPP. AWS D1.1 Sec. 8.1.4.5 makes
        CAWI an assistant qualification; whether that bars the listed use is a
        determination for the building official.
  U-04  HDOA course numbers 4322 / 4415 / 4422 / 4484. Portal returns Forbidden.
  U-05  Whether the certification set for any row is the CORRECT one. The code
        delegates this to the building official, so there is no external referent.
"""

RESULTS=[]
def check(cid, method, desc, fn, sources=()):
    """method: I=Inspection (file compare)  A=Analysis (derived)  T=Test (against source doc)
                M=Manifest (the source corpus checking itself)
    sources: manifest IDs this check reads. Declared, not inferred — M-02 checks both
    directions, so a check that quietly reads an undeclared source fails the build."""
    try:
        ok, detail = fn()
    except Exception as e:
        ok, detail = False, "harness error: %r" % (e,)
    RESULTS.append((cid, method, desc, ok, detail, list(sources)))

def norm(s):
    return re.sub(r"\s+", " ", (s or "")).strip().lower()

def main():
    ap=argparse.ArgumentParser()
    ap.add_argument("--new", required=True); ap.add_argument("--orig", required=True)
    ap.add_argument("--sources", default=os.path.join(os.path.dirname(os.path.abspath(__file__)),"sources"))
    a=ap.parse_args()

    NEW=openpyxl.load_workbook(a.new); ORIG=openpyxl.load_workbook(a.orig)
    zn=zipfile.ZipFile(a.new); zo=zipfile.ZipFile(a.orig)
    M=NEW["Matrix 2024"]
    def col(name):
        for c in range(1, M.max_column+1):
            if norm(M.cell(2,c).value)==norm(name) or norm(M.cell(1,c).value)==norm(name): return c
        raise KeyError(name)
    C_FLAG, C_REF, C_SEC, C_CODE = col("Flag"), col("Ref"), col("Section No."), col("Code")
    C_TITLE, C_CLS = col("Section title — as printed"), col("Classification")
    DATA=[r for r in range(3, M.max_row+1) if M.cell(r,C_REF).value]

    # ---------- I: the file has not been tampered with where it claims not to be ----------
    check("I-01","I","The 2018 original tab is cell-for-cell identical to the source workbook",
      lambda: (all(ORIG["SI Qualification"].cell(r,c).value ==
                   NEW["SI Qualification (2018 as received)"].cell(r,c).value
                   for r in range(1,57) for c in range(1,14)), "rows 1-56, cols A-M"))

    check("I-02","I","Drop-down column B rows 2-25 are unchanged (they feed the SI FORM validation)",
      lambda: (all(ORIG["Drop-down"].cell(r,2).value==NEW["Drop-down"].cell(r,2).value
                   for r in range(2,26)), "B2:B25"))

    def si_form_diff():
        o=zo.read("xl/worksheets/sheet4.xml").decode()
        n=zn.read("xl/worksheets/sheet4.xml").decode()
        o2=re.sub(r"\$B\$1:\$B\$\d+","RANGE",o).replace("\r","").strip()
        n2=re.sub(r"\$B\$1:\$B\$\d+","RANGE",n).replace("\r","").strip()
        return (o2==n2, "identical once the validation ranges are normalised" if o2==n2
                else "SI FORM differs beyond the validation ranges")
    check("I-03","I","The SI FORM tab differs from the original ONLY in its validation ranges", si_form_diff)

    def dv_rules():
        n=zn.read("xl/worksheets/sheet4.xml").decode()
        rules=re.findall(r"<xm:f>(.*?)</xm:f>", n)
        cnt=n.count("x14:dataValidation")
        ok = cnt==zo.read("xl/worksheets/sheet4.xml").decode().count("x14:dataValidation") \
             and len(rules)==3 and all(x=="'Drop-down'!$B$1:$B$29" for x in rules)
        return (ok, "%d validation nodes, list ranges=%s" % (cnt, set(rules)))
    check("I-04","I","All SI FORM dropdown rules survive and point at the extended range", dv_rules)

    def creds_untouched():
        bad=[]
        for r in DATA:
            ref=M.cell(r,C_REF).value
            if not str(ref).startswith("SI-"): continue
            orow=int(str(ref).split("-")[1])
            for off,ocol in enumerate([8,9,10,11,12]):
                mv=M.cell(r, col("ICC")+off).value
                ov=ORIG["SI Qualification"].cell(orow, ocol).value
                if norm(mv)!=norm(ov): bad.append("%s col%+d"%(ref,off))
        return (not bad, "all certification cells reproduced verbatim" if not bad else "differs: "+", ".join(bad[:6]))
    check("I-05","I","Certification columns are reproduced from the 2018 sheet without edits", creds_untouched)

    check("I-06","I","Workbook archive is not corrupt", lambda: (zn.testzip() is None, "zip integrity"))

    # ---------- A: internal consistency of the new work ----------
    check("A-01","A","Every matrix row carries a Ref, a Code and a Section",
      lambda: (all(M.cell(r,C_REF).value and M.cell(r,C_CODE).value and M.cell(r,C_SEC).value
                   for r in DATA), "%d rows"%len(DATA)))

    check("A-02","A","Every Ref is unique",
      lambda: (len({M.cell(r,C_REF).value for r in DATA})==len(DATA), "%d refs"%len(DATA)))

    def basis_rule():
        """Recompute Credential basis independently from the ORIGINAL sheet and compare."""
        try: C_B=col("Credential basis")
        except KeyError: return (False,"Credential basis column missing")
        bad=[]
        for r in DATA:
            ref=str(M.cell(r,C_REF).value)
            if not ref.startswith("SI-"):
                continue
            o=int(ref.split("-")[1]); s=ORIG["SI Qualification"]
            aci,aws=s.cell(o,9).value, s.cell(o,11).value
            oth=str(s.cell(o,12).value or ""); icc,nic=s.cell(o,8).value, s.cell(o,10).value
            tags=[]
            if aws or aci: tags.append("standard")
            if ("AABC" in oth or "Environmental Balancing" in oth) and "standard" not in tags: tags.append("standard")
            if "HDOA" in oth: tags.append("statute")
            if icc or nic or (oth and not ("HDOA" in oth and len(oth)<130)): tags.append("policy")
            want=" + ".join(dict.fromkeys(tags)) or "policy"
            if norm(M.cell(r,C_B).value)!=norm(want): bad.append("%s: is %r want %r"%(ref,M.cell(r,C_B).value,want))
        return (not bad, "rule reproduced on every row" if not bad else "; ".join(bad[:4]))
    check("A-03","A","Credential basis matches the rule stated on the sheet, recomputed from source", basis_rule)

    def flagged_have_decisions(word, cid, label):
        def fn():
            dec=NEW["Decisions"]; txt=norm(" ".join(str(dec.cell(r,c).value or "")
                  for r in range(1,dec.max_row+1) for c in range(1,dec.max_column+1)))
            rowsf=[M.cell(r,C_REF).value for r in DATA if norm(M.cell(r,C_FLAG).value)==word]
            missing=[ref for ref in rowsf if norm(str(ref)) not in txt]
            return (not missing, "%d %s rows, each named by Ref on the Decisions tab"%(len(rowsf),label) if not missing
                    else "Ref not named in any Decision: "+", ".join(sorted(set(map(str,missing))))[:90])
        return fn
    check("A-04","A","Every COLLISION row is traceable to a Decision entry", flagged_have_decisions("collision","A-04","collision"))
    check("A-07","A","Every CONFLICT row is traceable to a Decision entry", flagged_have_decisions("conflict","A-07","conflict"))

    def gaps_have_no_creds():
        try: C_ICC=col("ICC")
        except KeyError: return (False,"ICC column missing")
        bad=[M.cell(r,C_REF).value for r in DATA
             if "gap" in norm(M.cell(r,C_FLAG).value)
             and any(M.cell(r,C_ICC+o).value for o in range(5))]
        return (not bad, "gap rows carry no invented certifications" if not bad else "certs present on "+", ".join(map(str,bad)))
    check("A-05","A","Rows flagged as gaps have no certifications assigned to them", gaps_have_no_creds)

    def dropdown_in_range():
        dd=NEW["Drop-down"]
        entries=[dd.cell(r,2).value for r in range(2,30)]
        filled=[e for e in entries if e]
        beyond=[dd.cell(r,2).value for r in range(30,40) if dd.cell(r,2).value]
        return (len(filled)==28 and not beyond,
                "%d entries inside B1:B29, %d stranded beyond it"%(len(filled),len(beyond)))
    check("A-06","A","Every dropdown entry sits inside the validation range B1:B29", dropdown_in_range)

    # ---------- T: claims tested against the source documents on disk ----------
    def load(fn):
        for sub in ("", "public", "licensed"):
            p=os.path.join(a.sources, sub, fn)
            if os.path.exists(p):
                return open(p, encoding="utf-8", errors="ignore").read()
        return None

    IBC=load("ibc2024_ch17_index.txt") or load("ibc2024_ch17.txt"); ROH=load("roh_16_1_1.txt"); ROH19=load("roh_19_1.txt")

    def cite_in_source(codeval, src, label):
        if src is None: return (None, "source not supplied — skipped")
        miss=[]
        for r in DATA:
            if norm(M.cell(r,C_CODE).value)!=norm(codeval): continue
            for sec in re.split(r"[;/]", str(M.cell(r,C_SEC).value)):
                sec=sec.strip()
                if not re.match(r"^\d{3,4}(\.\d+)*$", sec): continue
                if not re.search(r"(?<![\d.])"+re.escape(sec)+r"(?![\d])", src):
                    miss.append("%s (%s)"%(sec, M.cell(r,C_REF).value))
        return (not miss, "every %s citation found in %s"%(codeval,label) if not miss
                else "not found: "+", ".join(sorted(set(miss))[:6]))

    check("T-01","T","Every citation marked IBC appears in the 2024 IBC Chapter 17 text",
      lambda: cite_in_source("IBC", IBC, "the 2024 IBC text"), sources=['ibc-2024-ch17-index'])
    check("T-02","T","Every citation marked ROH 16-1.1 appears in that ordinance text",
      lambda: cite_in_source("ROH 16-1.1", ROH, "ROH 16-1.1"), sources=['roh-16-1-1'])
    check("T-03","T","Every citation marked ROH 19-1 appears in that ordinance text",
      lambda: cite_in_source("ROH 19-1", ROH19, "ROH 19-1"), sources=['roh-19-1'])

    def titles_match():
        if IBC is None: return (None,"source not supplied — skipped")
        bad=[]
        for r in DATA:
            if norm(M.cell(r,C_CODE).value)!="ibc": continue
            sec=str(M.cell(r,C_SEC).value).strip()
            if not re.match(r"^\d{3,4}(\.\d+)*$", sec): continue
            title=str(M.cell(r,C_TITLE).value or "")
            title=re.split(r"\s*\[", title)[0].strip()          # drop bracketed commentary
            if not title: continue
            pat=re.compile(re.escape(sec)+r"\s+"+re.escape(title[:24]), re.I)
            if not pat.search(IBC): bad.append("%s %s"%(sec,title[:28]))
        return (not bad, "section titles match the code text" if not bad else "mismatch: "+"; ".join(bad[:4]))
    check("T-04","T","Each IBC section title is transcribed as printed in the code text", titles_match, sources=['ibc-2024-ch17-index'])

    def collision_is_real():
        if IBC is None or ROH is None: return (None,"sources not supplied — skipped")
        ibc_secs={str(M.cell(r,C_SEC).value).strip() for r in DATA if norm(M.cell(r,C_CODE).value)=="ibc"}
        roh_secs={str(M.cell(r,C_SEC).value).strip() for r in DATA if norm(M.cell(r,C_CODE).value)=="roh 16-1.1"}
        overlap=ibc_secs & roh_secs
        flagged={str(M.cell(r,C_SEC).value).strip() for r in DATA if norm(M.cell(r,C_FLAG).value)=="collision"}
        return (overlap.issubset(flagged) and overlap=={"1705.19","1705.20"},
                "shared numbers: %s ; flagged: %s"%(sorted(overlap), sorted(flagged & overlap)))
    check("T-05","T","The claimed collisions are exactly the section numbers used by both IBC and ROH", collision_is_real, sources=['ibc-2024-ch17-index', 'roh-16-1-1'])


    ACI=load("aci318_19_sec26_13.txt"); AWSD=load("aws_d1_1_sec8_1_4.txt")

    def aci_rebar_d14():
        if ACI is None: return (None,"source not supplied — skipped")
        ok = re.search(r"26\.13\.1\.4.{0,400}?D1\.4", ACI, re.S) is not None
        return (ok, "ACI 318-19 Sec. 26.13.1.4 names AWS D1.4 for reinforcement welding"
                    if ok else "could not locate the D1.4 requirement in Sec. 26.13.1.4")
    check("T-06","T","The claim that rebar welding inspection follows AWS D1.4 is in ACI 318-19", aci_rebar_d14, sources=['aci-318-19-sec26-13'])

    def aci_anchor_programs():
        if ACI is None: return (None,"source not supplied — skipped")
        a = "681.2" in ACI and "681.1" in ACI
        b = re.search(r"26\.13\.1\.[56]", ACI) is not None
        return (a and b, "ACI 318-19 Sec. 26.13.1.5/.1.6 name CPP 681.2 and CPP 681.1"
                         if (a and b) else "anchor inspector programs not found where claimed")
    check("T-07","T","The claim that post-installed anchors need the ACI anchor inspector programs is in ACI 318-19", aci_anchor_programs, sources=['aci-318-19-sec26-13'])

    def aws_tiers():
        if AWSD is None: return (None,"source not supplied — skipped")
        insp = re.search(r"8\.1\.4\.2.{0,600}?Certified Welding Inspector \(CWI\)", AWSD, re.S) is not None
        asst = re.search(r"8\.1\.4\.5.{0,600}?Certified Associate Welding Inspector \(CAWI\)", AWSD, re.S) is not None
        sup  = "under the supervision of the Inspector" in AWSD
        return (insp and asst and sup,
                "8.1.4.2 -> CWI/SCWI for the Inspector; 8.1.4.5 -> CAWI for the Assistant, under supervision"
                if (insp and asst and sup) else "tier structure not confirmed in the extract")
    check("T-08","T","The claim that CAWI is an ASSISTANT qualification, not an alternative to CWI, is in AWS D1.1", aws_tiers, sources=['aws-d1-1-2020-sec8-1-4'])

    def credcheck_traceable():
        try: CC=NEW["Credential check"]
        except KeyError: return (False,"Credential check tab missing")
        SRC_=NEW["Sources"]
        srctxt=norm(" ".join(str(SRC_.cell(r,c).value or "") for r in range(1,SRC_.max_row+1) for c in range(1,SRC_.max_column+1)))
        hard=[]
        for r in range(3, CC.max_row+1):
            status=norm(CC.cell(r,5).value); issuer=norm(CC.cell(r,1).value)
            if status in ("error","retired","unconfirmed","scope — check","naming","current","restructured"):
                if issuer and issuer.split()[0] not in srctxt: hard.append(CC.cell(r,1).value)
        miss=sorted(set(x for x in hard if x))
        return (not miss, "every issuer appearing in Credential check is accounted for on Sources"
                if not miss else "issuer not listed on Sources: "+", ".join(map(str,miss))[:80])
    check("A-08","A","Every issuer named in Credential check is accounted for on the Sources tab", credcheck_traceable)

    def unverifiable_declared():
        SRC_=NEW["Sources"]
        txt=norm(" ".join(str(SRC_.cell(r,c).value or "") for r in range(1,SRC_.max_row+1) for c in range(1,SRC_.max_column+1)))
        need=["nicet","hdoa","aisc","tms","awc"]
        miss=[n for n in need if n not in txt]
        return (not miss, "every un-consulted source is declared on the Sources tab"
                if not miss else "not declared: "+", ".join(miss))
    check("A-09","A","Sources declares what was NOT consulted, not only what was", unverifiable_declared)

    # ---------- M: the source corpus checked against its own manifest ----------
    # These answer "were all the sources used, and is each claim as strong as its source?"
    # in a form that can fail. Before the manifest existed, that question had no falsifiable
    # shape and was answered by assertion.
    MANP = os.path.join(a.sources, "manifest.json")
    MAN  = json.load(open(MANP, encoding="utf-8")) if os.path.exists(MANP) else None
    SRCID = {x["id"]: x for x in MAN["sources"]} if MAN else {}

    def manifest_integrity():
        """Bundled sources are byte-identical to what the manifest says, in both directions,
        and SOURCES.md has not drifted from the manifest."""
        if MAN is None: return (False, "manifest.json not found at "+MANP)
        problems=[]; listed=set()
        for src in MAN["sources"]:
            if not src.get("file"): continue
            rel=src["file"]; listed.add(rel)
            path=os.path.join(a.sources, rel)
            if not os.path.exists(path):
                problems.append("%s: file missing"%src["id"]); continue
            raw=open(path,"rb").read()
            if hashlib.sha256(raw).hexdigest()!=src["sha256"]:
                problems.append("%s: sha256 mismatch"%src["id"])
            elif src.get("bytes") and len(raw)!=src["bytes"]:
                problems.append("%s: size mismatch"%src["id"])
        on_disk=set()
        for sub in ("public","licensed"):
            d=os.path.join(a.sources, sub)
            if os.path.isdir(d):
                on_disk |= {sub+"/"+f for f in os.listdir(d) if not f.startswith(".")}
        for extra in sorted(on_disk - listed):
            problems.append("%s: on disk but not in the manifest"%extra)
        # SOURCES.md must be a faithful rendering of the manifest
        try:
            sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
            from gen_sources_md import render
            want=render(MAN).strip()
            md=os.path.join(a.sources,"SOURCES.md")
            have=open(md,encoding="utf-8").read().strip() if os.path.exists(md) else ""
            if want!=have:
                problems.append("SOURCES.md is out of date — run gen_sources_md.py")
        except Exception as e:
            problems.append("could not verify SOURCES.md: %r"%(e,))
        return (not problems,
                "%d bundled sources hashed and matched; SOURCES.md in sync"%len(listed)
                if not problems else "; ".join(problems[:4]))
    check("M-01","M","Bundled sources match the manifest by hash, nothing is unlisted, SOURCES.md is in sync",
          manifest_integrity)

    def manifest_coverage():
        """Both directions. Forward: every source a check declares exists in the manifest.
        Reverse: every bundled source is actually consumed by a check — an unused bundled
        source means either dead weight or a check somebody meant to write and didn't.
        Third: every issuer on the Credential check tab resolves to a manifest entry."""
        if MAN is None: return (False, "manifest.json not found")
        problems=[]
        declared=set()
        for cid,_m,_d,_ok,_det,srcs in RESULTS:
            for sid in srcs:
                declared.add(sid)
                if sid not in SRCID:
                    problems.append("%s declares unknown source %r"%(cid,sid))
        for src in MAN["sources"]:
            if src.get("frozen") and src["id"] not in declared:
                problems.append("%s is bundled but no check reads it"%src["id"])
            if src.get("frozen"):
                claimed=set(src.get("supports") or [])
                actual={cid for cid,_m,_d,_ok,_det,srcs in RESULTS if src["id"] in srcs}
                if claimed!=actual:
                    problems.append("%s: manifest says %s, harness uses %s"%(
                        src["id"], sorted(claimed) or "none", sorted(actual) or "none"))
        try:
            CC=NEW["Credential check"]
            covered={i for src in MAN["sources"] for i in src.get("issuers") or []}
            seen=[]
            for r in range(3, CC.max_row+1):
                v=CC.cell(r,1).value
                if v and v not in seen: seen.append(v)
            for iss in seen:
                if iss not in covered:
                    problems.append("issuer %r on Credential check has no source in the manifest"%iss)
        except KeyError:
            problems.append("Credential check tab missing")
        return (not problems,
                "%d sources; every bundled one consumed; %d issuers all resolve"%(
                    len(MAN["sources"]), len({i for s_ in MAN["sources"] for i in s_.get("issuers") or []}))
                if not problems else "; ".join(problems[:4]))
    check("M-02","M","Source coverage closes in both directions, and every issuer resolves to a source",
          manifest_coverage)

    def confidence_discipline():
        """No claim may be stated more strongly than the source under it. Mechanically:
        a source that was not read in full must say what it cannot establish; a source that
        could not be obtained must name what it leaves open; and every open item it names
        must actually appear in the NOT MACHINE-VERIFIABLE list."""
        if MAN is None: return (False, "manifest.json not found")
        problems=[]; declared_U=set()
        for src in MAN["sources"]:
            if src["completeness"]!="full" and not (src.get("limitation") or "").strip():
                problems.append("%s is %s but states no limitation"%(src["id"],src["completeness"]))
            if src["completeness"]=="unavailable" and not src.get("blocks"):
                problems.append("%s is unavailable but names nothing it blocks"%src["id"])
            for b in src.get("blocks") or []:
                m_=re.match(r"^(U-\d+)$", str(b).strip())
                if m_: declared_U.add(m_.group(1))
        for u in sorted(declared_U):
            if u not in UNVERIFIABLE:
                problems.append("%s is named in the manifest but not listed as unverifiable"%u)
        qualified=[cid for cid,_m,_d,ok,_det,srcs in RESULTS
                   if ok and any(SRCID.get(x,{}).get("completeness")!="full" for x in srcs)]
        return (not problems,
                "every partial source states its limit; %s rest on partial sources and are "
                "reported qualified"%(", ".join(qualified) or "no checks")
                if not problems else "; ".join(problems[:4]))
    check("M-03","M","Partial and unobtainable sources state their limits, and every open item they name is declared",
          confidence_discipline)

    # ---------- report ----------
    w=max(len(d) for _,_,d,_,_,_ in RESULTS)
    run=[x for x in RESULTS if x[3] is not None]
    passed=[x for x in run if x[3]]

    def partial_of(srcs):
        return [x for x in srcs if SRCID.get(x,{}).get("completeness") not in (None,"full")]

    print("="*(w+28))
    print("VERIFICATION REPORT — deterministic checks only, no model judgement")
    print("="*(w+28))
    qualified=[]
    for cid,meth,desc,ok,detail,srcs in RESULTS:
        part=partial_of(srcs)
        if ok is None:   tag="SKIP "
        elif not ok:     tag="FAIL "
        elif part:       tag="PASS*"; qualified.append((cid,part))
        else:            tag="PASS "
        print("%-5s %s  %-*s  %s" % (cid, tag, w, desc, detail))
    print("-"*(w+28))
    print("%d/%d executed checks passed, %d skipped"%(len(passed),len(run),len(RESULTS)-len(run)))

    if qualified:
        print("""
PASS* — PASSED AGAINST A SOURCE THAT WAS NOT READ IN FULL.
The check is sound; the source under it is an extract. Read each as "consistent with the
part of the document held here", not "confirmed against the document".""")
        for cid,part in qualified:
            for sid in part:
                print("  %-5s %s\n        %s" % (cid, sid, SRCID[sid]["limitation"]))

    if MAN:
        t={1:0,2:0,3:0}
        for x in MAN["sources"]: t[x["tier"]]=t.get(x["tier"],0)+1
        print("""
SOURCE CORPUS — %d sources enumerated in verification/sources/manifest.json
  tier 1  %2d  bundled here, hashed, and read by the checks above
  tier 2  %2d  read during the work but not frozen: live issuer pages, non-redistributable
  tier 3  %2d  could not be obtained — each one names what it leaves open"""
              % (len(MAN["sources"]), t.get(1,0), t.get(2,0), t.get(3,0)))

    print("\n"+UNVERIFIABLE)
    sys.exit(0 if len(passed)==len(run) else 1)

if __name__=="__main__":
    main()
