#!/usr/bin/env python3
"""
Turns the original source documents into the text files the harness reads, and writes
the catalog (sources/manifest.json) that records where each one came from.

Runs LOCALLY, by whoever holds the originals. The originals (PDFs, docx, saved web pages)
are NOT committed - they are large and most are copyrighted. What is committed is the
extracted text plus the sha256 of the original it came from, so anyone holding the same
original can re-run this and get a byte-identical text file.

    python3 tools/extract_sources.py --originals DIR [--snapshots DIR]

Every entry in CATALOG below is the only place a source is declared. To add a source:
add an entry, re-run, re-run gen_sources_md.py, commit.
"""
import argparse, hashlib, html, json, os, re, subprocess, sys, zipfile

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))   # verification/
SRC = os.path.join(HERE, "sources")
RETRIEVED = "2026-09-11"

# kind: code | standard | ordinance | state-rule | issuer-page | accreditation
# use:  evidence = must be cited by at least one evidence entry (check C-02)
#       context  = held for reference only; 'context_reason' is required
CATALOG = [
  # ---- base codes (ICC, copyrighted) -------------------------------------------------
  dict(id="ibc-2024-ch17", title="International Building Code 2024, Chapter 17 Special Inspections and Tests (full chapter, 1701-1709 incl. tables)",
       issuer="ICC", edition="2024", kind="code", folder="licensed", out="ibc2024_ch17.txt",
       original="2024 IBC Chapter 17_1.docx", extractor="docx", completeness="full", use="evidence"),
  dict(id="ibc-2024-ch35", title="International Building Code 2024, Chapter 35 Referenced Standards (UpCodes, Austin viewer)",
       issuer="ICC", edition="2024", kind="code", folder="licensed", out="ibc2024_ch35.txt",
       original="ibc2024_ch35_austin.html", extractor="html", url="https://up.codes/viewer/austin/ibc-2024/chapter/35/referenced-standards",
       completeness="full", use="evidence",
       limitation="Rendered by a local-jurisdiction viewer. Used only for referenced-standard EDITIONS, which local adoption does not change."),
  dict(id="ibc-2021-ch17", title="International Building Code 2021, Chapter 17 (UpCodes, Colorado viewer - no local 1705 amendments)",
       issuer="ICC", edition="2021", kind="code", folder="licensed", out="ibc2021_ch17.txt",
       original="ibc2021_ch17_colorado.html", extractor="html", url="https://up.codes/viewer/colorado/ibc-2021/chapter/17/special-inspections-and-tests",
       completeness="full", use="evidence",
       limitation="Used only to date when section numbers changed. Do NOT use the Connecticut 2021 viewer: it carries state amendments inside 1705.2."),
  dict(id="ibc-2018-ch17-hawaii", title="International Building Code 2018, Chapter 17, as amended by the Hawaii State Building Code (UpCodes, Hawaii viewer)",
       issuer="ICC", edition="2018 + Hawaii amendments", kind="code", folder="licensed", out="ibc2018_ch17_hawaii.txt",
       original="hi_ibc2018_ch17_upcodes.html", extractor="html", url="https://up.codes/viewer/hawaii/ibc-2018/chapter/17/special-inspections-and-tests",
       completeness="full", use="evidence",
       limitation="A third-party rendering with state amendments merged in. Where it matters, cite the State adoption document (hi-sbc-2018) for the amendment itself."),
  # ---- referenced standards (copyrighted) --------------------------------------------
  dict(id="aci-318-19-code", title="ACI 318-19 (SI), Building Code Requirements for Structural Concrete - CODE column",
       issuer="ACI", edition="318-19", kind="standard", folder="licensed", out="aci318_19_code.txt",
       original="ACI_318M_19.pdf", extractor="pdf-split-left", completeness="full", use="evidence",
       limitation="ACI prints code (left) and commentary (right) side by side. The split is by page gutter; full-width pages (front matter, some tables) may be split mid-line. Code text is what is mandatory."),
  dict(id="aci-318-19-commentary", title="ACI 318R-19 (SI) - COMMENTARY column (not mandatory)",
       issuer="ACI", edition="318R-19", kind="standard", folder="licensed", out="aci318_19_commentary.txt",
       original="ACI_318M_19.pdf", extractor="pdf-split-right", completeness="full", use="evidence",
       limitation="Commentary explains the code and is NOT a requirement. Evidence citing it must be typed as commentary."),
  dict(id="aws-d1-1-2020", title="AWS D1.1/D1.1M:2020 Structural Welding Code - Steel",
       issuer="AWS", edition="D1.1:2020", kind="standard", folder="licensed", out="aws_d1_1_2020.txt",
       original="AWS_D1_1_D1_1M_2020_Structural_Welding_C.pdf", extractor="pdf", completeness="full", use="evidence"),
  dict(id="aws-d1-4-2018", title="AWS D1.4/D1.4M:2018 Structural Welding Code - Steel Reinforcing Bars",
       issuer="AWS", edition="D1.4:2018", kind="standard", folder="licensed", out="aws_d1_4_2018.txt",
       original="AWS_D1_4_D1_4M.pdf", extractor="pdf", completeness="full", use="evidence"),
  dict(id="aws-d1-6-2017", title="AWS D1.6/D1.6M:2017 Structural Welding Code - Stainless Steel",
       issuer="AWS", edition="D1.6:2017", kind="standard", folder="licensed", out="aws_d1_6_2017.txt",
       original="AWS_D1_6_D1_6M_2017_Structural_Welding_C.pdf", extractor="pdf", completeness="full", use="evidence"),
  dict(id="aisc-360-22", title="ANSI/AISC 360-22 Specification for Structural Steel Buildings (the edition the 2024 IBC references)",
       issuer="AISC", edition="360-22", kind="standard", folder="licensed", out="aisc360_22.txt",
       original="pdfcoffee.com_aisc-360-22-16thedition-pdf-free.pdf", extractor="pdf", completeness="full", use="evidence"),
  dict(id="aisc-360-16", title="ANSI/AISC 360-16 Specification for Structural Steel Buildings (the edition the adopted 2018 IBC references)",
       issuer="AISC", edition="360-16", kind="standard", folder="licensed", out="aisc360_16.txt",
       original="AISC_360_2016_Specification_for_Structur.pdf", extractor="pdf", completeness="full", use="context",
       context_reason="Held to compare against 360-22 if DPP decides the matrix must be usable under the currently adopted 2018 IBC (Decision 4). No 2024 claim rests on it."),
  dict(id="upc-2018", title="IAPMO Uniform Plumbing Code 2018 (28th ed.)",
       issuer="IAPMO", edition="2018", kind="code", folder="licensed", out="upc_2018.txt",
       original="456452208-IAPMO-ANSI-UPC-1-2018-pdf.pdf", extractor="pdf", completeness="partial", use="evidence",
       limitation="A poor scan: much of the text layer is garbled or letter-spaced. Section 1319.4 could not be located. A negative search in this file proves nothing."),
  # ---- public law ---------------------------------------------------------------------
  dict(id="roh-16-1-1", title="Revised Ordinances of Honolulu Sec. 16-1.1, Building Code (amended through Ord. 25-11)",
       issuer="City & County of Honolulu", edition="as codified", kind="ordinance", folder="public", out="roh_16_1_1.txt",
       original="ROH 16-1.1.docx", extractor="docx", completeness="full", use="evidence"),
  dict(id="roh-16-1-2", title="Revised Ordinances of Honolulu Sec. 16-1.2, Residential Code",
       issuer="City & County of Honolulu", edition="as codified", kind="ordinance", folder="public", out="roh_16_1_2.txt",
       original="ROH 16-1.2.docx", extractor="docx", completeness="full", use="evidence"),
  dict(id="roh-16-2", title="Revised Ordinances of Honolulu Chapter 16 Article 2, Relocation of Buildings",
       issuer="City & County of Honolulu", edition="as codified", kind="ordinance", folder="public", out="roh_16_2.txt",
       original="ROH 16-2.docx", extractor="docx", completeness="full", use="evidence"),
  dict(id="roh-19-1", title="Revised Ordinances of Honolulu Sec. 19-1, Plumbing Code",
       issuer="City & County of Honolulu", edition="as codified", kind="ordinance", folder="public", out="roh_19_1.txt",
       original="ROH 19-1.docx", extractor="docx", completeness="full", use="evidence"),
  dict(id="hi-sbc-2018", title="Hawaii State Building Code as adopted 20 Apr 2021 (SBCC adoption document: 2018 IBC with state amendments)",
       issuer="State of Hawaii DAGS / SBCC", edition="adopted 2021-04-20", kind="state-rule", folder="public", out="hi_sbc_2018_adopted.txt",
       original="hi_sbc_2018_adopted.pdf", extractor="pdf", url="https://ags.hawaii.gov/wp-content/uploads/2021/09/2018StateBuildingCode_20210817.pdf",
       completeness="full", use="evidence"),
  # ---- issuer / accreditation pages, dated snapshots --------------------------------
  dict(id="ias-ac291", title="IAS AC291 Accreditation Criteria for Special Inspection Agencies (effective 2024-01-01), incl. Table 1 minimum inspector certifications",
       issuer="IAS", edition="2023-10-04, eff. 2024-01-01", kind="accreditation", folder="snapshots", out="ias_ac291_2024.txt",
       original="ias_ac291.pdf", extractor="pdf", url="https://cdn-v2.iasonline.org/wp-content/uploads/2023/11/AC291-Final.pdf",
       completeness="full", use="evidence"),
]
def page(id, title, issuer, original, url, use="evidence", completeness="full", limitation=None, context_reason=None):
    d = dict(id=id, title=title, issuer=issuer, edition="as published "+RETRIEVED, kind="issuer-page", folder="snapshots",
             out=id.replace("-", "_")+"_"+RETRIEVED+".txt", original=original, extractor="pdf" if original.endswith(".pdf") else "html",
             url=url, completeness=completeness, use=use)
    if limitation: d["limitation"] = limitation
    if context_reason: d["context_reason"] = context_reason
    CATALOG.append(d)
page("icc-si-exams", "ICC Special Inspector Certifications page", "ICC", "icc_si_exams.html",
     "https://www.iccsafe.org/professional-development/certifications-and-testing/special-inspector-exams/",
     limitation="Collapsed sections on the live page are included in this snapshot; Ctrl+F on the live page will not find them until expanded.")
page("icc-combo-designations", "ICC Combination Designations page", "ICC", "icc_combo.html", "https://www.iccsafe.org/credentialing/combo-designations/")
page("icc-cf-store", "ICC store: CF Commercial Fire Sprinkler Inspector", "ICC", "icc_cf_store.html", "https://shop.iccsafe.org/commercial-fire-sprinkler-inspector.html")
page("nicet-programs", "NICET Certification Programs page", "NICET", "nicet_programs.html", "https://www.nicet.org/certification-programs/")
page("nicet-geotech-manual", "NICET Geotechnical Engineering Technology Program Detail Manual (hosted PDF)", "NICET", "nicet_geotech.pdf",
     "https://www.nicet.org/nicetorg/assets/file/public/geotech.pdf",
     limitation="Establishes the program's structure and that no withdrawal notice is attached. Does NOT establish whether NICET still issues it.")
page("nfca-cap", "NFCA Contractor Accreditation Program page", "NFCA", "nfca_cap.html", "https://www.nfca-online.org/nfca-contractor-accreditation-program-cap-")
page("nfca-learning", "NFCA Programs & Exams page", "NFCA", "nfca_learning.html", "https://learning.nfca-online.org/")
page("icc-fire-inspector-1", "ICC store: 66 - Fire Inspector I", "ICC", "icc_fire_inspector_1.html", "https://shop.iccsafe.org/fire-inspector-i.html")
page("awc-learn-preservative", "AWCLearn course search: preservative", "AWC", "awc_learn_search.html", "https://learn.awc.org/course/search.php?search=preservative")
page("iapmo-pi", "IAPMO: UPC Residential and Commercial Plumbing Inspector (PI)", "IAPMO", "iapmo_pi.html",
     "https://iapmo.org/certification-testing/initial-certification/upc-residential-and-commercial-plumbing-inspector-pi")
page("iapmo-ppe", "IAPMO: UPC Residential and Commercial Plumbing Plans Examiner (PPE)", "IAPMO", "iapmo_ppe.html",
     "https://iapmo.org/certification-testing/initial-certification/upc-residential-and-commercial-plumbing-plans-examiner-ppe")
page("aabc-certification", "AABC Certification page", "AABC", "aabc_cert.html", "https://aabc.com/certification/", use="context",
     context_reason="Program landing page. The credential name is cited from AABC's own application form (aabc-tbt-application).")
page("aabc-tbt-application", "AABC Test and Balance Technician Certification Application (PDF)", "AABC", "aabc_tbt_application.pdf",
     "https://cdn.ymaws.com/members.aabc.com/resource/resmgr/aabc_applications/certified_tech_application.pdf")
page("pci-level-ii", "PCI Level II Quality Control Technician/Inspector Certification page", "PCI", "pci_level2.html",
     "https://www.pci.org/PCI/PCI/PCI-Certification/Personnel/Level-II.aspx")
page("pci-overview", "PCI Personnel Certification Program Overview", "PCI", "pci_overview.html",
     "https://www.pci.org/PCI/PCI/PCI-Certification/Personnel/Overview.aspx")
page("asse-6020", "ASSE 6020 Medical Gas Systems Inspector (certifying body page, MGTC)", "NFPA", "asse6020_mgtc.html",
     "https://mgtc.org/certifications/asse-6020-medical-gas-systems-inspector/")
page("nebb-certification", "NEBB Certification page", "NEBB", "nebb_cert.html", "https://www.nebb.org/certification/")
page("awci-eifs", "AWCI EIFS Inspector Certification page", "AWCI", "awci_eifs.html", "https://www.awci.org/eifs-inspector-certification/")
page("aci-shotcrete-inspector", "ACI Shotcrete Inspector certification program page", "ACI", "aci_shotcrete.html",
     "https://www.concrete.org/certification/certificationprograms.aspx?Program=Shotcrete%20Inspector")

# ---- declared but not held. Each names what its absence leaves open. -----------------
NOT_HELD = [
  dict(id="aisc-370-21", title="ANSI/AISC 370-21 Specification for Structural Stainless Steel Buildings", issuer="AISC",
       reason="not obtained", blocks="What IBC 1705.2.2 (stainless steel) requires of an inspector. The NEW-01 gap row cannot be given a code-backed credential."),
  dict(id="tms-402-602", title="TMS 402/602 Building Code Requirements and Specification for Masonry Structures", issuer="TMS",
       reason="not obtained", blocks="Whether the masonry rows' credentials are code-backed (IBC 1705.4)."),
  dict(id="sdi-qa-qc", title="SDI QA/QC Standard for Quality Control and Quality Assurance for Installation of Steel Deck", issuer="SDI",
       reason="not obtained", blocks="Welding inspector qualification for cold-formed steel deck, which IBC 1705.2.3 routes to SDI QA/QC."),
  dict(id="awci-12-b", title="AWCI Technical Manual 12-B (intumescent fire-resistive materials inspection)", issuer="AWCI",
       reason="not obtained", blocks="Whether IBC 1705.16 (via AWCI 12-B) names an inspector qualification."),
  dict(id="astm-e2174-e2393", title="ASTM E2174 / E2393 (firestop and joint system inspection)", issuer="ASTM",
       reason="not obtained", blocks="Whether IBC 1705.18 (via ASTM E2174/E2393) names an inspector qualification."),
  dict(id="aws-b5-1", title="AWS B5.1 Specification for the Qualification of Welding Inspectors", issuer="AWS",
       reason="not obtained", blocks="The exact WI / SWI / AWI definitions that AISC 360-22 N4 points to."),
  dict(id="aws-qc1", title="AWS QC1 Standard for AWS Certification of Welding Inspectors", issuer="AWS",
       reason="not obtained", blocks="Nothing on its own: D1.1 8.1.4.2 and 8.1.4.5 already state the CWI / CAWI tiers that reference it."),
  dict(id="ul-fm-firestop-exams", title="UL Solutions and FM Approvals firestop examination pages", issuer="UL / FM",
       reason="issuer pages blocked automated retrieval", blocks="The exams' current official names as UL and FM publish them. IAS AC291 Table 1 names both exams and is cited instead."),
  dict(id="hdoa-courses", title="Hawaii Department of Agriculture pesticide applicator certification and course listings", issuer="HDOA",
       reason="portal returned Forbidden", blocks="Whether HDOA courses 4322, 4415, 4422, 4484 exist and are current."),
  dict(id="dpp-registry-practice", title="What DPP's special inspector registry accepts in practice", issuer="DPP",
       reason="internal to DPP", blocks="Whether any credential set is the one DPP wants. IBC 1704.2.1 leaves competence to the building official."),
]

def sha(b): return hashlib.sha256(b).hexdigest()

def x_docx(path):
    x = zipfile.ZipFile(path).read("word/document.xml").decode("utf8")
    x = re.sub(r"</w:p>", "\n", x); x = re.sub(r"<w:tab/>", "\t", x); x = re.sub(r"<[^>]+>", "", x)
    return html.unescape(x)

def x_html(path):
    t = open(path, encoding="utf8", errors="ignore").read()
    t = re.sub(r"(?s)<script.*?</script>|<style.*?</style>|<noscript.*?</noscript>", "", t)
    t = re.sub(r"<br\s*/?>|</p>|</li>|</h\d>|</div>|</a>|</td>|</tr>|</section>", "\n", t)
    t = html.unescape(re.sub(r"<[^>]+>", " ", t))
    return "\n".join(re.sub(r"[ \t ]+", " ", l).strip() for l in t.splitlines() if l.strip()) + "\n"

def x_pdf(path):
    return subprocess.run(["pdftotext", "-layout", path, "-"], capture_output=True, check=True).stdout.decode("utf8", "replace")

def split_columns(text, side):
    """ACI prints code and commentary side by side. Find the gutter on each page and keep one side."""
    out = []
    for p in text.split("\f"):
        lines = p.split("\n"); best, score_best = 60, -1
        for g in range(45, 80):
            s = sum(1 if l[g-1:g+1] == "  " else 0.3 if (len(l) > g and (l[g-1] == " " or l[g] == " ")) else 0
                    for l in lines if len(l) > g)
            if s > score_best: best, score_best = g, s
        keep = []
        for l in lines:
            if len(l) <= best:
                if side == "left": keep.append(l.rstrip())
                continue
            m = re.search(r"\s{2,}", l[best-6:best+6]); cut = best-6+m.start() if m else best
            keep.append(l[:cut].rstrip() if side == "left" else l[cut:].strip())
        out.append("\n".join(keep))
    return "\f".join(out)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--originals", required=True, help="folder holding the supplied PDFs/docx")
    ap.add_argument("--snapshots", help="folder holding saved web pages (defaults to --originals)")
    a = ap.parse_args(); snaps = a.snapshots or a.originals
    entries = []
    for e in CATALOG:
        base = snaps if e["kind"] in ("issuer-page", "accreditation") or e["original"].endswith(".html") or e["id"] == "hi-sbc-2018" else a.originals
        src = os.path.join(base, e["original"])
        if not os.path.exists(src): sys.exit("missing original for %s: %s" % (e["id"], src))
        raw = open(src, "rb").read()
        ex = e["extractor"]
        if ex == "docx": text = x_docx(src)
        elif ex == "html": text = x_html(src)
        elif ex == "pdf": text = x_pdf(src)
        elif ex.startswith("pdf-split-"): text = split_columns(x_pdf(src), ex.rsplit("-", 1)[1])
        else: sys.exit("unknown extractor " + ex)
        header = "# %s\n# source id: %s | original: %s | original sha256: %s | retrieved: %s%s\n\n" % (
            e["title"], e["id"], e["original"], sha(raw), RETRIEVED, (" | " + e["url"]) if e.get("url") else "")
        body = (header + text).encode("utf8")
        os.makedirs(os.path.join(SRC, e["folder"]), exist_ok=True)
        rel = e["folder"] + "/" + e["out"]
        open(os.path.join(SRC, rel), "wb").write(body)
        entry = {k: e[k] for k in ("id", "title", "issuer", "edition", "kind", "completeness", "use") }
        entry.update(file=rel, sha256=sha(body), bytes=len(body),
                     origin=dict(original=e["original"], original_sha256=sha(raw), extractor=ex, retrieved=RETRIEVED, url=e.get("url")))
        for k in ("limitation", "context_reason"):
            if e.get(k): entry[k] = e[k]
        entries.append(entry)
        print("%-26s %9d bytes  %s" % (e["id"], len(body), rel))
    for n in NOT_HELD:
        entries.append(dict(n, kind="not-held", completeness="unavailable", use="declared"))
    man = dict(schema="dpp-si-source-catalog/2", document="Special Inspector Qualification 2024 IBC",
               compiled=RETRIEVED, field_notes={
        "completeness": "full = the whole document is in the file. partial = the file is incomplete or unreliable; 'limitation' says how.",
        "use": "evidence = at least one evidence entry must cite it (check C-02). context = held for reference, 'context_reason' says why. declared = not held; 'blocks' says what that leaves open.",
        "origin": "The original the text was extracted from. Re-running tools/extract_sources.py on the same original reproduces 'sha256' exactly."},
        sources=entries)
    json.dump(man, open(os.path.join(SRC, "manifest.json"), "w", encoding="utf8"), indent=1, ensure_ascii=False)
    print("catalog: %d held, %d declared not held" % (len(CATALOG), len(NOT_HELD)))

if __name__ == "__main__":
    main()
