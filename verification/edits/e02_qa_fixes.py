"""
Edit 02 - corrections from the independent QA pass of 2026-09-11 (handoff/QA-REPORT-2026-09-11.md).

What changes, and why, is written into the tab text itself so the workbook stands alone.
In summary:
  - ICC 49, EC and 86 restored as current ICC certifications (the rev18 tab called them errors).
  - Column references corrected: the DPOR columns are H and I, the certification columns L to P.
  - SI-004, SI-005, SI-006 flagged CONFLICT: steel fabrication rows that permit the DPOR for
    welding and bolting work, which ROH amend. (111) excludes.
  - SI-016 downgraded CONFLICT -> check: ACI 318-19 requires "a certified inspector"; the ACI
    anchor programs appear only in the commentary.
  - Firestop exams, AWC course, grading label, Fabricated Items "gap" corrected.
  - Credential basis (column Q) re-derived from the code chain, citing the clause.
  - Decisions 16 (City amendments are written against the 2018 IBC) and 17 (reading of the
    2018 footnote notation) added.
  - Credential check and Change Log get stable IDs (CC-nn, CL-nn) so evidence can point at them.
  - QA log tab removed; its history is in git and in the Change Log.
Certification columns L-P are NOT edited.
"""
import os, sys
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, os.path.join(HERE, "..", "tools"))
from xlsx import Workbook, table_sheet, read_all, STYLE_BODY

def refs(nums): return ", ".join("SI-%03d" % n for n in nums)

# ---------------------------------------------------------------------------------------
# Matrix 2024
# ---------------------------------------------------------------------------------------
BASIS = {
 "SI-003": "standard: ACI 318-19 26.12.1.1(d) certified field testing technicians (via IBC 1705.3) + DPP policy",
 "SI-004": "DPP policy",
 "SI-005": "standard: AISC 360-22 N4.2 (via IBC 1705.2.6 and 1705.2.1) + DPP policy",
 "SI-006": "standard: AISC 360-22 N4.2 (via IBC 1705.2.1) + DPP policy",
 "SI-007": "DPP policy",
 "SI-008": "DPP policy (AISC 360-22 N4.2 asks bolting inspectors only for documented training and experience)",
 "SI-009": "standard: AISC 360-22 N4.2 (via IBC 1705.2.1) + DPP policy",
 "SI-010": "unverified: SDI QA/QC not held (IBC 1705.2.3) + DPP policy",
 "SI-011": "DPP policy",
 "SI-012": "standard: ACI 318-19 26.12.1.1(d) certified field testing technicians (via IBC 1705.3)",
 "SI-013": "standard: ACI 318-19 26.12.1.1(d) certified field testing technicians (via IBC 1705.3) + DPP policy",
 "SI-014": "standard: AWS D1.4 9.1.2 (via IBC 1705.3.1) + DPP policy",
 "SI-015": "standard: ACI 318-19 26.12.1.1(d) certified field testing technicians (via IBC 1705.3) + DPP policy",
 "SI-016": "standard: ACI 318-19 26.13.1.5 and 26.13.1.6 certified inspector (via IBC 1705.3) + DPP policy",
 "SI-017": "unverified: TMS 402/602 not held (IBC 1705.4) + DPP policy",
 "SI-018": "unverified: TMS 402/602 not held (IBC 1705.4) + DPP policy",
 "SI-019": "DPP policy",
 "SI-020": "DPP policy", "SI-021": "DPP policy", "SI-022": "DPP policy", "SI-023": "DPP policy",
 "SI-026": "no credential listed", "SI-027": "DPP policy", "SI-028": "no credential listed",
 "SI-029": "DPP policy",
 "SI-030": "unverified: AWCI 12-B not held (IBC 1705.16) + DPP policy",
 "SI-031": "DPP policy",
 "SI-032": "unverified: ASTM E2174 / E2393 not held (IBC 1705.18.1 and 1705.18.2) + DPP policy",
 "SI-033": "code: IBC 1705.19.2 certification as air balancers + DPP policy",
 "SI-034": "DPP policy",
 "SI-024": "unverified: HDOA certification not held + DPP policy",
 "SI-025": "DPP policy",
 "SI-035": "ordinance: ROH 19-1 105.2.7 Registered Design Professional + DPP policy",
 "SI-036": "ordinance: ROH 19-1 105.2.7 Registered Design Professional + DPP policy",
 "SI-037": "ordinance: ROH 19-1 105.2.7 Registered Design Professional + DPP policy",
 "SI-038": "DPP policy",
}
FLAG_CHANGES = {"SI-004": "CONFLICT", "SI-005": "CONFLICT", "SI-006": "CONFLICT", "SI-016": "check"}

LEGEND_FLAGS = ("FLAGS.  COLLISION = the section number is used by both the City and the 2024 base code.  "
  "CONFLICT = the row permits something a code, standard or ordinance appears to restrict; the Flags tab says which.  "
  "GAP — no row = a 2024 section with no row yet.  check = needs DPP confirmation; nothing found contradicts it.  "
  "policy = department policy with no code citation.  Blank = checked, nothing found.  "
  "Shaded rows are not base IBC: City & County amendments (ROH 16-1.1), the plumbing rows, and department policy.  "
  "Every flag and every credential basis is backed by an entry on the Evidence tab (column R).")
LEGEND_BASIS = ("HOW THE CREDENTIAL BASIS COLUMN IS DERIVED: follow the code chain from the cited section.  "
  "\"code\" = the IBC section itself names the qualification.  \"standard\" = the section routes to a referenced standard "
  "that names it (clause given).  \"ordinance\" = a City ordinance sets it.  \"unverified\" = the chain leads to a standard "
  "that is not held (named in the cell).  \"DPP policy\" = the chain names no qualification, so the choice is the "
  "department's.  Most rows combine these with \"+\".")

# ---------------------------------------------------------------------------------------
# Flags - "What it is" text by Ref (None = keep existing text)
# ---------------------------------------------------------------------------------------
EXC = "ROH amend. (111) excludes the DPOR and their personnel from welding and high strength bolting inspection, but columns H and I permit them here"
FLAGS_TEXT = {
 "SI-004": "Was cited to 1705.1 (the general section); 1705.2.4 is the joist section. " + EXC +
           ", and the row's credentials are welding (S2, CWI). Confirm whether the exception applies to shop fabrication. See Decision 3.",
 "SI-005": "The 2018 sheet cited 1705.11, which in the 2018 IBC is Special Inspections for Wind Resistance. 1705.2.6 Metal Building Systems, "
           "with Table 1705.2.6, is new in the 2024 base code (the 2021 edition has no 1705.2.6) and routes to the structural steel sections. "
           "Same City welding exception as SI-004. See Decision 3.",
 "SI-006": "Structural steel fabrication. Same City welding exception as SI-004: columns H and I permit the DPOR and their personnel, "
           "and the credentials are welding and bolting (S1, S2, CWI). IBC 1705.2.1 routes to AISC 360 (360-22 is the edition the 2024 IBC "
           "references), whose N4.2 sets QA inspector qualifications. See Decision 3.",
 "NEW-04": "Not new: this was 1705.10 Fabricated Items in the 2018 IBC. The five Fabricator rows (SI-003 to SI-007) already cover fabrication "
           "inspection but cite material sections; 1705.11 routes to 1704.2.5. Confirm whether a separate row is wanted. See Decision 5.",
 "SI-008": "ROH amend. (111) lets the DPOR and their personnel act as special inspector \"with the exception of welding and high strength bolting.\" "
           "Columns H and I permit them on this row. The exception is City-only: the 2024 IBC 1704.2.1 and the Hawaii State version of 1704.2.1 "
           "have none, so it reaches a 2024 adoption only if the City re-adopts (111). AISC 360-22 N4.2 asks bolting inspectors only for documented "
           "training and experience. See Decisions 3 and 16.",
 "SI-009": "Same City exception (welding), columns H and I. The 2018 third-party cell reads \"SE1 & 2\" (ICC + ACI) but this row has no ACI credential; "
           "likely meant \"& 4\" (AWS). AISC 360-22 N4.2 qualifies QA welding inspectors as AWS B5.1 WI or SWI (AWI only under direct WI supervision) "
           "or under AWS D1.1 8.1.4.2(5). See Decision 3.",
 "SI-010": "2024 inserts 1705.2.2 Structural Stainless Steel (in 2021, 1705.2.2 was Cold-Formed Steel Deck), shifting these down one. IBC 1705.2.3 routes "
           "welding inspector qualification for steel deck to SDI QA/QC, which is not held. The 2018 supervised-staff cell requires no credential "
           "(\"X\") while the registered SI cell requires ICC (\"X1\"); confirm that is deliberate.",
 "NEW-01": "New in 2024 (in 2021, 1705.2.2 was Cold-Formed Steel Deck). Routes to AISC 370 (370-21 per Chapter 35), which is not held. Stainless "
           "welding is AWS D1.6, whose 8.1.4.1 accepts CWI, CWB / CSA W178.2, or training and experience.",
 "NEW-02": "New, with Table 1705.5.3. ICC Category 93 Tall Mass Timber Buildings Special Inspector exists.",
 "NEW-05": "New. Periodic special inspection of sealants or adhesives required by Section 703.7. ICC Category 93 likely applies. "
           "The City uses this number for termite protection (SI-024).",
 "SI-012": "The 2018 third-party cell names licences (\"SE, AR, CE\") with no credential, while the registered SI cell requires ACI (\"X2\"). Confirm.",
 "SI-013": "The 2018 sheet cited 1705.3.1, which is titled \"Welding of Reinforcing Bars\"; that fits the welding sub-row, not this one.",
 "SI-011": "The 2018 sheet cited \"1705.5; 1705.1; 1705.2\". The specific subsections are 1705.5.1 and 1705.5.2. Left as flagged in case the row "
           "deliberately covers steel truss elements.",
 "SI-014": "City welding exception applies (columns H and I). IBC 1705.3.1 itself requires rebar welding inspection and inspector qualification per "
           "AWS D1.4, and D1.4 9.1.2 accepts AWS CWI, so the listed CWI satisfies it. See Decisions 3 and 8.",
 "SI-015": "Pre-stressed Concrete (92) and Reinforced Concrete (49) are both current ICC certifications, and 92 requires 49. The PCI credential's "
           "printed name is \"Level II Quality Control Technician/Inspector\".",
 "SI-016": "ACI 318-19 26.13.1.6 requires a certified inspector for adhesive anchors; 26.13.1.5 allows a certified inspector, or a qualified inspector "
           "approved by the licensed design professional and the building official, for mechanical anchors. The ACI anchor programs (CPP 681.1, "
           "681.2) are named only in the commentary, as examples. The registered SI cell requires ICC Reinforced Concrete, a certification that "
           "IAS AC291 accepts for post-installed anchors. See Decision 9.",
 "SI-017": "2024 routes masonry inspection to TMS 402 and TMS 602, which are not held.",
 "SI-018": "City welding exception applies (columns H and I). Also: the 2018 supervised-staff cell requires ICC only (\"X1\") while SI-014 requires "
           "ICC + AWS (\"X1 & 4\"); same sub-category, different answer. See Decisions 3 and 11.",
 "SI-021": "+ Table 1705.7. 1705.10 Structural Integrity of Deep Foundation Elements also applies; no row exists for it.",
 "NEW-03": "Added in the 2021 edition. This insertion is what shifts every later 1705 section by +1. Applies alongside the rows for 1705.7 to 1705.9.",
 "SI-029": "2018: 1705.14 Sprayed Fire-Resistant Materials. 2024: 1705.15 Sprayed Fire-Resistive Materials (SFRM).",
 "SI-030": "2018: 1705.15 Mastic and Intumescent Fire-Resistant Coatings. 2024: 1705.16 Intumescent Fire-Resistive Materials, which drops "
           "\"Mastic and\" and references AWCI 12-B (not held). Consider retitling the classification.",
 "SI-033": "1705.19.2 is code-backed: approved agencies for smoke control testing need \"certification as air balancers\", which the Other column "
           "provides. Collides with ROH 1705.19 (SI-034). ROH 919.5 still cites 1705.18, which was Testing for Smoke Control in the 2018 IBC.",
 "SI-034": "ROH amend. (112). Collides with IBC-2024 1705.19 Testing for Smoke Control (SI-033). DPP to renumber; the 2024 base code stops at "
           "1705.20 and ROH uses 1705.21, so the first free number is 1705.22.",
 "SI-024": "ROH amend. (113); the 2018 sheet read \"1705.2\" (dropped zero). Collides with IBC-2024 1705.20 Sealing of Mass Timber. Residential "
           "counterpart is ROH 16-1.2 R109.5.2. HDOA credentials are Department of Agriculture licensing, a separate statute, not checked.",
 "SI-025": "ROH amend. (113). 1705.21 requires special inspection where soil conditions warrant a geotechnical investigation or are specified in ROH "
           "Chapter 18A, and amend. (114) places grading, excavation and earthwork under Chapter 18A. The \"Special Grading\" label is therefore "
           "defensible; confirm the intended scope. 1705.21 also blocks that number as a renumbering target. See Decisions 2 and 6.",
 "SI-035": "Citation updated from the 2018 sheet's \"2021 Uniform Plumbing Code Chapter 13\". ROH 19-1 adopts the Hawaii State Plumbing Code, which "
           "adopts the Uniform Plumbing Code, 28th Edition (2018), 6th Printing. DPP may be applying the 2021 UPC deliberately; recorded, not "
           "overwritten. ROH 19-1 105.2.7 requires plumbing special inspections by Registered Design Professionals licensed in Hawaii and not "
           "involved in the work. The UPC 2018 section 1319.4 itself could not be located in the scan held. See Decision 12.",
 "SI-036": "Citation updated from \"2021 UPC Chapter 15\". Same edition and framework notes as SI-035. UPC 2018 1501.2 is \"System Design\". "
           "ROH 19-1 amend. (80) deletes UPC Chapter 14 entirely.",
 "SI-037": "Citation updated from \"2021 UPC Chapter 16\". Same edition and framework notes as SI-035. UPC 2018 1602.2 is \"Plumbing Plan Submission\".",
 "SI-038": "Labelled \"Pending DPP Policy\" in the 2018 sheet.",
}

# ---------------------------------------------------------------------------------------
# Decisions
# ---------------------------------------------------------------------------------------
ALL_REFS = "all rows"
DECISIONS = [
 (1, "Renumber the two colliding City sections.",
  "ROH 1705.19 \"Fire-protection systems\" collides with IBC-2024 1705.19 \"Testing for Smoke Control\". ROH 1705.20 \"Termite protection\" collides "
  "with IBC-2024 1705.20 \"Sealing of Mass Timber\". Both appear in the matrix printing the same number as a different requirement.",
  "The 2024 base code stops at 1705.20 and ROH already uses 1705.21, so the first free numbers are 1705.22 and 1705.23. The ordinance text has to "
  "move with them. NO NUMBER HAS BEEN ASSIGNED IN THIS FILE.", "SI-033, SI-034, SI-024"),
 (2, "Does ROH 1705.21 stay where it is?", "It does not collide, but it sits directly above the obvious renumbering targets.",
  "Confirm it stays, or move all three together.", "SI-025"),
 (3, "The welding / high strength bolting exception.",
  "ROH amend. (111) lets the DPOR/EOR and their personnel act as special inspector \"with the exception of welding and high strength bolting.\" "
  "Seven rows permit them in columns H and I for welding or bolting work: the two structural steel rows, both Welding of Reinforcement rows, "
  "and the three steel Fabricator rows.",
  "Either those cells should read — (not permitted), or the exception is read more narrowly (for example, not applying to shop fabrication). "
  "This is a permissions question, not a formatting one.", refs([4, 5, 6, 8, 9, 14, 18])),
 (4, "Target edition.",
  "ROH 16-1.1 (amended through Ord. 25-11) adopts the Hawaii State Building Code of 20 April 2021, which adopts the 2018 IBC. Is this matrix staged "
  "for a future adoption or meant to be usable now?",
  "If usable now, the edition labels revert to 2018 and only the flags survive.", ALL_REFS),
 (5, "Add rows for the 2024 sections that have none?",
  "Four sections new since the 2018 IBC have no row: 1705.2.2 Structural Stainless Steel, 1705.5.3 Mass Timber Construction, 1705.10 Structural "
  "Integrity of Deep Foundation Elements, 1705.20 Sealing of Mass Timber. 1705.11 Fabricated Items is not new (it was 1705.10 in 2018) and may "
  "already be covered by the five Fabricator rows.",
  "Marked as gaps, with no credentials assigned. ICC Category 93 Tall Mass Timber Buildings Special Inspector covers the two timber sections.",
  "NEW-01, NEW-02, NEW-03, NEW-04, NEW-05"),
 (6, "The \"Special Grading, Excavation/Filling\" label.",
  "ROH 1705.21 is titled \"Soils and foundation\" and requires special inspection where conditions are specified in ROH Chapter 18A, which amend. "
  "(114) makes the chapter for excavation, grading and earthwork. The label is defensible but broader than the title.",
  "Confirm the intended scope, or retitle the classification to match the section.", "SI-025"),
 (7, "CAWI listed alongside CWI in seven rows.",
  "AWS D1.1:2020 8.1.4.2 qualifies the Inspector as CWI or SCWI; 8.1.4.5 qualifies CAWI (or higher) as an Assistant Inspector, and 8.1.4.4 has "
  "assistants work under the Inspector's supervision. AISC 360-22 N4.2 likewise permits an Associate Welding Inspector only under the direct "
  "supervision of a Welding Inspector on the premises. The 2018 cell joins them with \"&\".",
  "If the intent is that a CAWI may serve alone, that conflicts with both standards.", refs([4, 5, 6, 8, 9, 14, 18])),
 (8, "Rebar welding inspector qualification.",
  "IBC 1705.3.1 and ACI 318-19 26.13.1.4 both route reinforcing bar welding inspection to AWS D1.4 (2018-AMD1 per Chapter 35). D1.4 9.1.2 accepts "
  "AWS CWI, CWB / CSA W178.2, or training and experience.",
  "The listed CWI already satisfies D1.4. No change needed unless DPP wants the rows to name D1.4.", "SI-014, SI-018"),
 (9, "Post-installed anchor inspector qualification.",
  "ACI 318-19 26.13.1.6 requires a certified inspector for adhesive anchors; 26.13.1.5 allows a certified inspector or an approved qualified "
  "inspector for mechanical anchors. The code names no program: the commentary gives ACI CPP 681.1 and 681.2 as examples, \"or similar program\". "
  "IAS AC291 accepts ICC Reinforced Concrete SI or ACI Concrete/Masonry Construction SI for this work.",
  "Confirm whether ICC Reinforced Concrete plus ACI Field Testing Technician is the intended certified-inspector route, or name the ACI anchor programs.",
  "SI-016"),
 (10, "ICC certification names in column L.",
  "ICC lists special inspector certifications by category (47, 48, 49, 84, 86, 92, 93, EC, S1, S2), each earned through a General Requirements "
  "module plus codes and plans exams (for example 47C and 47P). Category 49 Reinforced Concrete Special Inspector exists and is the prerequisite "
  "for 92. The building inspector exams are B1 and B2; the 2018 sheet writes 1B and 2B.",
  "The category numbers in column L are current ICC names. Only 1B / 2B need correcting, to B1 / B2.", refs([7, 10, 11, 24, 27])),
 (11, "Two rows for the same sub-category disagree - internal to DPP policy.",
  "Both Welding of Reinforcement rows carry the same AWS credential and the same registered SI requirement, but the supervised-employee cell requires "
  "AWS on one (\"X1 & 4\") and not the other (\"X1\").", "Only the department can reconcile it. Flagged, not changed.", "SI-014, SI-018"),
 (12, "The plumbing rows - edition, and who may inspect.",
  "The 2018 sheet cites the \"2021 Uniform Plumbing Code\". ROH 19-1 adopts the Hawaii State Plumbing Code, which adopts the Uniform Plumbing Code, "
  "28th Edition (2018), 6th Printing. The special inspection requirements are City amendments: 1319.4 (medical gas and vacuum), 1501.2 (alternate "
  "water source), 1602.2 (nonpotable rainwater catchment). ROH 19-1 105.2.7 requires plumbing special inspections by Registered Design "
  "Professionals licensed in Hawaii and not involved in the work.",
  "(a) Confirm whether the 2021 UPC reference is intentional. (b) Confirm how the who-may-perform cells should read given 105.2.7, which is narrower "
  "than the building-side pool; the IAPMO certifications in the Other column appear nowhere in ROH 19-1.", "SI-035, SI-036, SI-037"),
 (13, "The SI FORM acknowledgement citation.",
  "The SI FORM cites \"Section 1704.4 ... Access for Special Inspection\" and \"ROH Chapter 16-2 Section R109.6\". In the 2024 IBC, 1704.4 is "
  "Contractor Responsibility and Access for Special Inspection is 1704.2.2. ROH Chapter 16 Article 2 is Relocation of Buildings; the residential "
  "code is ROH 16-1.2, where R109.6 is Building permit requirement and R109.8 is Contractor responsibility.",
  "The sentence cites the wrong subsections and the wrong ROH article. NOT changed: the SI FORM tab was left untouched.", "SI FORM tab"),
 (14, "Add NEBB to the acronym table.",
  "The smoke control row requires \"Nat'l Environmental Balancing Bureau\" certification, spelled out rather than abbreviated, and IBC-2024 1705.19.2 "
  "backs it: approved agencies need \"certification as air balancers\". NEBB is the only organisation in the matrix with no acronym-table entry.",
  "Housekeeping. Add NEBB alongside AABC.", "SI-033"),
 (15, "Does DPP accept the other qualification bases the welding codes allow?",
  "AWS D1.1:2020 8.1.4.2 lists five bases for a welding inspector, including CSA W178.2 Level 2 or 3 and ASNT SNT-TC-1A-VT Level II; D1.4 9.1.2 "
  "and D1.6 8.1.4.1 accept CWI, CWB / CSA W178.2, or training and experience. The matrix names only AWS CWI.",
  "NOT A DEFECT. A jurisdiction may be more stringent than the standard it adopts. Recorded so the choice is visible to an applicant with a CSA- or "
  "ASNT-qualified inspector.", refs([4, 5, 6, 8, 9, 14, 18])),
 (16, "City amendments are written against the 2018 IBC.",
  "Every ROH 16-1.1 amendment amends the 2018 IBC as adopted by the State. Neither the 2024 IBC 1704.2.1 nor the Hawaii State version of 1704.2.1 "
  "has a welding or bolting exception; it exists only in ROH amend. (111). ROH 1705.19 to 1705.21 are likewise City additions. This matrix assumes "
  "the City will re-adopt its amendments against the 2024 edition.",
  "Confirm the City intends to carry these forward. If not, the CONFLICT flags for the exception and the COLLISION flags fall away.",
  refs([4, 5, 6, 8, 9, 14, 18]) + ", SI-033, SI-034, SI-024, SI-025"),
 (17, "Reading the 2018 footnote notation.",
  "The 2018 sheet keys credentials to footnotes (headers ICC1, ACI2, NICET3, AWS4, Other5) but footnotes 1 to 6 are not in the file. This matrix "
  "reads a comma as AND in the employee and registered SI columns (\"X1, 2, & 3\" = ICC + ACI + NICET) and as OR between professions in the "
  "third-party column (\"SE1 & 4, AR1 & 4, IAS\" = SE or AR, each with ICC + AWS, or an IAS agency). All 144 permission cells were translated "
  "this way and are checked against the 2018 sheet.",
  "Confirm the reading, and supply footnotes 1 to 6 if they exist.", ALL_REFS),
]

# ---------------------------------------------------------------------------------------
# Credential check
# ---------------------------------------------------------------------------------------
BB = refs([7, 10, 11, 24, 27]); WELD = refs([4, 5, 6, 8, 9, 14, 18]); SOIL = refs([20, 21, 22, 23, 25])
CRED = [
 ("CC-01", "ICC", refs([3, 13, 16]), "Reinforced Concrete (47)",
  "Category 47 Reinforced Concrete: General Requirements module + 47C Reinforced Concrete Codes + 47P Reinforced Concrete Plans", "current", "Current certification name."),
 ("CC-02", "ICC", refs([15]), "Reinforced Concrete (49)",
  "Category 49 Reinforced Concrete Special Inspector: Category 47 plus ACI Concrete Field Testing Technician Grade I", "current",
  "Consistent with ICC's prerequisite chain: Category 92 requires Category 49. (rev18 called 49 an error; that was wrong.)"),
 ("CC-03", "ICC", refs([15]), "Pre-stressed Concrete (92)", "Category 92 Prestressed Concrete: Category 49 + 92C + 92P", "current", ""),
 ("CC-04", "ICC", refs([17]), "Structural Masonry (84)", "Category 84 Structural Masonry: General Requirements + 84C + 84P", "current", ""),
 ("CC-05", "ICC", SOIL, "Soils (EC)", "Category EC Soils: General Requirements + ECC Soils Codes + ECP Soils Plans", "current",
  "EC is the certification; ECC and ECP are its exams. (rev18 called EC an error; that was wrong.)"),
 ("CC-06", "ICC", refs([4, 5, 6, 8, 14, 18]), "Structural Steel & Bolting (S1)", "Category S1 Structural Steel and Bolting: General Requirements + S1C + S1P", "current", ""),
 ("CC-07", "ICC", refs([4, 5, 6, 9, 14, 18]), "Structural Steel Welding (S2)", "Category S2 Structural Welding: Category S1 + S2C + S2P", "current", "S2 requires S1."),
 ("CC-08", "ICC", refs([29, 30]), "Spray-Applied Fireproofing (86)", "Category 86 Spray-applied Fireproofing: General Requirements + exam 86 (codes and plans)",
  "current", "No \"86M\" designation appears on ICC's page. (rev18 said 86M.)"),
 ("CC-09", "ICC", BB, "Commercial Bldg Inspector (2B)", "B2 Commercial Building Inspector", "naming", "The exam ID is B2, not 2B."),
 ("CC-10", "ICC", BB, "Residential Bldg Inspector (1B)", "B1 Residential Building Inspector", "naming", "The exam ID is B1, not 1B."),
 ("CC-11", "ICC", refs([29, 30]), "Fire Inspector I (66)", "66 - Fire Inspector I", "current", ""),
 ("CC-12", "ICC", refs([34]), "Commercial Fire Sprinkler (CF)", "CF - Commercial Fire Sprinkler Inspector", "current", "Exact match."),
 ("CC-13", "ICC", "NEW-02, NEW-05", "(none listed)", "Category 93 Tall Mass Timber Buildings: General Requirements + exam 93", "current",
  "Candidate for the two mass timber gap rows."),
 ("CC-14", "ACI", refs([3, 12, 13, 15, 16, 17, 21, 22, 23]), "Conc Construction & Conc Field Testing Technician Grade 1",
  "ACI 318-19 26.12.1.1(d): certified field testing technicians perform tests on fresh concrete at the job site. ICC Category 49 requires ACI Grade I.",
  "current", "Code-backed for testing, via IBC 1705.3."),
 ("CC-15", "ACI", refs([16]), "Conc Field Testing Technician Grade 1, for post-installed anchors",
  "ACI 318-19 26.13.1.6: adhesive anchors need a certified inspector. 26.13.1.5: mechanical anchors need a certified inspector or an approved "
  "qualified inspector. ACI CPP 681.1 / 681.2 are named only in the commentary, as examples. IAS AC291 accepts ICC Reinforced Concrete SI.",
  "SCOPE — check", "See Decision 9."),
 ("CC-16", "ACI", refs([19]), "Shotcrete SI", "ACI Shotcrete Inspector certification program", "current",
  "Full name \"Shotcrete Inspector\". ACI 318-19 does not name a shotcrete inspector certification."),
 ("CC-17", "AWS", WELD, "Certified Welding Inspector",
  "AWS D1.1:2020 8.1.4.2(1): CWI or SCWI qualifies the Inspector. AWS D1.4 9.1.2(1): CWI.", "current", ""),
 ("CC-18", "AWS", WELD, "Certified Assoc Welding Inspector",
  "AWS D1.1:2020 8.1.4.5(1): CAWI or higher qualifies an Assistant Inspector, who works under the Inspector's supervision (8.1.4.4). "
  "AISC 360-22 N4.2: AWI only under direct supervision of a WI.", "SCOPE — check", "See Decision 7."),
 ("CC-19", "AWS", refs([14, 18]), "(welding inspector for reinforcing bars)",
  "IBC 1705.3.1 routes to AWS D1.4. D1.4 9.1.2 accepts CWI, CWB / CSA W178.2, or training and experience.", "current",
  "The listed CWI satisfies D1.4. See Decision 8."),
 ("CC-20", "NICET", SOIL, "Geotechnical engineering technology",
  "Not on NICET's current certification programs page. NICET still hosts the program manual (published June 1994). IAS AC291, effective 2024, "
  "still lists NICET Geotechnical Engineering Technology Levels I to IV as acceptable.", "UNCONFIRMED",
  "Whether NICET still issues it needs NICET. Existing holders appear to remain accepted."),
 ("CC-21", "NICET", SOIL, "Lvl II", "IAS AC291 Table 1: NICET II (geotechnical or construction or construction material testing or soils)", "current", ""),
 ("CC-22", "NICET", refs([3]), "NICET Lvl III with 2 yrs experience", "Level exists; no program named in the cell", "naming", "Which NICET program is not stated."),
 ("CC-23", "NICET", refs([34]), "NICET: a) Level III; b) Level II",
  "NICET program \"Inspection and Testing of Water-Based Systems\" (distinct from \"Water-Based Systems Layout\")", "naming", "Program not named in the cell."),
 ("CC-24", "PCI", refs([15]), "PCI - Inspector Level II", "PCI Level II Quality Control Technician/Inspector Certification", "naming", "Printed name differs."),
 ("CC-25", "NFCA", refs([29, 30]), "NFCA - CAP Program Education & Exam",
  "CAP accredits fireproofing contractors. NFCA's Fireproofing Education and Exams are offered for CAP, UL Designated Responsible Individuals, and "
  "Inspectors. IAS AC291 lists ICC Spray-Applied Fireproofing SI or ICC Fire Inspector I for this work, not NFCA.", "SCOPE — check",
  "The education and exam is an inspector route; CAP itself is not."),
 ("CC-26", "AWCI", refs([31]), "AWCI - EIFS", "AWCI EIFS Inspector certification; IAS AC291 names \"Exterior Insulated Finish System Inspector (EIFS-I)\"",
  "naming", "The cell names it loosely."),
 ("CC-27", "AABC", refs([33]), "AABC - Tech Certification", "AABC Test and Balance Technician certification", "naming",
  "IBC 1705.19.2 requires certification as air balancers."),
 ("CC-28", "NEBB", refs([33]), "Nat'l Environmental Balancing Bureau - Tech Certification",
  "NEBB certifies Professionals and Technicians. IAS AC291 accepts AABC, NEBB or equivalent balancing technician certification.", "naming",
  "The exact NEBB credential title is not pinned. See Decision 14."),
 ("CC-29", "IAPMO", refs([35, 36, 37]), "UPC Residential & Commercial Plumbing Inspector (PI)", "UPC Residential and Commercial Plumbing Inspector (PI)", "current", "Exact match."),
 ("CC-30", "IAPMO", refs([38]), "Plans Examiner (PPE)", "UPC Residential and Commercial Plumbing Plans Examiner (PPE)", "current", ""),
 ("CC-31", "UL / FM", refs([32]), "UL - Firestop Exam, & Factory Mutual Firestop Exam",
  "IAS AC291 Table 1, fire-resistant penetrations and joints: ICC CLA or UL Firestop Examination or FM Firestop Examination (or the IFC exam). "
  "The exams are taken by individuals; FM 4991 and UL's qualified contractor program are separate company accreditations.", "current",
  "rev18 called these contractor-only; the exams are recognised inspector qualifications."),
 ("CC-32", "AWC", refs([7]), "AWC Durable Construction with Preservative-Treated Wood",
  "A course on AWC's own learning portal (AWCLearn), not a certification", "naming",
  "The issuer is AWC; the item is a course. DPP to confirm what an applicant presents."),
 ("CC-33", "HDOA", refs([24]), "Certification of Pesticide Applicators in Hawaii; HDOA courses #4322, 4415, 4422, 4484",
  "Not checked: the HDOA portal returned Forbidden.", "not checked", "Separate statute; HDOA to confirm."),
 ("CC-34", "NFPA", refs([35]), "NFPA 99 Certification",
  "NFPA 99 is the Health Care Facilities Code. ASSE 6020 Medical Gas Systems Inspector is a personnel certification whose training covers NFPA 99.",
  "SCOPE — check", "The cell names a code where a credential belongs."),
]

# ---------------------------------------------------------------------------------------
# Change Log
# ---------------------------------------------------------------------------------------
CHANGELOG = [
 ("CL-01", "Structure", "Rows re-ordered by authority layer, then section number.",
  "Base IBC first, then City & County amendments, then the plumbing code rows, then department policy. The \"2018 row\" column on Flags maps every row back to its original position."),
 ("CL-02", "Structure", "Added Code and Section title columns.",
  "A bare section number could previously mean base IBC, a state amendment, a City amendment or department policy. The Code column separates them, and the section title is reproduced as printed so a number can be checked against its words."),
 ("CL-03", "Structure", "Superscript pointer notation replaced with explicit text.",
  "\"X1, 2, & 3\" now reads \"✓ ICC + ACI + NICET\". \"SE1 & 4, AR1 & 4, IAS\" now reads \"SE or AR (ICC + AWS), or IAS agency\". The reading of the 2018 shorthand is an interpretation; see Decision 17."),
 ("CL-04", "Structure", "Credential columns reproduced verbatim.",
  "Columns L through P are copied cell-for-cell from the 2018 sheet. No credential has been edited. Findings about them are on Decisions and Credential check."),
 ("CL-05", "Structure", "Credential basis column (Q).",
  "Re-derived 2026-09-11 from the code chain: IBC section, then any referenced standard, then the clause naming an inspector qualification. \"unverified\" where the standard is not held."),
 ("CL-06", "Citation", "Steel Joists & Joist Girders: 1705.1 -> 1705.2.4", "The 2018 cell pointed at the general section."),
 ("CL-07", "Citation", "Metal Buildings: 1705.11 -> 1705.2.6",
  "1705.2.6 Metal Building Systems and Table 1705.2.6 are new in the 2024 base code. In the 2018 IBC, 1705.11 is Special Inspections for Wind Resistance."),
 ("CL-08", "Citation", "Cold-Formed Steel: 1705.2.2/.3/.4 -> 1705.2.3/.4/.5", "The 2024 IBC inserts 1705.2.2 Structural Stainless Steel."),
 ("CL-09", "Citation", "Wind 1705.11 -> 1705.12; Seismic 1705.12 -> 1705.13; SFRM 1705.14 -> 1705.15; Intumescent 1705.15 -> 1705.16; EIFS 1705.16 -> 1705.17; Penetrations 1705.17 -> 1705.18; Smoke control 1705.18 -> 1705.19.",
  "The +1 shift. Caused by 1705.10 Structural Integrity of Deep Foundation Elements, added in the 2021 IBC. Honolulu adopts the 2018 IBC, so the 2021 and 2024 changes arrive together."),
 ("CL-10", "Citation", "Termite Protection: 1705.2 -> ROH 1705.20", "The 2018 sheet read \"1705.2\", a dropped zero. ROH 16-1.1 amend. (113)."),
 ("CL-11", "Citation", "Water-Based Fire Protection: kept at ROH 1705.19", "ROH amend. (112) \"Fire-protection systems\". Left at its own number; renumbering is DPP's call."),
 ("CL-12", "Citation", "Special Grading: kept at ROH 1705.21", "Citation correct. The section is titled \"Soils and foundation\" and refers to ROH Chapter 18A; see Decision 6."),
 ("CL-13", "Title", "Two title changes worth noting.",
  "1705.15 is now \"Sprayed Fire-Resistive Materials (SFRM)\" (2018: Sprayed Fire-Resistant Materials). 1705.16 is now \"Intumescent Fire-Resistive Materials\" (2018: Mastic and Intumescent Fire-Resistant Coatings)."),
 ("CL-14", "Not changed", "SI FORM and blurb tabs.",
  "SI FORM: only its three dropdown validation ranges changed (see SI FORM dropdown tab). blurb: unchanged. Neither tab's text was touched."),
 ("CL-15", "Not changed", "Drop-down tab column B, rows 2 to 25.",
  "Rows 26 to 29 were added (see SI FORM dropdown tab). Column C, added 2026-09-11, lists the Matrix 2024 rows each entry covers."),
 ("CL-16", "Not changed", "Every credential in columns L through P.", "See Decisions 7 through 10 and 15 for what was found but not applied."),
 ("CL-17", "Repair", "2026-09-11: workbook repaired so Excel opens it.",
  "Excel 16 refused to open rev18, even in repair mode. Causes: Drop-down rows written out of order; the Matrix filter written in the wrong place in the file; a tab name over Excel's 31-character limit (now \"SI Qual 2018 (as received)\"). No values changed."),
 ("CL-18", "QA", "2026-09-11: corrections from the independent QA pass.",
  "ICC 49, EC and 86 restored as current; column references corrected (H and I; L through P); firestop exam, AWC course, anchor requirement and grading label findings corrected; SI-004 to SI-006 flagged for the City welding exception; SI-016 downgraded to check; credential basis re-derived; Decisions 16 and 17 added; README and Evidence tabs added; QA log tab removed. Report: handoff/QA-REPORT-2026-09-11.md in the repository."),
]

DROPDOWN_MAP = {
 2: "SI-008", 3: "SI-009, SI-014, SI-018", 4: "SI-010", 5: "SI-011", 6: "SI-012", 7: "SI-016", 8: "SI-015", 9: "SI-013", 10: "SI-017",
 11: "SI-019", 12: "SI-021, SI-022, SI-023", 13: "SI-024", 14: "SI-025", 15: "SI-026", 16: "SI-027", 17: "SI-028", 18: "SI-033",
 19: "SI-029, SI-030", 20: "SI-031", 21: "SI-032", 22: "SI-034", 23: "SI-035", 24: "SI-036", 25: "SI-037",
 26: "SI-003, SI-004, SI-005, SI-006, SI-007", 27: "SI-020", 28: "SI-038", 29: "—",
}
DROPDOWN_NOTE = ("WHAT THIS TAB IS FOR: column B is the pick-list behind the Classification dropdown on the SI FORM tab (validation range B1:B29, applied "
  "to SI FORM cells A9:A21). Do not insert, delete or reorder rows in column B. Column A is reference only. Column C lists the Matrix 2024 rows "
  "each entry covers; check D-02 fails if a matrix row has no entry. If a classification is added or renamed in Matrix 2024, update columns B "
  "and C together and extend the validation range.")


def apply(path_in, path_out):
    cur = read_all(path_in)
    w = Workbook(path_in)
    M = cur["Matrix 2024"]
    row_of = {M[k]: int(k[1:]) for k in M if k.startswith("B") and k[1:].isdigit() and int(k[1:]) >= 3 and M[k].startswith(("SI-", "NEW-"))}

    # Matrix: flags, basis, legend
    cells = {}
    for ref, fl in FLAG_CHANGES.items(): cells["A%d" % row_of[ref]] = fl
    for ref, b in BASIS.items(): cells["Q%d" % row_of[ref]] = b
    cells["A46"] = LEGEND_FLAGS; cells["A48"] = LEGEND_BASIS
    w.set_cells("Matrix 2024", cells)

    # Flags: regenerate, adding SI-006, in matrix order
    F = cur["Flags"]
    old = {F["A%d" % r]: r for r in range(3, 100) if "A%d" % r in F}
    rows = []
    for ref, r in sorted(row_of.items(), key=lambda kv: kv[1]):
        m_flag = FLAG_CHANGES.get(ref, M.get("A%d" % r, ""))
        if ref not in old and ref not in FLAGS_TEXT: continue
        fr = old.get(ref)
        section = M.get("C%d" % r, ""); cls = M.get("F%d" % r, "")
        flag = m_flag or (F.get("D%d" % fr) if fr else "verified") or "verified"
        if fr and not m_flag: flag = F.get("D%d" % fr, "verified")
        orig_row = ref.split("-")[1].lstrip("0") if ref.startswith("SI-") else "—"
        text = FLAGS_TEXT.get(ref) or (F.get("G%d" % fr) if fr else "")
        rows.append([ref, section, cls, flag, orig_row, BASIS.get(ref, M.get("Q%d" % r, "")), text])
    w.replace_sheet("Flags", table_sheet("FLAGS — the note behind every flagged or checked row in Matrix 2024, keyed by Ref. Evidence IDs point to the Evidence tab.",
        ["Ref", "Section", "Classification", "Flag", "2018 row", "Credential basis", "What it is", "Evidence"],
        rows, [8, 13, 30, 14, 8, 30, 96, 16], autofilter=True))

    w.replace_sheet("Decisions", table_sheet("DECISIONS — items only DPP can settle. Columns F to H are for DPP.",
        ["#", "Decision", "Why it matters", "Recommendation / first free option", "Rows affected", "DPP answer", "Initials", "Date", "Evidence"],
        [[str(n), d, why, rec, rows_, "", "", "", ""] for n, d, why, rec, rows_ in DECISIONS], [5, 34, 70, 58, 26, 24, 10, 12, 16]))

    w.replace_sheet("Credential check", table_sheet(
        "CREDENTIAL CHECK — the certification columns are DPP policy of record. This tab asks only (1) does each credential exist under that name with its issuer, and (2) does a code or referenced standard name a different credential for the work. SCOPE rows are the second kind.",
        ["ID", "Issuer", "Matrix rows", "What the matrix says", "What the issuer lists now / what the standard says", "Status", "Note", "Evidence"],
        [list(r) + [""] for r in CRED], [7, 9, 22, 34, 66, 14, 44, 16], autofilter=True))

    w.replace_sheet("Change Log", table_sheet("CHANGE LOG — what changed between the 2018 sheet (tab \"SI Qual 2018 (as received)\", preserved untouched) and \"Matrix 2024\".",
        ["ID", "Area", "Change", "Detail", "Evidence"], [list(r) + [""] for r in CHANGELOG], [7, 12, 52, 100, 16]))

    dd = {"C1": "Matrix rows (Ref)"}
    for r, v in DROPDOWN_MAP.items(): dd["C%d" % r] = v
    dd["A33"] = DROPDOWN_NOTE
    w.set_cells("Drop-down", dd)
    w.remove_sheet("QA log")
    w.save(path_out)


if __name__ == "__main__":
    apply(sys.argv[1], sys.argv[2])
