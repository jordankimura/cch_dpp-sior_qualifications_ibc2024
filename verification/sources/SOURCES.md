# Sources — full enumeration

<!-- GENERATED FROM manifest.json BY gen_sources_md.py — DO NOT HAND-EDIT. -->

Compiled 2026-09-11 for `Special Inspector Qualification 2024 IBC rev18.xlsx`.

**32 sources: 5 bundled and machine-tested, 20 consulted but not frozen, 7 declared unobtainable.**

Every claim in the workbook traces to an entry below. The manifest is the master copy; this page is a rendering of it.

## Tier 1 — bundled in this repo, tested by the harness

A byte-identical copy is committed here and hashed. Check M-01 re-hashes each one, so a silently swapped or edited source fails the build.

| ID | Document | File | Completeness | Used by |
|---|---|---|---|---|
| `ibc-2024-ch17-index` | International Building Code, 2024 Edition, Chapter 17 - Special Inspections and Tests | `licensed/ibc2024_ch17_index.txt` | partial | T-01, T-04, T-05 |
| `roh-16-1-1` | Revised Ordinances of Honolulu, Sec. 16-1.1 - Building Code (Ord. 24-15, eff. 9 Aug 2024) | `public/roh_16_1_1.txt` | full | T-02, T-05 |
| `roh-19-1` | Revised Ordinances of Honolulu, Sec. 19-1 - Plumbing Code | `public/roh_19_1.txt` | full | T-03 |
| `aci-318-19-sec26-13` | ACI 318M-19, Building Code Requirements for Structural Concrete, Sec. 26.13 - Inspection | `licensed/aci318_19_sec26_13.txt` | partial | T-06, T-07 |
| `aws-d1-1-2020-sec8-1-4` | AWS D1.1/D1.1M:2020, Structural Welding Code - Steel, Sec. 8.1.4 - Inspector Qualification | `licensed/aws_d1_1_sec8_1_4.txt` | partial | T-08 |

**Limits on the partial sources.** A claim may not be stated more strongly than the source under it. The harness marks any check resting on a partial source with `*`.

- `ibc-2024-ch17-index` — Section numbers and headings only. This file carries no provision text, so it establishes what a section is CALLED and that it EXISTS - never what it requires. Any claim about the substance of a 2024 IBC provision rests on the full text read during the session, which is not frozen here.
- `aci-318-19-sec26-13` — Sec. 26.13 only. Supports claims about inspection and inspector qualification under ACI 318-19 and nothing else in that standard. It does not contain AWS D1.4 or CPP 681.x themselves - it only establishes that ACI 318 NAMES them.
- `aws-d1-1-2020-sec8-1-4` — Sec. 8.1.4 only. Establishes the Inspector / Assistant Inspector tier structure and the acceptable qualification bases. It does not establish anything about welding procedures, acceptance criteria, or any other clause of D1.1.

## Tier 2 — read, but not frozen

Live issuer pages and documents that cannot be redistributed. Each is re-checkable by hand; none can be frozen into a test. All consulted 2026-09-11.

| ID | Source | Issuer token | What it established |
|---|---|---|---|
| `icc-exam-catalog` | ICC, Current Certification Exams - National Exams catalog | `ICC` | 47C/47P, 84C/84P, 92C/92P, S1C/S1P, S2C/S2P, ECC/ECP, 86M, 93, B1, B2, CF, 66 are current; 49, 1B, 2B and EC do not appear. |
| `icc-si-exam-listing` | ICC special inspector exam listing | `ICC` | The modular structure: General Requirements module plus Codes and Plans modules. |
| `nicet-programs` | NICET certification-programs and civil-engineering-technology program pages | `NICET` | Current civil programs are CMT Asphalt / Concrete / Soils and Highway Construction Inspection. 'Geotechnical Engineering Technology' is not among them. |
| `nicet-geotech-manual-1994` | NICET, Geotechnical Engineering Technology program manual (1994), still hosted by the issuer | `NICET` | The program existed, Levels I-IV. This is the contradiction behind U-01. |
| `pci-personnel-cert` | PCI personnel certification program overview | `PCI` | Plant QC certifications are Level I / II / III; field programs are CFA and CCA. No credential named 'Inspector Level II' exists. |
| `aabc-certification` | AABC certification application and program pages | `AABC` | AABC's individual credential is 'Test and Balance Technician'. |
| `nebb-certification` | NEBB certification pages | `NEBB` | NEBB technician certifications exist. |
| `iapmo-certification` | IAPMO initial-certification pages | `IAPMO` | 'UPC Residential and Commercial Plumbing Inspector (PI)' and '...Plans Examiner (PPE)' are exact current matches, codes included. |
| `nfca-cap` | NFCA CAP program page | `NFCA` | CAP is a Contractor Accreditation Program, not an individual inspector credential. |
| `fcia-fm4991-ul` | FCIA pages for FM 4991 and the UL firestop examination | `UL / FM` | Both are contractor accreditation routes. FM 4991 is the Standard for the Approval of Firestop Contractors. |
| `awci-eifs` | AWCI EIFS inspector pages and certification registry | `AWCI` | An EIFS inspector certification with a published registry exists. |
| `aci-shotcrete-program` | ACI Shotcrete Inspector program page and certification policy CPP 6611-24 | `ACI` | Current program, introduced 2020. Full name is 'Shotcrete Inspector'; an Associate level also exists. |
| `asse-6000-series` | ASSE 6000-series medical gas personnel certification pages | `NFPA` | NFPA 99 is the Health Care Facilities Code, not a personnel credential. The personnel certifications are ASSE 6020 and 6030. |
| `aws-qc1-2016` | AWS QC1:2016-AMD1, Standard for AWS Certification of Welding Inspectors | `AWS` | A CAWI may inspect only under the active supervision of a CWI or SCWI. Second, independent confirmation of the tier structure in D1.1 Sec. 8.1.4. |
| `hi-dags-state-building-code` | Hawaii DAGS, adopted Hawaii State Building Code | — | The State adopts the 2018 IBC. This is why Honolulu is on the 2018 IBC and a 2024 update is forward-looking. |
| `har-title-3` | Hawaii Administrative Rules, Title 3 - codified state amendment index | — | State amendments to IBC Chapter 17 are limited to 1704.2, 1704.2.1, 1704.2.3, 1704.2.4, 1704.3, 1705.3, 1705.11 and 1705.11.1. |
| `roh-16-1-2` | Revised Ordinances of Honolulu, Sec. 16-1.2 - Residential Code, amendment (11) | — | R109.5 through R109.9, including R109.5.2 termite protection on the residential side. |
| `ibc-2018-ch17` | A published 2018-IBC-based full Chapter 17 | — | The baseline numbering that the prior-citation column rests on. |
| `ibc-2021-inspection-manual` | A 2021-IBC-based special inspection manual (Montgomery County MD) | — | 1705.10 and 1705.11 already carry their 2024 numbering in the 2021 edition - so the +1 shift originates in 2021, not 2024. This corrected an earlier wrong attribution. |
| `peer-jurisdiction-guidance` | San Antonio IB 132; ACEC/SEAOG Georgia special inspection guidelines | — | Early corroboration of the 2024 numbering. Superseded by the code text itself and retained only as the second source required for layer-IBC claims. |

Partial among these:

- `nicet-geotech-manual-1994` — A hosted historical manual. It establishes that the program existed with Levels I-IV and that no withdrawal notice accompanies it. It does NOT establish current status either way - absence from a marketing page is not retirement, and a hosted manual is not an active program.
- `nebb-certification` — Confirms that NEBB technician certifications exist. The exact printed credential name that the matrix should carry was NOT pinned down to an issuer page, which is why the Credential check row for NEBB is marked 'naming' rather than 'current'.

## Tier 3 — not obtained

Declared so the gap is visible rather than implied by silence. Check M-03 fails if any entry here stops naming what it leaves open.

| ID | Document | Why not | Leaves open |
|---|---|---|---|
| `aisc-360-ch-n-and-370` | AISC 360 Chapter N (Quality Control and Quality Assurance) and AISC 370 | not obtained - paywalled | Structural steel rows (matrix rows 4, 5, 6, 8, 14, 18) unverified against the referenced standard; The new IBC-2024 1705.2.2 stainless steel row |
| `tms-402-602` | TMS 402/602, Building Code Requirements and Specification for Masonry Structures | not obtained - paywalled | Masonry rows unverified against the referenced standard |
| `aws-d1-4-and-d1-6` | AWS D1.4 (Reinforcing Steel) and AWS D1.6 (Stainless Steel) | not obtained - paywalled | Decision 8 - whether rebar welding rows need a D1.4-specific qualification |
| `awc-preservative-treated-wood` | AWC, 'Durable Construction with Preservative-Treated Wood' - credential or course | searched; no matching credential located | Credential check row 34 - whether the matrix names a course rather than a certification |
| `hdoa-pesticide-applicator` | Hawaii Department of Agriculture, Certification of Pesticide Applicators; courses 4322 / 4415 / 4422 / 4484 | portal returned Forbidden | U-04 |
| `upc-2021-and-2024` | Uniform Plumbing Code, 2021 and 2024 editions | not obtained - paywalled | U-02 |
| `dpp-registry-practice` | What DPP's special inspector registry accepts in practice | no external referent exists | U-03; U-05 |

## Issuer coverage

Every issuer named on the workbook's **Credential check** tab must resolve to at least one source above. Check M-02 enforces this.

- **AABC** → `aabc-certification`
- **ACI** → `aci-318-19-sec26-13`, `aci-shotcrete-program`
- **AWC** → `awc-preservative-treated-wood`
- **AWCI** → `awci-eifs`
- **AWS** → `aws-d1-1-2020-sec8-1-4`, `aws-qc1-2016`
- **HDOA** → `hdoa-pesticide-applicator`
- **IAPMO** → `iapmo-certification`
- **ICC** → `ibc-2024-ch17-index`, `icc-exam-catalog`, `icc-si-exam-listing`
- **NEBB** → `nebb-certification`
- **NFCA** → `nfca-cap`
- **NFPA** → `asse-6000-series`
- **NICET** → `nicet-programs`, `nicet-geotech-manual-1994`
- **PCI** → `pci-personnel-cert`
- **UL / FM** → `fcia-fm4991-ul`

## Refreshing for the next code cycle

Replace the tier 1 files with the new edition's equivalents, update their `sha256` and `edition` in `manifest.json`, re-run `gen_sources_md.py`, and re-run the harness. T-01 and T-04 report exactly which citations and titles no longer match. Tier 2 has to be re-walked by hand — that is the cost of a credential requirement whose issuer can renumber it without telling anyone.

