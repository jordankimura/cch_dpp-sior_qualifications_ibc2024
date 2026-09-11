# Sources — full enumeration

Three tiers. Only tier 1 can be tested offline, which is why the harness's `T` checks
cover what they cover and no more.

## Tier 1 — BUNDLED, and tested by the harness

| File | Document | Used by | Status |
|---|---|---|---|
| `ibc2024_ch17_index.txt` | 2024 IBC Ch.17 — section numbers and headings only | T-01, T-04, T-05 | ICC copyright; index only, no provisions |
| `roh_16_1_1.txt` | ROH 16-1.1, building code as adopted by CCH | T-02, T-05 | Public ordinance |
| `roh_19_1.txt` | ROH 19-1, plumbing code | T-03 | Public ordinance |
| `aci318_19_sec26_13.txt` | ACI 318-19 (SI), Sec. 26.13 Inspection | T-06, T-07 | ACI copyright; extract only |
| `aws_d1_1_sec8_1_4.txt` | AWS D1.1/D1.1M:2020, Sec. 8.1.4 | T-08 | AWS copyright; extract only |

## Tier 2 — CONSULTED, NOT BUNDLED

Live issuer pages. They cannot be frozen into a test, but each is re-checkable by hand.
All consulted 2026-09-11.

| Source | What was taken from it |
|---|---|
| ICC, *Current Certification Exams — National Exams* catalog | 47C/47P, 84C/84P, 92C/92P, S1C/S1P, S2C/S2P, ECC/ECP, 86M, 93, B1, B2, CF, 66 — and that **49, 1B, 2B and EC do not appear** |
| ICC special inspector exam listing | The modular structure (General Requirements + Codes + Plans) |
| NICET certification-programs and civil-engineering pages | Current civil programs are CMT Asphalt/Concrete/Soils and Highway Construction Inspection; "Geotechnical Engineering Technology" is not listed |
| NICET, hosted Geotechnical Engineering Technology manual (1994) | The program existed, Levels I–IV, with no withdrawal notice — the contradiction behind U-01 |
| PCI personnel certification overview | Plant QC is Level I/II/III; field is CFA/CCA; no credential named "Inspector Level II" |
| AABC certification pages | The individual credential is "Test and Balance Technician" |
| IAPMO initial-certification pages | "UPC Residential and Commercial Plumbing Inspector (PI)" and "…Plans Examiner (PPE)" — exact matches |
| NFCA CAP program page | CAP is a **Contractor** Accreditation Program |
| FCIA pages for FM 4991 and the UL firestop exam | Both are **contractor** accreditation routes |
| AWCI EIFS inspector pages and registry | An EIFS inspector certification with a registry exists |
| ACI Shotcrete Inspector program page, policy CPP 6611-24 | Current program, introduced 2020 |
| ASSE 6000-series pages | NFPA 99 is a code; the credentials are ASSE 6020 and 6030 |
| AWS QC1:2016-AMD1 | CAWI may inspect only under active supervision of a CWI/SCWI |
| Hawaii DAGS adopted State Building Code; HAR Title 3 amendment index | State amendments to Ch.17 limited to 1704.2, 1704.2.1, 1704.2.3, 1704.2.4, 1704.3, 1705.3, 1705.11, 1705.11.1 |
| A 2018-IBC-based full Chapter 17; a 2021-IBC-based inspection manual | The baseline numbering, and that the +1 shift originates in 2021 |
| San Antonio IB 132; ACEC/SEAOG Georgia guidelines; Montgomery County MD manual | Early corroboration of 2024 numbering, later superseded by the code text itself |

## Tier 3 — NOT CONSULTED

Declared so the gap is visible. Check A-09 fails if this list disappears.

- AISC 360 Chapter N and AISC 370 — the structural steel rows and the new stainless section
- TMS 402/602 — the masonry rows
- AWS D1.4 and D1.6 in full — named by ACI 318-19 and by 1705.2.2 respectively
- AWC — no credential found under the name the matrix uses
- HDOA — pesticide applicator certification and courses 4322 / 4415 / 4422 / 4484; portal returns Forbidden
- The 2021 and 2024 Uniform Plumbing Codes — only the adopting ordinance was read
- What DPP's registry accepts in practice

## Refreshing for the next code cycle

Replace the tier 1 files with the new edition's equivalents and re-run. T-01 and T-04
will report exactly which citations and titles no longer match. Tier 2 has to be
re-walked by hand — that is the cost of a credential requirement whose issuer can
renumber it without telling anyone.
