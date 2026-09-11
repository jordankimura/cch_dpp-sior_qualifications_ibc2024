# Special Inspector Qualification — 2024 IBC revision

![verify](https://github.com/OWNER/REPO/actions/workflows/verify.yml/badge.svg)

> Replace `OWNER/REPO` in the badge URL above once the repository is created.

Draft revision of the DPP Special Inspection Qualifications matrix, moved from the
2018 to the 2024 International Building Code, as amended by the Hawaii State Building
Code and ROH 16-1.1.

**Status: draft, not adopted.** Honolulu currently adopts the Hawaii State Building
Code of 20 April 2021, which adopts the **2018** IBC (ROH 16-1.1, Ord. 24-15, effective
9 August 2024). This revision is forward-looking.

---

## What's here

```
Special Inspector Qualification 2024 IBC rev18.xlsx   the deliverable
verification/
    verify_si.py                the check harness — run this
    verification-report.txt     its output as of this bundle
    sources/
        public/                 Revised Ordinances of Honolulu — public law
        licensed/               extracts of copyrighted standards — SEE LICENSING.md
handoff/
    VERIFICATION-BRIEF.md       brief for an independent reviewer
reference/
    Special Inspection Qualifications 082625 (original from DPP).xlsx
    si-lookup.html              browser lookup tool, open in any browser
```

**Before making this repository public, read `LICENSING.md`.**

## The workbook

Thirteen tabs. The four that matter first:

| Tab | What it is |
|---|---|
| **Matrix 2024** | The deliverable. One row per classification, filterable. |
| **Flags** | The note behind every flagged row, keyed by Ref. |
| **Decisions** | 15 items only DPP can settle, with answer / initials / date columns. |
| **Credential check** | Every certification checked against its issuer. |

Then: `SI FORM dropdown` (what changed in the pick list), the five original tabs
including `SI Qualification (2018 as received)` untouched, `Change Log`, `Sources`,
and `QA log`.

**Certification columns are reproduced from the 2018 sheet and have not been edited.**
They are DPP policy of record; findings about them are recorded on Decisions and
Credential check, not applied.

## Running the verification

```
cd verification
python3 verify_si.py \
  --new  "../Special Inspector Qualification 2024 IBC rev18.xlsx" \
  --orig "../reference/Special Inspection Qualifications 082625 (original from DPP).xlsx"
```

Requires `python3` and `openpyxl` (`pip install openpyxl`). Sources are found
automatically in `verification/sources`. Exit code 0 if every check passes, 1 if not.

23 deterministic checks. **No language model is involved in any assertion** — each one
compares files or greps a source document. Checks are tagged by method:

- **I — Inspection.** File comparison. Proves what was *not* touched.
- **A — Analysis.** Internal consistency, recomputed independently.
- **T — Test.** Checked against a source document on disk.

The report ends with five items marked **NOT MACHINE-VERIFIABLE**, each with its reason.
Those are not failures; they are questions no program can settle.

## What still needs DPP

1. The 15 items on the Decisions tab, chiefly: renumbering the two City sections that
   collide with base-code numbers under the 2024 edition, and the welding /
   high-strength-bolting exception in ROH 16-1.1 amendment (111).
2. The five NOT MACHINE-VERIFIABLE items, chiefly the status of NICET's
   "Geotechnical Engineering Technology" program, which five rows depend on.
3. Confirmation of the attribution sentence at the foot of Matrix 2024.

## A caution

`reference/si-lookup.html` is generated from the same data as the workbook but is
**not linked to it**. If the matrix changes, the page does not. Treat the workbook as
the record and the page as a reading aid.
