# Special Inspector Qualification — 2024 IBC revision

![verify](https://github.com/jordankimura/cch_dpp-sior_qualifications_ibc2024/actions/workflows/verify.yml/badge.svg)

A draft revision of the Honolulu DPP Special Inspection Qualifications matrix, moved from the
2018 to the 2024 International Building Code, with the Hawaii State and City & County amendments
traced.

**Status: draft, not adopted.** Honolulu adopts the Hawaii State Building Code of 20 April 2021,
which adopts the **2018** IBC (ROH 16-1.1). This revision is forward-looking.

## What's here

```
Special Inspector Qualification 2024 IBC.xlsx   the deliverable (open the README tab first)
reference/
    Special Inspection Qualifications 082625 (original from DPP).xlsx   DPP's 2018 sheet, the baseline
verification/
    verify_si.py        the checks - run this
    evidence.json       what backs every claim in the workbook (generates the Evidence tab)
    workbook.txt        plain-text dump of every tab, so git diffs show cell-level changes
    sources/            full text of every document relied on, plus manifest.json (the catalog)
    edits/              every scripted change to the workbook, in order, with the reason
    tools/              extraction, safe workbook editing, and sync
handoff/
    QA-REPORT-2026-09-11.md    independent QA of rev18 and what was fixed
    VERIFIER-BRIEF.md          instructions for the next independent check
```

**Before making this repository public, read `LICENSING.md`.**

## Running the checks

```
cd verification
python verify_si.py --new "../Special Inspector Qualification 2024 IBC.xlsx" --orig "../reference/Special Inspection Qualifications 082625 (original from DPP).xlsx"
```

Python 3 only, no packages. Add `--excel` on a Windows machine with Excel to also have Excel open
the file and read the SI FORM dropdown back. GitHub runs the checks on every push; the badge above
is the result for the latest commit.

What the checks establish, in plain terms:

- DPP's 2018 data, SI FORM and dropdowns are untouched where they must be.
- The file obeys the rules Excel enforces when opening (rev18 did not; see Change Log CL-17).
- Every "who may perform" cell is a faithful translation of the 2018 notation.
- Every IBC and City citation is a real heading in the full code text, titles exact.
- Collisions are computed from the code text, not asserted.
- Every flag, credential basis, Flags note, Decision, Credential check row and citation change has
  evidence, and every quote is found in the full text of its source.
- A "does not exist" claim is only accepted against a complete source.
- Every source file matches its fingerprint and is used.

Each check was also run against deliberately broken copies (a wrong quote, a missing flag, rows out
of order, an over-long tab name, a false "does not exist" claim, and 13 others) to confirm it fails.
The workbook also passes after being opened and re-saved by Excel.

What no check can establish is listed at the end of every run under **OPEN**.

## Changing the workbook

Opening and re-saving the workbook with common Python libraries **deletes the SI FORM dropdowns**.
Either:

- **edit in Excel**, save, then run `python verification/tools/sync_workbook.py "<workbook>"` and the
  checks; or
- **script the change** in `verification/edits/` using `tools/xlsx.py`, which edits only what it
  touches.

Then commit with a message that says what changed and why. When DPP answers the Decisions tab, one
commit per decision (`Decision 3: welding exception applies to shop fabrication`).

To add or change a source: add it to `CATALOG` in `tools/extract_sources.py`, re-run it with the
originals, run `gen_sources_md.py`, then cite it from `evidence.json`.
