# Verifier brief

You are checking someone else's work. **Do not edit the workbook.** Produce a report of findings
in `handoff/`, nothing else.

## The deliverable and who it is for

`Special Inspector Qualification 2024 IBC.xlsx` updates the Honolulu Department of Planning and
Permitting (DPP) special inspector qualification matrix from the 2018 to the 2024 IBC. A DPP
engineer will use it to decide what to adopt. Start with the workbook's README tab.

## What "done" means (the owner's acceptance criteria)

1. Every section number and title is right for the 2024 IBC, with Hawaii State and City & County of
   Honolulu amendments accounted for.
2. DPP's 2018 data, the SI FORM and its dropdowns were not broken, and the file opens in Excel.
3. Anything called an error, conflict, collision or gap has a source that proves it.
4. Anything nobody could confirm is labelled for DPP to decide, not presented as settled.

## How the document is layered

| Layer | Document in `verification/sources` |
|---|---|
| Base code | 2024 IBC Chapter 17 (`licensed/ibc2024_ch17.txt`), editions in Chapter 35 |
| State | Hawaii State Building Code adoption document (`public/hi_sbc_2018_adopted.txt`) |
| City | ROH 16-1.1 building, 16-1.2 residential, 19-1 plumbing (`public/`) |
| Referenced standards | ACI 318-19 (code and commentary split), AWS D1.1 / D1.4 / D1.6, AISC 360-22, UPC 2018 |
| Issuers | dated snapshots in `snapshots/`, and IAS AC291 |
| DPP policy | no upstream document exists |

Two namespaces both have a "Chapter 17": IBC Chapter 17 is special inspections; ROH Chapter 17 is
Honolulu's electrical code and is irrelevant.

## Run the checks first

```
cd verification
python verify_si.py --new "../Special Inspector Qualification 2024 IBC.xlsx" --orig "../reference/Special Inspection Qualifications 082625 (original from DPP).xlsx" --excel
```

They should pass. **Passing is not your verdict.** The checks prove that quotes exist and that the
workbook is internally consistent. They do not prove that a quote means what the claim says, or
that the right thing was checked.

## What to attack

- **Meaning, not presence.** For a sample of at least 30 Evidence rows, read the quote in context in
  the source file and decide whether it actually supports the claim and the cell that uses it.
  Pay special attention to anything typed `absent` and to every credential basis that says
  `standard` or `code`.
- **What was not checked.** Look for claims in the workbook that no evidence covers because the
  coverage rule does not require them (for example, Matrix section titles for City rows, the SI
  FORM dropdown tab, the README tab). Look for 2024 IBC 1705 sections that have no row and are not
  listed as gaps.
- **Source fidelity.** Pick five source text files and compare passages against the originals if you
  have them (the originals' fingerprints are in `sources/manifest.json`). The ACI code/commentary
  split is done by page gutter; test it where a claim depends on it.
- **The checks themselves.** Read `verify_si.py`. For any check, describe a change to the workbook
  or evidence that would be wrong but would still pass.
- **Negative findings.** Any claim that something does not exist, is not required, or is not held:
  establish whether the search was complete. An honest UNRESOLVED beats a confident wrong answer.

## Output

`handoff/VERIFIER-REPORT-<date>.md`: a table of findings (what, where, severity, the source passage),
then anything stated more confidently than its evidence supports, then any check that can be
passed while wrong.
