# Second pass: working notes

Notes carried over from the session that did the QA and the evidence rebuild (2026-09-11). They
come from the owner's review of the workbook plus that session's own findings. You are the
**builder** for the second pass. A separate, fresh session verifies afterwards using
`handoff/VERIFIER-BRIEF.md`, so leave the repository in a state that session can check.

The owner is new to git and the terminal. Explain what you are doing in plain language, and
commit and push only once the checks are green.

---

## Ground rules learned the hard way

- **Never save the workbook with openpyxl.** It silently deletes the SI FORM dropdowns. Edit
  through `verification/tools/xlsx.py`, as a numbered script in `verification/edits/` (next is
  `e03_...`), or have the owner edit in Excel.
- After any workbook edit: `python verification/tools/sync_workbook.py "<workbook>"`, then
  `python verification/verify_si.py ... --excel`, then `python verification/tools/mutation_tests.py`.
  Excel is on this machine. Excel COM will not open files under the user's Temp folder, so test copies
  go in a folder next to the repo.
- Quotes in `verification/evidence.json` are pulled from the source text with
  `evidence_lib.find(source_id, regex)`, never typed from memory.
- Every new check gets a mutation in `tools/mutation_tests.py` that proves it can fail.
- A "does not exist" claim needs a complete source (check E-02). A web page that did not show
  something is not proof it doesn't exist: that exact mistake produced the false "ICC has no exam 49".
- The certification columns L-P are DPP policy of record. Change them only under the rule in item 5.

## Answers already given to the owner (so you don't re-derive them)

- **ROH 1705.19 and IBC 1705.19 really are two different sections with one number.** ROH 16-1.1
  amend. (112) literally adds "1705.19 Fire-protection systems". It was written against the 2018 IBC,
  whose 1705 ends at 1705.18, so the number was free. The 2021 and 2024 IBC then added their own
  1705.19 (Testing for Smoke Control) and 1705.20 (Sealing of Mass Timber). That is the COLLISION;
  renumbering is DPP's call (Decision 1). The matrix shows both because both exist.
- **The IBC almost never names certifications.** 1704.2.1 leaves competence to the building
  official. Certification names arrive through referenced standards (AISC 360, ACI 318, AWS D1.x,
  TMS 402/602, SDI QA/QC). No 2018 certification cell has yet been shown to conflict with a 2024
  requirement. But that was only a check against the 2024 chain, never a 2018-chain vs 2024-chain
  comparison (item 7).

---

## Work items, in suggested order

### 1. Formatting consistency (owner noticed)
The owner can see which cells were edited afterwards:
- Matrix A4:A6 (SI-004/005/006 CONFLICT) are left-aligned with no fill. Existing CONFLICT cells are
  style 127 (centered, orange).
- A20 (SI-016, now "check") still carries the orange CONFLICT fill.
- Column R (Evidence) is style 125 (left, no fill) on every row, including the shaded City and
  plumbing rows (37-43), which use style 131.

Make the flag cell style a function of the flag value (CONFLICT/COLLISION 127, GAP 128, others 126 or
131 on shaded rows). Match column R to each row's shading. Add a check that fails when style and
value disagree.

Text overflows many boxes. Fixed row heights (ht=27) are the cause. Propose a scripted step that has
Excel auto-fit row heights (Excel COM: `Rows.AutoFit()` then save). The harness already passes after
an Excel re-save. The owner said this could be manual, but a scripted auto-fit is better if it works.

### 2. Dropdown: close the loop in both directions
D-02 checks that every matrix row is reachable from a dropdown entry. **Nothing checks the reverse**:
that every SI FORM pick-list entry (Drop-down B2:B29) maps to at least one matrix row. Column C
currently allows "—" (used for "Other"). Add the reverse check with an explicit exemption for
"Other" only, plus a mutation.

### 3. Credential check: lead with the credential
The owner wants the tab organised around the credential, since that's what is being counted. Suggested
columns: Credential (as printed on the matrix) | Current official name | Issuer | Matrix rows | What
the issuer / standard says | Status | Note | Evidence. Keep the CC-nn IDs stable so evidence keys
still resolve (or migrate them in evidence.json in the same edit).

### 4. Human-readable source names
Source IDs like `ibc-2024-ch17` show up on the Evidence and Sources tabs. Keep the IDs as internal
keys, but give each catalog entry a short display name ("2024 IBC, Chapter 17", "ACI 318-19 (code)",
"ROH 16-1.1") and show that name in the workbook. Put it in `CATALOG` in `tools/extract_sources.py`
so it lands in manifest.json, then use it in `sync_workbook.py`.

### 5. Certification changes from 2018 to 2024 (owner decision + an open design question)
**Owner decision:** where the 2024 code chain conflicts with a 2018 certification cell, change the
cell in columns L-P, mark it, and have a check that the change is grounded in evidence.

**Guardrail (proposed to and accepted by the owner):** change a cell only when the 2024 chain makes it
non-compliant (the code requires something the cell doesn't provide, or the cell names something that
no longer exists). Where DPP is stricter than the code, that's allowed: leave it and record it as a
Decision.

**Open design question: the owner asked for a flag for changed certifications, and whether naming
fixes overlap with requirement changes.** They can overlap on the same row, and a row can also be
CONFLICT or COLLISION, while the Flag column holds one value. Recommendation to put to the owner:
a **separate column** next to the certifications, "Changed from 2018", with values blank / `code` /
`name` / `code + name`. Confirm with the owner before building. Also ask whether pure naming fixes
(1B→B1, 2B→B2, PCI's real title) should be applied. The previous session recommended yes.

Mechanism for every changed cell:
- a Change Log row with Area `Credential`, quoting the exact 2018 value and the exact new value;
- evidence from a 2024-edition source, code text not commentary (for `name` changes, an issuer
  snapshot);
- a check that the "before" matches the 2018 sheet, the "after" matches the matrix, and the evidence
  type fits the change type. I-05 already refuses unregistered changes; extend it.

### 6. Gap rows: the owner questions why they are there
Rows NEW-01 to NEW-05 have no "who may perform" and no certifications. The owner's view: the matrix
is DPP's 2018 sheet adopting the 2024 standards, and a row with nothing in it has no purpose unless a
special inspector for that work must be credentialed through DPP. They also carry no evidence of
their own (only their Flags notes do), and A-05 only checks that no certification was invented.

For each gap row, establish with evidence:
- **Does the 2024 IBC actually require a special inspection** here?
  - **1705.10 probably does not.** It requires "an engineering assessment" only "whenever there is a
    reasonable doubt" about a deep foundation element. That's a trigger, not a routine special
    inspection. It may not belong in the matrix at all.
  - **1705.11 Fabricated Items** routes to 1704.2.5, and the five Fabricator rows already cover it.
    Likely a duplicate.
  - **1705.2.2 stainless steel** routes to AISC 370-21 (not held), with welding under AWS D1.6 (held).
    The closest existing row is Structural Steel.
  - **1705.5.3 Mass Timber and 1705.20 Sealing of Mass Timber:** ICC Category 93 Tall Mass Timber
    Buildings Special Inspector exists (evidence already held).
- If yes, **propose** who-may-perform and certifications, marked PROPOSED for DPP, each backed by
  evidence or by analogy to a named existing row. If no, move it off the matrix and into a Decision.
- Update Decision 5 to match. Add checks: every GAP or PROPOSED row has evidence on its flag and on
  every non-empty cell, and every row still on the matrix names the 2024 section that requires the
  inspection.

### 7. Second pass: what changed in the whole chain, 2018 → 2024
For every sentence about special inspection or inspector qualification, record added / removed /
changed between the editions the 2018 IBC references and the ones the 2024 IBC references. Confirm
editions from each IBC's Chapter 35 (the 2024 Ch. 35 is held; the 2018 one is not yet).

| Pair | Held? |
|---|---|
| IBC 2018 Ch. 17 (with Hawaii amendments) vs IBC 2024 Ch. 17 | both held |
| AISC 360-16 vs 360-22, Chapter N | both held |
| ACI 318-14 vs 318-19 | **318-14 not held** |
| AWS D1.1:2015 vs D1.1:2020 | **2015 not held** |
| AWS D1.4:2011 vs D1.4:2018 | **2011 not held** |
| TMS 402/602-16 vs -22 | **neither held** |
| SDI QA/QC 2011 vs 2022 | **neither held** |

**Lead already seen, not yet a finding:** AISC 360-22 N4.2 changed the fallback route for QA welding
inspectors from "AWS D1.1 clause 6.1.4" (360-16) to "AWS D1.1 clause 8.1.4.2(5)" only. It also adds
QA **coating** inspector qualifications (AMPP/NACE CIP Level 1 or SSPC PCI Level 1). Determine whether
IBC 1705.2.1's routing into AISC 360 QA makes coating inspection a required special inspection. If
so, the matrix has no row for it.

**Also check** the 2018 → 2024 change in 1704.2 / 1704.2.1 wording: 2021 onward speaks of "approved
agencies" documenting inspector competence, where the Hawaii 2018 text says each special inspector
does. Does that change who may perform?

### 8. Special inspection requirements outside IBC Chapter 17
Search for requirements that create or change a classification:
- Other 2024 IBC chapters: 909 smoke control, Chapter 18 foundations, Chapter 21 masonry, Chapter 22
  steel (2203 stainless), Chapter 23 wood. Only Ch. 17 and Ch. 35 are held, so the owner needs to
  supply the others.
- ROH 16-1.1: 110.3.10 (special inspections listed on plans), Section 919 fire protection special
  inspections (919.1 to 919.6), and any other amendment that says "special inspection".
- ROH 16-1.2 R109.5 residential special inspections (held): does each have a matrix row?
- ROH 19-1: 105.2.7, and **301.5.6, which requires plumbing special inspection for alternative
  engineered designs** (no matrix row).
- ROH Chapter 18A (grading) and the Honolulu fire code, if they exist. **Not held.**

### 9. Known soft spots in the harness (fix or leave explicitly for the verifier)
- Coverage (E-05) does not require evidence for City row section titles, SI FORM dropdown tab text,
  or the README tab.
- Evidence proves a quote exists, not that it means what the claim says. The verifier samples for this.
- The ACI code/commentary split is a page-gutter heuristic.
- `upc-2018` is a garbled scan; 1319.4 was not found in it.
- `absent ... within` depends on the heading regex in `evidence_lib.section_span`.

---

## Documents the owner was asked to find

ACI 318-14; AWS D1.1:2015; AWS D1.4:2011; TMS 402/602 (2016 and 2022); SDI QA/QC; AISC 370-21;
2018 IBC Chapter 35; other 2024 IBC chapters (9, 18, 21, 22, 23); ROH Chapter 18A; the Honolulu fire
code. Anything not found gets a `NOT_HELD` entry in `tools/extract_sources.py` that says what it
leaves open.

## When you finish

1. All checks green with `--excel`; all mutations caught, including new ones.
2. One commit per work item, messages saying what and why. Push.
3. Add Change Log rows for anything that changed in the workbook. Update the README tab text if tabs
   or columns changed (`README_ROWS` in `sync_workbook.py`).
4. Tell the owner it's ready for the verifier: a fresh session told to read `handoff/VERIFIER-BRIEF.md`.

Context for whoever reads this: the deliverable goes to a DPP Civil Engineering Branch reviewer who
received an earlier, different version on 2026-09-10. That version is superseded. The `rev18` copies
in the owner's Downloads folder do not open in Excel and must not be sent.
