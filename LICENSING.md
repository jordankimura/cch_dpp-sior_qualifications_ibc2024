# Licensing note — read before making this repository public

`verification/sources/licensed/` contains extracts of copyrighted standards. They are
here so the verification harness can check citations offline. They are inputs to a
test, not reference copies.

| File | Rights holder | What it contains |
|---|---|---|
| `ibc2024_ch17_index.txt` | International Code Council | Section numbers and headings only — no provisions |
| `aci318_19_sec26_13.txt` | American Concrete Institute | Extract of Sec. 26.13, Inspection |
| `aws_d1_1_sec8_1_4.txt` | American Welding Society | Extract of Sec. 8.1.4, Qualification of Inspection Personnel |

`verification/sources/public/` contains the Revised Ordinances of Honolulu, which are
public law and carry no such restriction.

## Keep this repository private

That is the simple answer, and the recommended one. A private repo under the
department's account raises no question at all.

## If it must be made public

1. Uncomment the `verification/sources/licensed/` line in `.gitignore`
2. `git rm --cached -r verification/sources/licensed`
3. Commit

The harness degrades gracefully: checks T-01, T-04, T-05, T-06, T-07 and T-08 will
report **SKIP — source not supplied** rather than failing, and anyone with licensed
access can restore the folder locally to run the full suite.

Note that removing the files from the working tree does **not** remove them from git
history. If they have already been pushed to a public repository, rewriting history or
starting a fresh repository is the only real remedy.
