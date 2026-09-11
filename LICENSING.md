# Licensing note — read before making this repository public

`verification/sources/licensed/` holds the **full text** of copyrighted codes and standards, so the
checks can confirm every quote against the whole document rather than an extract:

| File | Rights holder |
|---|---|
| `ibc2024_ch17.txt`, `ibc2024_ch35.txt`, `ibc2021_ch17.txt`, `ibc2018_ch17_hawaii.txt` | International Code Council |
| `aci318_19_code.txt`, `aci318_19_commentary.txt` | American Concrete Institute |
| `aws_d1_1_2020.txt`, `aws_d1_4_2018.txt`, `aws_d1_6_2017.txt` | American Welding Society |
| `aisc360_22.txt`, `aisc360_16.txt` | American Institute of Steel Construction |
| `upc_2018.txt` | IAPMO |

`verification/sources/snapshots/` holds dated text copies of issuer web pages and IAS AC291; those
publishers also hold copyright. `verification/sources/public/` holds the Revised Ordinances of
Honolulu and the Hawaii State Building Code adoption document, which are public law.

## Keep this repository private

That is the simple answer. A private repository shared only with the people doing the work raises
no question.

## If it must ever be made public

Do not flip the existing repository to public: the licensed text is in its history. Start a fresh
repository from a copy with `verification/sources/licensed/` and `verification/sources/snapshots/`
removed. The checks that read those files will then fail rather than pass silently, which is the
intended behaviour; anyone holding licensed copies can restore the folders locally with
`tools/extract_sources.py`.
