# Independent verification brief — Honolulu DPP special inspector qualification matrix

You are checking someone else's work. Your job is **not** to agree with it. Assume the
claims below are wrong until a source says otherwise, and report every disagreement,
including ones that seem minor.

Do not rebuild the spreadsheet. Do not improve anything. Produce a list of verdicts.

---

## Context you need

A Honolulu DPP (Department of Planning and Permitting) spreadsheet lists, per class of
special inspection, who may perform it and what certification they must hold. It was
built on the 2018 IBC. It was updated to the 2024 IBC. You are checking that update.

Citations in that document come from **four different authorities that look identical as
bare numbers**, and conflating them is the single largest source of error here:

| Layer | Document |
|---|---|
| Base IBC | 2024 International Building Code, Chapter 17 |
| State | Hawaii State Building Code (SBCC, adopted 20 Apr 2021 — adopts the **2018** IBC) |
| City | ROH 16-1.1 (building), 16-1.2 (residential), 19-1 (plumbing) |
| Department | DPP policy — no upstream document exists |

**Two namespaces both have a "Chapter 17."** IBC Chapter 17 is Special Inspections.
**ROH Chapter 17 is Honolulu's Electrical Code** and is irrelevant. The City's amendments
to IBC Chapter 17 live inside **ROH Chapter 16**. A keyword search for "Chapter 17" against
Honolulu ordinances returns the electrical code and looks like a negative result. It is not.

## Known failure mode in the work you are checking

The prior work made the same error three times: **fetching a web page, not finding a thing,
and reporting it as absent.** Twice this produced a false finding (a truncated ordinance read
as "the City does not amend Section 1705"; a program manual read as "this NICET program is
retired"). Once it produced a false defect report (footnote markers misread as missing
footnotes when they were column references).

So: for every negative claim below, establish whether the source was **complete**. If a
retrieval was truncated, paywalled, or partial, the correct verdict is **UNRESOLVED**, not
confirmation.

---

## Claims to check

Mark each: **CONFIRMED** / **CONTRADICTED** / **UNRESOLVED**, with the source you used.

### A. 2024 IBC Chapter 17 numbering — claimed VERIFIED from full text
1. 1705.10 is "Structural Integrity of Deep Foundation Elements".
2. 1705.11 is "Fabricated Items" (it was 1705.10 in the 2018 edition).
3. 1705.12 wind, 1705.13 seismic, 1705.14 testing for seismic, 1705.15 SFRM,
   1705.16 Intumescent, 1705.17 EIFS, 1705.18 penetrations and joints,
   1705.19 Testing for Smoke Control, 1705.20 Sealing of Mass Timber.
4. **1705.2.2 "Structural Stainless Steel" is new in 2024**, shifting cold-formed steel from
   1705.2.2/.3/.4 to 1705.2.3/.4/.5.
5. **1705.2.6 "Metal Building Systems" and Table 1705.2.6 are new in 2024.**
6. The 2024 title of 1705.16 **drops the words "Mastic and"** from the 2018 title.
7. The insertion causing the +1 shift is a **2021** change, not a 2024 one.

### B. City & County amendments — claimed VERIFIED from ROH 16-1.1
8. Amendment (112) adds **1705.19 "Fire-protection systems"**.
9. Amendment (113) adds **1705.20 "Termite protection"** and **1705.21 "Soils and foundation"**.
10. Amendment (111) replaces 1704.2.1 and permits the design professional of record and
    their personnel to act as special inspector **"with the exception of welding and high
    strength bolting."** ← check this wording exactly; a lot rests on it.
11. ROH 16-1.1 adds **Section 919** "Fire Protection Systems Special Inspections" to IBC
    Chapter 9, with 919.5 covering smoke control, and 919.5 **cites IBC 1705.18** — the 2018
    number, which shifts on adoption.
12. Amendment (114) places grading, excavation and earthwork under **ROH Chapter 18A**, not
    Chapter 17.

### C. Collisions — the central finding
13. Under the 2024 IBC, **City 1705.19 collides with base 1705.19**, and **City 1705.20
    collides with base 1705.20**. City 1705.21 does not collide.
14. Therefore the first free number above the base code is 1705.22.

### D. Adoption status
15. ROH 16-1.1 (Ord. 24-15, eff. 9 Aug 2024) adopts the Hawaii State Building Code as adopted
    20 Apr 2021, **which adopts the 2018 IBC**. So Honolulu is on the 2018 IBC via a 2024
    ordinance, and a 2024-IBC update is forward-looking.
16. ROH 19-1 adopts the Hawai'i State Plumbing Code (SBCC, 19 May 2020), which adopts the
    **Uniform Plumbing Code, 28th Edition (2018), 6th Printing** — **not** the 2021 UPC that
    the original sheet cited.
17. ROH 19-1 §105.2.7 requires plumbing special inspections be performed by **Registered
    Design Professionals licensed in Hawaii** not involved in the work.

### E. Referenced standards
18. **AWS D1.1:2020 §8.1.4.2** qualifies the Inspector as CWI or SCWI; **§8.1.4.5** qualifies
    **CAWI as an *Assistant* Inspector** working under supervision. So CAWI is a tier, not an
    alternative.
19. §8.1.4.2 lists five acceptable bases, including **CSA W178.2** and **ASNT SNT-TC-1A-VT
    Level II** — meaning the DPP sheet, which names only AWS CWI, is narrower than the standard.
20. **ACI 318-19 §26.13.1.4** requires reinforcement welding inspection per **AWS D1.4**, not D1.1.
21. **ACI 318-19 §26.13.1.5 / .1.6** require a certified inspector for post-installed and
    adhesive anchors, naming **ACI CPP 681.2-19** and **CPP 681.1-17**.
22. **IBC 1705.19.2** requires smoke control testing agencies to have "expertise in fire
    protection engineering, mechanical engineering and certification as air balancers."

### F. Certification bodies — the weakest set, check hardest
23. **ICC has no exam 49.** Reinforced Concrete is 47.
24. ICC designates the building inspector exams **B1** and **B2**; the sheet writes 1B and 2B.
25. ICC exams are now modular — **47C/47P, 84C/84P, 92C/92P, ECC/ECP, S1C/S1P, S2C/S2P**,
    plus a General Requirements module. Exam **86** is unchanged.
26. **ICC exam 93 "Tall Mass Timber Buildings Special Inspector" exists.**
27. **NICET's "Geotechnical Engineering Technology" program is NOT on its current programs
    page**, but NICET still hosts a 1994 manual for it with no retirement notice. Prior work
    marked this UNRESOLVED after first wrongly calling it retired. **Do not repeat that
    error — establish the actual status, or confirm it cannot be established online.**
28. **PCI has no credential named "Inspector Level II"**; its plant QC certifications are
    Level I / II / III.
29. **NFCA's CAP is a *Contractor* Accreditation Program**, and **FM 4991 / the UL firestop
    exam are also contractor accreditation** — yet three rows use them as *inspector*
    qualifications.
30. **AABC's individual credential is "Test and Balance Technician."**

### G. Not checked by the prior work — verify or confirm unavailable
31. IAPMO exam codes **(PI)** and **(PPE)**.
32. ICC code **(CF)** Commercial Fire Sprinkler.
33. **AWC "Durable Construction with Preservative-Treated Wood"** — is it a certification or
    a course?
34. **NFPA 99 certification** — does a credential by that name exist?
35. **HDOA** Certification of Pesticide Applicators, and courses **#4322, 4415, 4422, 4484**.
    The HDOA portal returned "Forbidden" on the prior attempt.

---

## What sources were and were not available

**Read in full:** 2024 IBC Chapter 17; ROH 16-1.1, 16-1.2, 19-1; ACI 318M-19; AWS D1.1:2020;
AWS QC1:2016-AMD1; the adopted Hawaii State Building Code and the codified state amendments;
a 2018-IBC-based full Chapter 17; ICC's published certification catalog.

**Not available:** the 2018 and 2024 IBC behind ICC's paywall were only reached for 2024;
AWS D1.4 and D1.6 in full; TMS 402/602; AISC 360 Ch. N and AISC 370; the HDOA course portal.

If you reach a claim whose source you cannot obtain, say so. An honest UNRESOLVED is more
useful than a confirmation built on a secondary source.

## Output

A table: claim number, verdict, source used, and — where you disagree — what the source
actually says. Then, separately: anything the prior work appears to have **missed**, and any
claim that is stated more confidently than its evidence supports.
