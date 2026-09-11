#!/usr/bin/env python3
"""
Proves the checks can fail. Copies the repository to a temp folder, breaks one thing at a
time, runs verify_si.py, and reports which checks caught it. A break caught by NOTHING is a
hole in the harness.

    python3 tools/mutation_tests.py

Add a mutation here whenever you add a check. Exit code 1 if any mutation goes uncaught.
"""
import json, os, re, shutil, subprocess, sys, tempfile

VER = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REPO = os.path.dirname(VER)
WB = "Special Inspector Qualification 2024 IBC.xlsx"
ORIG = "../reference/Special Inspection Qualifications 082625 (original from DPP).xlsx"


def fresh():
    base = tempfile.mkdtemp(prefix="si-mut-")
    dst = os.path.join(base, "repo")
    shutil.copytree(REPO, dst, ignore=shutil.ignore_patterns(".git", "__pycache__"))
    return base, dst


def run(dst):
    r = subprocess.run([sys.executable, "verify_si.py", "--new", "../" + WB, "--orig", ORIG],
                       cwd=os.path.join(dst, "verification"), capture_output=True, text=True, encoding="utf8")
    return r.returncode, [l.split()[0] for l in r.stdout.splitlines() if re.match(r"^[A-Z]-\d\d\s+FAIL", l)]


def ev_edit(dst, fn):
    p = os.path.join(dst, "verification", "evidence.json"); d = json.load(open(p, encoding="utf8"))
    fn(d["evidence"]); json.dump(d, open(p, "w", encoding="utf8"), indent=1, ensure_ascii=False)


def wb_edit(dst, fn):
    sys.path.insert(0, os.path.join(dst, "verification", "tools"))
    import importlib, xlsx; importlib.reload(xlsx)
    w = xlsx.Workbook(os.path.join(dst, WB)); fn(w); w.save(os.path.join(dst, WB))
    sys.path.pop(0)


TESTS = [
 ("rev18 mistake: claim ICC has no Category 49", lambda d: ev_edit(d, lambda E: E.append(dict(id="E-999", claim="x", type="absent", source="icc-si-exams", where="p", part="text", terms=["Category 49"], used_by=["Credential check:CC-02"])))),
 ("rev18 mistake: anchor programs cited as code", lambda d: ev_edit(d, lambda E: [e.update(part="code") for e in E if e["source"] == "aci-318-19-commentary"][:1])),
 ("quote altered by one word", lambda d: ev_edit(d, lambda E: [e.update(quote=e["quote"].replace("certified inspector", "qualified inspector")) for e in E if "adhesive anchors shall" in e.get("quote", "")])),
 ("absence claimed against a partial source", lambda d: ev_edit(d, lambda E: E.append(dict(id="E-998", claim="x", type="absent", source="upc-2018", where="p", part="text", terms=["zzzz"], used_by=["Flags:SI-035"])))),
 ("remove all evidence for one Flags note", lambda d: ev_edit(d, lambda E: [e.update(used_by=[u for u in e["used_by"] if u != "Flags:SI-038"]) for e in E])),
 ("rev18 bug: Drop-down rows out of order", lambda d: wb_edit(d, lambda w: w.set_xml("Drop-down", re.sub(r'(<row r="27".*?</row>)(.*?)(<row r="33".*?</row>)', r'\3\2\1', w.xml("Drop-down"), flags=re.S)))),
 ("rev18 bug: tab name over 31 characters", lambda d: wb_edit(d, lambda w: w.parts.__setitem__("xl/workbook.xml", w.parts["xl/workbook.xml"].replace(b'name="Change Log"', b'name="Change Log with a very long tab name here"')))),
 ("unflag a welding-exception row (SI-006)", lambda d: wb_edit(d, lambda w: w.set_cells("Matrix 2024", {"A6": ""}))),
 ("mistype a section title", lambda d: wb_edit(d, lambda w: w.set_cells("Matrix 2024", {"E6": "Structural Steels"}))),
 ("edit a certification cell without registering it", lambda d: wb_edit(d, lambda w: w.set_cells("Matrix 2024", {"L6": "Structural Steel Welding (S2)"}))),
 ("mis-translate a permission cell (AND -> OR)", lambda d: wb_edit(d, lambda w: w.set_cells("Matrix 2024", {"I3": "✓  ICC or ACI or NICET"}))),
 ("remove a matrix row from the dropdown map", lambda d: wb_edit(d, lambda w: w.set_cells("Drop-down", {"C27": ""}))),
 ("use an undefined flag value", lambda d: wb_edit(d, lambda w: w.set_cells("Matrix 2024", {"A3": "maybe"}))),
 ("shrink an SI FORM dropdown range", lambda d: wb_edit(d, lambda w: w.set_xml("SI FORM", w.xml("SI FORM").replace("$B$1:$B$29", "$B$1:$B$20", 1)))),
 ("edit the 2018 tab", lambda d: wb_edit(d, lambda w: w.set_cells("SI Qual 2018 (as received)", {"B3": "tampered"}))),
 ("change a source text file", lambda d: open(os.path.join(d, "verification", "sources", "public", "roh_19_1.txt"), "a", encoding="utf8").write("x")),
 ("hand-edit the generated Evidence tab", lambda d: wb_edit(d, lambda w: w.set_cells("Evidence", {"B3": "tampered"}))),
 ("flag a non-colliding row COLLISION", lambda d: wb_edit(d, lambda w: w.set_cells("Matrix 2024", {"A7": "COLLISION"}))),
]


def main():
    base, dst = fresh(); code, fails = run(dst); shutil.rmtree(base)
    print("baseline: exit=%d, failing checks=%s" % (code, fails or "none"))
    missed = 0
    for name, mut in TESTS:
        base, dst = fresh()
        try:
            mut(dst); code, fails = run(dst)
        finally:
            shutil.rmtree(base, ignore_errors=True)
        if code == 0 or not fails: missed += 1
        print("%-50s caught by: %s" % (name, ", ".join(fails) or "NOTHING"))
    print("%d of %d mutations caught" % (len(TESTS) - missed, len(TESTS)))
    sys.exit(1 if missed else 0)


if __name__ == "__main__":
    main()
