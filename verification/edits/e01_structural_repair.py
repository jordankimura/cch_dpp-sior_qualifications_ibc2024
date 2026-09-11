"""
Edit 01 - make the workbook open in Excel.

Found 2026-09-11 by opening rev18 with Excel 16 through COM: Excel refused to open it, even
in repair mode. Bisection against DPP's original isolated three defects, all introduced
when the 2024 revision was built outside Excel:

  1. Drop-down tab: row 33 (the explanatory note) was written BEFORE rows 27-29. Excel
     requires rows in ascending order.
  2. Matrix 2024 tab: <autoFilter> was written after <mergeCells>; the file format
     requires it before.
  3. Tab "SI Qualification (2018 as received)" is 35 characters; Excel's limit is 31.
     Renamed "SI Qual 2018 (as received)"; the print-area names that point at it follow.

Also drops empty folder entries from the zip (Excel-saved files do not have them).
No cell value changes. Check X-01 now fails on any of these.
"""
import os, re, sys
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, os.path.join(HERE, "..", "tools"))
from xlsx import Workbook

def apply(path_in, path_out):
    w = Workbook(path_in)
    w.order = [n for n in w.order if not n.endswith("/")]
    x = w.xml("Drop-down"); m = re.search(r"<sheetData>(.*)</sheetData>", x, re.S)
    rows = re.findall(r'<row r="\d+"[^>]*?(?:/>|>.*?</row>)', m.group(1), re.S)
    rows.sort(key=lambda r: int(re.match(r'<row r="(\d+)"', r).group(1)))
    w.set_xml("Drop-down", x[:m.start(1)] + "".join(rows) + x[m.end(1):])
    x = w.xml("Matrix 2024"); af = re.search(r"<autoFilter[^>]*?(?:/>|>.*?</autoFilter>)", x, re.S).group(0)
    x = x.replace(af, "", 1).replace("<mergeCells", af + "<mergeCells", 1); w.set_xml("Matrix 2024", x)
    old, new = "SI Qualification (2018 as received)", "SI Qual 2018 (as received)"
    wb = w.parts["xl/workbook.xml"].decode("utf8").replace('name="%s"' % old, 'name="%s"' % new).replace("'%s'!" % old, "'%s'!" % new)
    w.parts["xl/workbook.xml"] = wb.encode("utf8")
    w.save(path_out)

if __name__ == "__main__":
    apply(sys.argv[1], sys.argv[2])
