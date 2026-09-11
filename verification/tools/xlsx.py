"""
Edits the workbook at the file-format level, so nothing it does not touch changes.

Why this exists: opening and re-saving the workbook with openpyxl silently DELETES the
SI FORM dropdowns (openpyxl warns "Data Validation extension is not supported and will be
removed"). Every edit to the deliverable goes through this module instead. Checks I-03 and
I-04 in verify_si.py fail if an edit ever damages the SI FORM tab.

All text written here is stored as inline strings, the same way the added tabs already are.
"""
import re, zipfile
from xml.sax.saxutils import escape

NS_WS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
CT_WS = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"

# Styles already defined in styles.xml by the tabs added in the 2024 revision.
STYLE_TITLE, STYLE_HEADER, STYLE_BODY = 130, 124, 125


def col_letter(n):
    s = ""
    while n: n, r = divmod(n - 1, 26); s = chr(65 + r) + s
    return s


def col_number(letters):
    n = 0
    for ch in letters: n = n * 26 + ord(ch) - 64
    return n


def split_ref(ref):
    m = re.match(r"^([A-Z]+)(\d+)$", ref)
    return m.group(1), int(m.group(2))


def cell_xml(ref, text, style):
    if text is None or text == "":
        return '<c r="%s" s="%d"/>' % (ref, style)
    return '<c r="%s" s="%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (
        ref, style, escape(str(text)))


class Workbook:
    def __init__(self, path):
        z = zipfile.ZipFile(path)
        self.order = [i.filename for i in z.infolist()]
        self.info = {i.filename: i for i in z.infolist()}
        self.parts = {n: z.read(n) for n in self.order}

    # ---- structure -----------------------------------------------------------------
    def _wb(self): return self.parts["xl/workbook.xml"].decode("utf8")
    def _rels(self): return self.parts["xl/_rels/workbook.xml.rels"].decode("utf8")

    def sheets(self):
        """[(name, sheetId, rId, part path)] in tab order."""
        rels = dict(re.findall(r'<Relationship Id="([^"]+)" Type="%s" Target="([^"]+)"/>' % re.escape(NS_WS), self._rels()))
        out = []
        for name, sid, rid in re.findall(r'<sheet name="([^"]+)" sheetId="(\d+)" r:id="([^"]+)"/>', self._wb()):
            out.append((unescape_attr(name), int(sid), rid, "xl/" + rels[rid]))
        return out

    def part_of(self, name):
        for n, _sid, _rid, part in self.sheets():
            if n == name: return part
        raise KeyError(name)

    def xml(self, name): return self.parts[self.part_of(name)].decode("utf8")
    def set_xml(self, name, xml): self.parts[self.part_of(name)] = xml.encode("utf8")

    def _reindex_defined_names(self, old_names, new_names):
        wb = self._wb()
        def fix(m):
            old = int(m.group(1)); nm = old_names[old]
            return 'localSheetId="%d"' % new_names.index(nm)
        wb = re.sub(r'localSheetId="(\d+)"', fix, wb)
        self.parts["xl/workbook.xml"] = wb.encode("utf8")

    def add_sheet(self, name, xml, position):
        old_names = [s[0] for s in self.sheets()]
        if name in old_names: raise ValueError("sheet exists: " + name)
        wb, rels = self._wb(), self._rels()
        sid = max(s[1] for s in self.sheets()) + 1
        nums = [int(x) for x in re.findall(r"worksheets/sheet(\d+)\.xml", rels)]
        part = "xl/worksheets/sheet%d.xml" % (max(nums) + 1)
        rid = "rId%d" % (max(int(x) for x in re.findall(r'Id="rId(\d+)"', rels)) + 1)
        rels = rels.replace("</Relationships>", '<Relationship Id="%s" Type="%s" Target="%s"/></Relationships>' % (rid, NS_WS, part[3:]))
        tags = re.findall(r'<sheet name="[^"]+" sheetId="\d+" r:id="[^"]+"/>', wb)
        tags.insert(position, '<sheet name="%s" sheetId="%d" r:id="%s"/>' % (escape(name, {'"': "&quot;"}), sid, rid))
        wb = re.sub(r"<sheets>.*?</sheets>", "<sheets>" + "".join(tags) + "</sheets>", wb, flags=re.S)
        self.parts["xl/workbook.xml"] = wb.encode("utf8")
        self.parts["xl/_rels/workbook.xml.rels"] = rels.encode("utf8")
        ct = self.parts["[Content_Types].xml"].decode("utf8").replace(
            "</Types>", '<Override PartName="/%s" ContentType="%s"/></Types>' % (part, CT_WS))
        self.parts["[Content_Types].xml"] = ct.encode("utf8")
        self.parts[part] = xml.encode("utf8"); self.order.append(part)
        self._reindex_defined_names(old_names, [s[0] for s in self.sheets()])

    def remove_sheet(self, name):
        old_names = [s[0] for s in self.sheets()]
        _n, _sid, rid, part = [s for s in self.sheets() if s[0] == name][0]
        if re.search(r'localSheetId="%d"' % old_names.index(name), self._wb()):
            raise ValueError("sheet has defined names; refusing to remove: " + name)
        wb = re.sub(r'<sheet name="[^"]+" sheetId="\d+" r:id="%s"/>' % rid, "", self._wb())
        self.parts["xl/workbook.xml"] = wb.encode("utf8")
        rels = re.sub(r'<Relationship Id="%s" [^>]*/>' % rid, "", self._rels())
        self.parts["xl/_rels/workbook.xml.rels"] = rels.encode("utf8")
        ct = re.sub(r'<Override PartName="/%s" [^>]*/>' % re.escape(part), "", self.parts["[Content_Types].xml"].decode("utf8"))
        self.parts["[Content_Types].xml"] = ct.encode("utf8")
        for p in (part, part.replace("worksheets/", "worksheets/_rels/") + ".rels"):
            if p in self.parts: del self.parts[p]; self.order.remove(p)
        new_names = [s[0] for s in self.sheets()]
        self._reindex_defined_names_after_remove(old_names, new_names)

    def _reindex_defined_names_after_remove(self, old_names, new_names):
        wb = self._wb()
        wb = re.sub(r'localSheetId="(\d+)"', lambda m: 'localSheetId="%d"' % new_names.index(old_names[int(m.group(1))]), wb)
        self.parts["xl/workbook.xml"] = wb.encode("utf8")

    def replace_sheet(self, name, xml):
        if name in [s[0] for s in self.sheets()]: self.set_xml(name, xml)
        else: raise KeyError(name)

    # ---- cells ---------------------------------------------------------------------
    def set_cells(self, name, values, style=None):
        """values: {ref: text}. Keeps a cell's existing style unless one is given."""
        xml = self.xml(name)
        for ref, text in values.items():
            col, row = split_ref(ref)
            m = re.search(r'<c r="%s"(?P<attrs>[^>]*?)(?:/>|>(?P<body>.*?)</c>)' % ref, xml, re.S)
            if m:
                sm = re.search(r's="(\d+)"', m.group("attrs"))
                st = style if style is not None else int(sm.group(1)) if sm else STYLE_BODY
                xml = xml[:m.start()] + cell_xml(ref, text, st) + xml[m.end():]
                continue
            st = style if style is not None else STYLE_BODY
            rm = re.search(r'<row r="%d"(?P<attrs>[^>]*?)(?:/>|>(?P<body>.*?)</row>)' % row, xml, re.S)
            if rm:
                body = rm.group("body") or ""
                cells = re.findall(r'<c r="([A-Z]+)\d+"[^>]*?(?:/>|>.*?</c>)', body, re.S)
                pieces = re.findall(r'<c r="[A-Z]+\d+"[^>]*?(?:/>|>.*?</c>)', body, re.S)
                idx = sum(1 for c in cells if col_number(c) < col_number(col))
                pieces.insert(idx, cell_xml(ref, text, st))
                attrs = rm.group("attrs")
                xml = xml[:rm.start()] + '<row r="%d"%s>%s</row>' % (row, attrs, "".join(pieces)) + xml[rm.end():]
                continue
            rows = [(int(r), mm) for mm in re.finditer(r'<row r="(\d+)"', xml) for r in [mm.group(1)]]
            after = [mm for r, mm in rows if r > row]
            new_row = '<row r="%d">%s</row>' % (row, cell_xml(ref, text, st))
            if after: xml = xml[:after[0].start()] + new_row + xml[after[0].start():]
            else: xml = xml.replace("</sheetData>", new_row + "</sheetData>")
        self.set_xml(name, xml)

    def get_cell(self, name, ref):
        m = re.search(r'<c r="%s"[^>]*?(?:/>|>(.*?)</c>)' % ref, self.xml(name), re.S)
        if not m or not m.group(1): return None
        t = re.findall(r"<t[^>]*>(.*?)</t>", m.group(1), re.S)
        return unescape_attr("".join(t)) if t else None

    def save(self, path):
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for n in self.order:
                if n in self.parts:
                    zi = zipfile.ZipInfo(n, date_time=self.info[n].date_time if n in self.info else (2026, 9, 11, 0, 0, 0))
                    zi.compress_type = zipfile.ZIP_DEFLATED
                    z.writestr(zi, self.parts[n])


def unescape_attr(s):
    return s.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"').replace("&apos;", "'").replace("&amp;", "&")


def row_height(values, widths):
    """Excel does not grow a wrapped row on open, so estimate it: about one character per
    unit of column width at the 9-10pt fonts these tabs use."""
    import math
    lines = 1
    for v, w in zip(values, widths):
        if v in (None, ""): continue
        n = sum(max(1, math.ceil(len(part) / max(1.0, float(w) * 1.05))) for part in str(v).split("\n"))
        lines = max(lines, n)
    return max(15.0, round(12.6 * lines + 3, 1))


def table_sheet(title, headers, rows, widths, freeze_row=2, note_rows=(), autofilter=False, merge_notes=False):
    """A tab in the same style as the other tabs added in the 2024 revision:
    title in row 1, headers in row 2, data from row 3, optional notes after a blank row."""
    ncols = len(headers)
    cols = "".join('<col min="%d" max="%d" width="%s" customWidth="1"/>' % (i + 1, i + 1, w) for i, w in enumerate(widths))
    out = ['<row r="1" ht="24" customHeight="1">%s</row>' % cell_xml("A1", title, STYLE_TITLE),
           '<row r="2" ht="%s" customHeight="1">%s</row>' % (row_height(headers, widths),
                                                            "".join(cell_xml(col_letter(i + 1) + "2", h, STYLE_HEADER) for i, h in enumerate(headers)))]
    r = 3
    for row in rows:
        out.append('<row r="%d" ht="%s" customHeight="1">%s</row>' % (r, row_height(row, widths), "".join(
            cell_xml(col_letter(i + 1) + str(r), v, STYLE_BODY) for i, v in enumerate(row) if v not in (None, ""))))
        r += 1
    last_data = r - 1
    merges = []
    if note_rows:
        r += 1
        total = sum(float(w) for w in widths)
        for n in note_rows:
            out.append('<row r="%d" ht="%s" customHeight="1">%s</row>' % (r, row_height([n], [total if merge_notes else widths[0]]),
                                                                        cell_xml("A%d" % r, n, STYLE_BODY)))
            if merge_notes: merges.append("A%d:%s%d" % (r, col_letter(ncols), r))
            r += 1
    last = "%s%d" % (col_letter(ncols), max(r - 1, 2))
    pane = ('<pane ySplit="%d" topLeftCell="A%d" activePane="bottomLeft" state="frozen"/>' % (freeze_row, freeze_row + 1)) if freeze_row else ""
    af = '<autoFilter ref="A2:%s%d"/>' % (col_letter(ncols), max(last_data, 2)) if autofilter else ""
    mc = ('<mergeCells count="%d">%s</mergeCells>' % (len(merges), "".join('<mergeCell ref="%s"/>' % m for m in merges))) if merges else ""
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>'
            '<dimension ref="A1:%s"/><sheetViews><sheetView showGridLines="0" workbookViewId="0">%s</sheetView></sheetViews>'
            '<sheetFormatPr defaultRowHeight="12.75"/><cols>%s</cols><sheetData>%s</sheetData>%s%s'
            '<pageMargins left="0.3" right="0.3" top="0.4" bottom="0.4" header="0.3" footer="0.3"/>'
            '<pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0"/></worksheet>') % (
            last, pane, cols, "".join(out), af, mc)


def read_all(path):
    """{sheet name: {ref: text}} for every tab, straight from the file (shared + inline strings)."""
    wbk = Workbook(path)
    shared = []
    if "xl/sharedStrings.xml" in wbk.parts:
        for si in re.findall(r"<si>(.*?)</si>", wbk.parts["xl/sharedStrings.xml"].decode("utf8"), re.S):
            shared.append(unescape_attr("".join(re.findall(r"<t[^>]*>(.*?)</t>", si, re.S))))
    out = {}
    for name, _sid, _rid, part in wbk.sheets():
        cells = {}
        for m in re.finditer(r'<c r="([A-Z]+\d+)"([^>]*?)(?:/>|>(.*?)</c>)', wbk.parts[part].decode("utf8"), re.S):
            ref, attrs, body = m.group(1), m.group(2), m.group(3) or ""
            if 't="s"' in attrs:
                v = re.search(r"<v>(\d+)</v>", body); val = shared[int(v.group(1))] if v else None
            elif 't="inlineStr"' in attrs:
                val = unescape_attr("".join(re.findall(r"<t[^>]*>(.*?)</t>", body, re.S)))
            else:
                v = re.search(r"<v>(.*?)</v>", body); val = unescape_attr(v.group(1)) if v else None
            if val not in (None, ""): cells[ref] = val
        out[name] = cells
    return out
