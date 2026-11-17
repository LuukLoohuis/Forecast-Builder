"""Gedeelde helpers: workbook-zip bewerken, bladen toevoegen."""
import re, zipfile
from patch_ak import Sheet, sheet_paths, N
from lxml import etree
MAIN='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships'

def skeleton(cols='', dv='', tab=None, freeze=None, hidden_grid=False):
    pr = f'<sheetPr><tabColor rgb="{tab}"/></sheetPr>' if tab else ''
    sv = '<sheetView workbookViewId="0"' + (' showGridLines="0"' if hidden_grid else '')
    if freeze: sv += f'><pane ySplit="{freeze-1}" topLeftCell="A{freeze}" activePane="bottomLeft" state="frozen"/></sheetView>'
    else: sv += '/>'
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="{MAIN}" xmlns:r="{REL}">{pr}'
            f'<dimension ref="A1"/><sheetViews>{sv}</sheetViews><sheetFormatPr defaultRowHeight="14.4"/>{cols}<sheetData/>{dv}'
            '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>').encode()

class Book:
    """Laadt alle zip-entries in geheugen; bladen als Sheet-objecten op aanvraag."""
    def __init__(self, path):
        z = zipfile.ZipFile(path)
        self.infos = z.infolist(); self.data = {i.filename: z.read(i.filename) for i in self.infos}
        self.paths = sheet_paths(z); self.sheets = {}; self.new = []
    def sheet(self, name):
        p = self.paths[name]
        if p not in self.sheets: self.sheets[p] = Sheet(self.data[p])
        return self.sheets[p]
    def text(self, fn): return self.data[fn].decode()
    def settext(self, fn, s): self.data[fn] = s.encode()
    def add_sheet(self, name, xml, state=None):
        n = max(int(re.search(r'sheet(\d+)\.xml', p).group(1)) for p in self.data if re.match(r'xl/worksheets/sheet\d+\.xml$', p)) + 1
        fn = f'xl/worksheets/sheet{n}.xml'; self.data[fn] = xml
        wb = self.text('xl/workbook.xml'); rels = self.text('xl/_rels/workbook.xml.rels')
        rid = max(int(x) for x in re.findall(r'Id="rId(\d+)"', rels)) + 1
        sid = max(int(x) for x in re.findall(r'sheetId="(\d+)"', wb)) + 1
        st = f' state="{state}"' if state else ''
        nm = name.replace("'", '&apos;')
        wb = wb.replace('</sheets>', f'<sheet name="{nm}" sheetId="{sid}"{st} r:id="rId{rid}"/></sheets>')
        rels = rels.replace('</Relationships>', f'<Relationship Id="rId{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{n}.xml"/></Relationships>')
        self.settext('xl/workbook.xml', wb); self.settext('xl/_rels/workbook.xml.rels', rels)
        ct = self.text('[Content_Types].xml').replace('</Types>', f'<Override PartName="/xl/worksheets/sheet{n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
        self.settext('[Content_Types].xml', ct)
        self.paths[name] = fn; self.new.append(name)
        # app.xml
        d = self.text('docProps/app.xml')
        m = re.search(r'(<vt:lpstr>Werkbladen</vt:lpstr></vt:variant><vt:variant><vt:i4>)(\d+)', d)
        if m:
            k = int(m.group(2))
            d = d[:m.start(2)] + str(k+1) + d[m.end(2):]
            d = re.sub(r'(<TitlesOfParts><vt:vector size=")(\d+)', lambda x: x.group(1)+str(int(x.group(2))+1), d)
            i0 = d.index('<TitlesOfParts>')
            parts = list(re.finditer(r'<vt:lpstr>.*?</vt:lpstr>', d[i0:]))
            off = i0 + parts[k-1].end()
            d = d[:off] + f'<vt:lpstr>{nm}</vt:lpstr>' + d[off:]
            self.settext('docProps/app.xml', d)
        return self.sheet(name)
    def set_name(self, name, ref):
        wb = self.text('xl/workbook.xml')
        pat = re.compile(r'<definedName name="%s"( [^>]*)?>[^<]*</definedName>' % re.escape(name))
        new = f'<definedName name="{name}">{ref}</definedName>'
        if pat.search(wb): wb = pat.sub(new, wb, count=1)
        else:
            # alfabetisch invoegen
            names = list(re.finditer(r'<definedName name="([^"]+)"', wb))
            pos = None
            for m in names:
                if m.group(1).lower() > name.lower(): pos = m.start(); break
            if pos is None: wb = wb.replace('</definedNames>', new + '</definedNames>')
            else: wb = wb[:pos] + new + wb[pos:]
        self.settext('xml/workbook.xml' if False else 'xl/workbook.xml', wb)
    def save(self, dst):
        for p, s in self.sheets.items(): self.data[p] = s.xml()
        # calcChain weg, volledige herberekening
        self.data.pop('xl/calcChain.xml', None)
        ct = re.sub(r'<Override PartName="/xl/calcChain.xml"[^>]*/>', '', self.text('[Content_Types].xml')); self.settext('[Content_Types].xml', ct)
        rels = re.sub(r'<Relationship [^>]*Target="calcChain.xml"/>', '', self.text('xl/_rels/workbook.xml.rels')); self.settext('xl/_rels/workbook.xml.rels', rels)
        wb = self.text('xl/workbook.xml')
        if 'fullCalcOnLoad' not in wb: wb = re.sub(r'<calcPr([^>]*?)/>', r'<calcPr\1 fullCalcOnLoad="1"/>', wb)
        self.settext('xl/workbook.xml', wb)
        z = zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED)
        done = set()
        for i in self.infos:
            if i.filename in self.data: z.writestr(i, self.data[i.filename]); done.add(i.filename)
        for fn, d in self.data.items():
            if fn not in done: z.writestr(fn, d)
        z.close()

def q(name):
    return "'" + name.replace("'", "''") + "'"
