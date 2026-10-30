"""Maakt van de (van cache voorziene) .xlsx een .xlsm: VBA-project toevoegen en knoppen (vormen met macro) plaatsen.

Gebruik: python tools/package_xlsm.py <invoer.xlsx> <uitvoer.xlsm>
"""
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree
from openpyxl.utils import column_index_from_string

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import vbaproject  # noqa: E402
from build_rdfo import BUTTONS, code_name  # noqa: E402

ROOT = HERE.parent
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
      "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
      "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
REL_DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
REL_VBA = "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
AFTER_DRAWING = ["legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"]


def vba_source(path):
    lines = Path(path).read_text(encoding="utf-8").replace("\r\n", "\n").split("\n")
    return "\n".join(ln for ln in lines if not ln.startswith("Attribute VB_"))


def cell(ref):
    m = re.fullmatch(r"([A-Z]+)(\d+)", ref)
    return column_index_from_string(m.group(1)) - 1, int(m.group(2)) - 1


def button_xml(spec, shape_id):
    _sheet, c1, c2, caption, macro, stijl = spec
    (fc, fr), (tc, tr) = cell(c1), cell(c2)
    fill, txt = ("F5B800", "14284B") if stijl == "geel" else ("14284B", "FFFFFF")
    return f"""<xdr:twoCellAnchor xmlns:xdr="{NS['xdr']}" xmlns:a="{NS['a']}" editAs="absolute">
<xdr:from><xdr:col>{fc}</xdr:col><xdr:colOff>57150</xdr:colOff><xdr:row>{fr}</xdr:row><xdr:rowOff>19050</xdr:rowOff></xdr:from>
<xdr:to><xdr:col>{tc + 1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{tr + 1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
<xdr:sp macro="[0]!{macro}" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{shape_id}" name="Knop {shape_id}" descr="{caption}"/><xdr:cNvSpPr/></xdr:nvSpPr>
<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1800000" cy="540000"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill><a:ln><a:noFill/></a:ln></xdr:spPr>
<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" lIns="72000" tIns="36000" rIns="72000" bIns="36000" rtlCol="0" anchor="ctr"/>
<a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="nl-NL" sz="1100" b="1"><a:solidFill><a:srgbClr val="{txt}"/></a:solidFill>
<a:latin typeface="Arial"/><a:cs typeface="Arial"/></a:rPr><a:t>{caption}</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>"""


def xml(el):
    return etree.tostring(el, xml_declaration=True, encoding="UTF-8", standalone=True)


def main(src, out):
    zin = zipfile.ZipFile(src)
    names = zin.namelist()
    files = {n: zin.read(n) for n in names}

    wb = etree.fromstring(files["xl/workbook.xml"])
    wbrels = etree.fromstring(files["xl/_rels/workbook.xml.rels"])
    target = {r.get("Id"): r.get("Target") for r in wbrels}
    sheets = []
    for s in wb.find("m:sheets", NS):
        t = target[s.get(f"{{{NS['r']}}}id")].lstrip("/")
        sheets.append((s.get("name"), t if t.startswith("xl/") else "xl/" + t))
    sheet_path = dict(sheets)

    ct = etree.fromstring(files["[Content_Types].xml"])
    for o in ct.findall("ct:Override", NS):
        if o.get("PartName") == "/xl/workbook.xml":
            o.set("ContentType", "application/vnd.ms-excel.sheet.macroEnabled.main+xml")
    etree.SubElement(ct, f"{{{NS['ct']}}}Default", Extension="bin", ContentType="application/vnd.ms-office.vbaProject")
    etree.SubElement(wbrels, f"{{{NS['rel']}}}Relationship", Id="rIdVBA1", Type=REL_VBA, Target="vbaProject.bin")

    # knoppen als vormen met macro (bestaande tekening uitbreiden of een nieuwe maken)
    n_drawings = len([n for n in files if re.fullmatch(r"xl/drawings/drawing\d+\.xml", n)])
    shape_id = 1001
    for sheet in sorted({b[0] for b in BUTTONS}):
        sp = sheet_path[sheet]
        relp = sp.replace("worksheets/", "worksheets/_rels/") + ".rels"
        rels = etree.fromstring(files[relp]) if relp in files else etree.Element(f"{{{NS['rel']}}}Relationships", nsmap={None: NS["rel"]})
        drawing = None
        for r_ in rels:
            if r_.get("Type") == REL_DRAWING:
                drawing = "xl/drawings/" + r_.get("Target").split("/")[-1]
        if drawing is None:
            n_drawings += 1
            drawing = f"xl/drawings/drawing{n_drawings}.xml"
            files[drawing] = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                              f'<xdr:wsDr xmlns:xdr="{NS["xdr"]}" xmlns:a="{NS["a"]}"></xdr:wsDr>').encode()
            etree.SubElement(rels, f"{{{NS['rel']}}}Relationship", Id="rIdKnop1", Type=REL_DRAWING,
                             Target=f"../drawings/drawing{n_drawings}.xml")
            etree.SubElement(ct, f"{{{NS['ct']}}}Override", PartName="/" + drawing,
                             ContentType="application/vnd.openxmlformats-officedocument.drawing+xml")
            sx = etree.fromstring(files[sp])
            el = etree.Element(f"{{{NS['m']}}}drawing", nsmap={"r": NS["r"]})
            el.set(f"{{{NS['r']}}}id", "rIdKnop1")
            pos = len(sx)
            for i, child in enumerate(sx):
                if etree.QName(child).localname in AFTER_DRAWING:
                    pos = i
                    break
            sx.insert(pos, el)
            files[sp] = xml(sx)
            files[relp] = xml(rels)
        dr = etree.fromstring(files[drawing])
        for b in BUTTONS:
            if b[0] == sheet:
                dr.append(etree.fromstring(button_xml(b, shape_id)))
                shape_id += 1
        files[drawing] = xml(dr)

    # VBA-project: één documentmodule per blad (codeName) + ThisWorkbook + standaardmodule
    docs = ["ThisWorkbook"] + [code_name(n) for n, _p in sheets]
    assert b'codeName="ThisWorkbook"' in files["xl/workbook.xml"], "workbook codeName ontbreekt"
    for n, p in sheets:
        assert f'codeName="{code_name(n)}"'.encode() in files[p], f"codeName ontbreekt voor {n}"
    files["xl/vbaProject.bin"] = vbaproject.build_vba_project(
        {"modRDFO": vba_source(ROOT / "macros" / "modRDFO.bas")}, docs,
        sheet_code={code_name("Stoplicht"): vba_source(ROOT / "macros" / "shStoplicht.cls")})
    files["[Content_Types].xml"] = xml(ct)
    files["xl/_rels/workbook.xml.rels"] = xml(wbrels)

    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        zout.writestr("[Content_Types].xml", files["[Content_Types].xml"])
        for n in files:
            if n != "[Content_Types].xml":
                zout.writestr(n, files[n])
    return shape_id - 1001


if __name__ == "__main__":
    print(main(sys.argv[1], sys.argv[2]), "knoppen geplaatst")
