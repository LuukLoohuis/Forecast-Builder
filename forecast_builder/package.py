"""Nabewerking van het door openpyxl bewaarde bestand: echte grafiek-XML en (optioneel) de VBA-module."""
import re
import zipfile

VBA_REL = '<Relationship Id="rIdVba1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>'
CT_BIN = '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>'
CT_XLSM = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
CT_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"


def lees_vba(pad_xlsm):
    """Haalt xl/vbaProject.bin uit een bestaand .xlsm."""
    with zipfile.ZipFile(pad_xlsm) as z:
        return z.read("xl/vbaProject.bin")


def nabewerken(pad_in, pad_uit, chart_xmls, vba_bin=None):
    """Kopieert het pakket, vervangt chart1..n.xml en voegt de VBA toe als vba_bin is meegegeven (.xlsm)."""
    with zipfile.ZipFile(pad_in) as zin, zipfile.ZipFile(pad_uit, "w", zipfile.ZIP_DEFLATED) as zout:
        namen = zin.namelist()
        for naam in namen:
            data = zin.read(naam)
            m = re.fullmatch(r"xl/charts/chart(\d+)\.xml", naam)
            if m and int(m.group(1)) <= len(chart_xmls):
                data = chart_xmls[int(m.group(1)) - 1].encode("utf-8")
            elif vba_bin is not None and naam == "[Content_Types].xml":
                s = data.decode("utf-8")
                if 'Extension="bin"' not in s:
                    s = s.replace("<Default ", CT_BIN + "<Default ", 1)
                s = s.replace(CT_XLSX, CT_XLSM)
                data = s.encode("utf-8")
            elif vba_bin is not None and naam == "xl/_rels/workbook.xml.rels":
                s = data.decode("utf-8")
                if "vbaProject" not in s:
                    s = s.replace("</Relationships>", VBA_REL + "</Relationships>")
                data = s.encode("utf-8")
            elif vba_bin is not None and naam == "xl/workbook.xml":
                s = data.decode("utf-8")
                if "codeName" not in s:
                    s = s.replace("<workbookPr/>", '<workbookPr codeName="ThisWorkbook"/>').replace("<workbookPr />", '<workbookPr codeName="ThisWorkbook"/>')
                data = s.encode("utf-8")
            zout.writestr(naam, data)
        if vba_bin is not None and "xl/vbaProject.bin" not in namen:
            zout.writestr("xl/vbaProject.bin", vba_bin)
