"""Hulpmiddel: de macro 'Ophalen uit FO' uitvoeren in LibreOffice (VBA-compatibiliteit) zonder dialogen.

testmodule(): maakt van vba/CashflowNaarPowerPoint_v10.bas een variant zonder #If-blokken en zonder dialogen (het FO-pad
komt uit Invoer!A1, de afsluitende melding gaat naar Invoer!A2). voer_uit(): bouwt daarmee een .xlsm, start LibreOffice
headless, roept UitFOOphalen aan via UNO en bewaart het resultaat als .xlsx. Alleen bruikbaar met pyuno en soffice."""
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time

HIER = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HIER)
BAS = os.path.join(REPO, "vba", "CashflowNaarPowerPoint_v10.bas")


def beschikbaar():
    try:
        import uno  # noqa: F401
    except ImportError:
        return False
    return shutil.which("soffice") is not None


def testmodule(bas_pad=BAS):
    s = open(bas_pad, encoding="utf-8").read()
    t = re.sub(r"#If Mac Then\n.*?\n#Else\n(.*?)\n#End If\n", r"\1\n", s, flags=re.S)          # LibreOffice Basic kent geen #If
    old = ('    pad = Application.GetOpenFilename("Excel-werkboeken (*.xlsx;*.xlsm;*.xlsb),*.xlsx;*.xlsm;*.xlsb", , "FO-werkboek kiezen")\n'
           '    If VarType(pad) = vbBoolean Then Exit Sub\n')
    assert t.count(old) == 1, "bestandskeuze niet gevonden in de macro"
    t = t.replace(old, '    pad = ThisWorkbook.Worksheets(SH_INVOER).Range("A1").Value\n')
    i = t.index('    If MsgBox("Tab Invoer (cashflow, verkocht, transport)')
    j = t.index("Exit Sub\n", i) + len("Exit Sub\n")
    t = t[:i] + t[j:]
    vervang = [
        ('    MsgBox bericht, IIf(Len(waarschuwing) > 0, vbExclamation, vbInformation), "Ophalen uit FO"',
         '    ThisWorkbook.Worksheets(SH_INVOER).Range("A2").Value = "KLAAR: " & bericht'),
        ('    MsgBox bericht, vbExclamation, "Ophalen uit FO"', '    ThisWorkbook.Worksheets(SH_INVOER).Range("A2").Value = "FOUT: " & bericht'),
        ("Set wbF = Workbooks.Open(CStr(pad), UpdateLinks:=0, ReadOnly:=True)", "Set wbF = Workbooks.Open(CStr(pad), 0, True)"),
    ]
    for old, new in vervang:
        assert t.count(old) == 1, old[:60]
        t = t.replace(old, new)
    return t


def voer_uit(fo_pad, uit_pad, werkmap=None, project=None, poort=2093):
    """Voert UitFOOphalen uit op fo_pad; geeft de tekst in Invoer!A2 terug ('KLAAR: ...' of 'FOUT: ...') en schrijft uit_pad (.xlsx)."""
    import uno
    from com.sun.star.beans import PropertyValue
    sys.path.insert(0, REPO)
    import build
    from forecast_builder import data as D
    werkmap = werkmap or tempfile.mkdtemp(prefix="macro_lo_")
    os.makedirs(werkmap, exist_ok=True)
    bas = os.path.join(werkmap, "module_test.bas")
    with open(bas, "w", encoding="utf-8", newline="") as f:
        f.write(testmodule())
    vba_bin = build.vba_bin_maken(None, bas_pad=bas)
    paden = build.bouw(project or D.voorbeeld(), os.path.join(werkmap, "test"), vba_bin, caches=False)
    xlsm = [p for p in paden if p.endswith(".xlsm")][0]
    profiel = os.path.join(werkmap, "lo_profiel")
    os.makedirs(os.path.join(profiel, "user"), exist_ok=True)
    with open(os.path.join(profiel, "user", "registrymodifications.xcu"), "w") as f:
        f.write('<?xml version="1.0" encoding="UTF-8"?>\n<oor:items xmlns:oor="http://openoffice.org/2001/registry" '
                'xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n'
                '<item oor:path="/org.openoffice.Office.Common/Security/Scripting"><prop oor:name="MacroSecurityLevel" oor:op="fuse">'
                '<value>0</value></prop></item>\n</oor:items>\n')
    proc = subprocess.Popen(["soffice", "--headless", "--invisible", "--nologo", "--norestore", f"-env:UserInstallation=file://{profiel}",
                             f"--accept=socket,host=127.0.0.1,port={poort};urp;StarOffice.ComponentContext"],
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    desktop = None
    try:
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
        ctx = None
        for _ in range(90):
            try:
                ctx = resolver.resolve(f"uno:socket,host=127.0.0.1,port={poort};urp;StarOffice.ComponentContext")
                break
            except Exception:
                time.sleep(1)
        if ctx is None:
            raise RuntimeError("geen verbinding met LibreOffice")
        desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

        def prop(n, v):
            p = PropertyValue()
            p.Name, p.Value = n, v
            return p

        doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(xlsm), "_blank", 0, (prop("Hidden", True), prop("MacroExecutionMode", 4)))
        libs = doc.BasicLibraries
        gevonden = None
        for lib in libs.ElementNames:
            libs.loadLibrary(lib)
            for mod in libs.getByName(lib).ElementNames:
                if "Sub UitFOOphalen" in libs.getByName(lib).getByName(mod):
                    gevonden = (lib, mod)
        if not gevonden:
            raise RuntimeError("UitFOOphalen niet gevonden in " + str(list(libs.ElementNames)))
        invoer = doc.Sheets.getByName("Invoer")
        invoer.getCellRangeByName("A1").setString(fo_pad)
        invoer.getCellRangeByName("A2").setString("")
        script = doc.getScriptProvider().getScript(f"vnd.sun.star.script:{gevonden[0]}.{gevonden[1]}.UitFOOphalen?language=Basic&location=document")
        script.invoke((), (), ())
        melding = invoer.getCellRangeByName("A2").getString()
        doc.storeToURL(uno.systemPathToFileUrl(uit_pad), (prop("FilterName", "Calc MS Excel 2007 XML"),))
        doc.close(True)
        return melding
    finally:
        try:
            if desktop is not None:
                desktop.terminate()
        except Exception:
            pass
        time.sleep(1)
        proc.kill()


def vergelijk(project_fo, project_macro):
    """Verschillen tussen wat fo.lees_fo leest en wat de macro op tab Invoer/Woningtypes zette (gelezen met data.lees_werkboek)."""
    fouten = []

    def chk(wat, x, y):
        if isinstance(x, float) or isinstance(y, float):
            ok = x is not None and y is not None and abs(float(x) - float(y)) < 1e-6
        else:
            ok = x == y
        if not ok:
            fouten.append(f"{wat}: python={x!r} macro={y!r}")

    a, b = project_fo, project_macro
    per_b = [r for r in b.periodes if r.get("jaar") is not None]
    chk("aantal periodes", len(a.periodes), len(per_b))
    for i, (pa, pb) in enumerate(zip(a.periodes, per_b)):
        for k in ("jaar", "kw", "kosten", "pct_kosten", "omzet", "pct_omzet", "cf", "cf1000", "cf_vorig"):
            chk(f"periode {i} {k}", pa.get(k), pb.get(k))
    chk("actuals", (a.params["actuals_jaar"], a.params["actuals_kw"]), (b.params["actuals_jaar"], b.params["actuals_kw"]))
    chk("aantal typen", len(a.types), len(b.types))
    for k, (ta, tb) in enumerate(zip(a.types, b.types)):
        for attr in ("naam", "aantal", "koopsom", "grond_pct", "start_jaar", "start_kw", "soort"):
            chk(f"type {k} {attr}", getattr(ta, attr), getattr(tb, attr))
        for wat, xa, xb in (("termijn", ta.termijnen, [t for t in tb.termijnen if t[0]]), ("extra", ta.extras, [t for t in tb.extras if t[0]])):
            chk(f"type {k} {wat}s (aantal)", len(xa), len(xb))
            for i, (x, y) in enumerate(zip(xa, xb)):
                chk(f"type {k} {wat} {i} naam", x[0], y[0])
                chk(f"type {k} {wat} {i} waarde", x[1], y[1])
                chk(f"type {k} {wat} {i} kwartaal", x[2], y[2])
        chk(f"type {k} fee_comp", [c for c in ta.fee_comp if c[1] is not None], [c for c in tb.fee_comp if c[1] is not None])
        chk(f"type {k} fee_termijnen", [c for c in ta.fee_termijnen if c[1] is not None], [c for c in tb.fee_termijnen if c[1] is not None])
    for i in range(min(len(a.periodes), len(per_b))):
        for k in range(min(len(a.types), len(b.types))):
            chk(f"verkocht periode {i} type {k}", a.verkocht[i][k], b.verkocht[i][k])
            chk(f"transport periode {i} type {k}", a.transport[i][k], b.transport[i][k])
    for k in ("bestand", "actuals", "kosten", "opbrengsten"):
        chk(f"fo {k}", a.fo.get(k), b.fo.get(k))
    return fouten
