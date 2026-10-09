"""Statische controles op de VBA-module (geen Excel nodig)."""
import os
import re

HIER = os.path.dirname(os.path.abspath(__file__))
BAS = os.path.join(os.path.dirname(HIER), "vba", "CashflowNaarPowerPoint_v10.bas")


def _procedures(src):
    """{naam: body} van alle Subs en Functions."""
    uit = {}
    for m in re.finditer(r"^(?:Private |Public )?(?:Sub|Function) (\w+)\(.*?\n(.*?)^End (?:Sub|Function)", src, flags=re.M | re.S):
        uit[m.group(1)] = m.group(2)
    return uit


def test_geen_functie_roept_zichzelf_aan():
    """Een functie die haar eigen naam aanroept (Tekst = Tekst(v)) loopt in Excel vast op 'Out of stack space'."""
    src = open(BAS, encoding="utf-8").read()
    procs = _procedures(src)
    assert len(procs) > 40
    recursief = [n for n, body in procs.items() if re.search(rf"^\s*(?!')[^'\n]*\b{n}\s*\(", body, flags=re.M | re.I)]
    assert recursief == [], recursief


def test_foutafhandelaars_bewaren_err():
    """Een On Error-statement wist het Err-object: in een handler moet Err.Number vóór 'On Error Resume Next' zijn vastgelegd."""
    src = open(BAS, encoding="utf-8").read()
    for m in re.finditer(r"^Fout:\n(.*?)^End Sub", src, flags=re.M | re.S):
        body = m.group(1)
        i = body.find("On Error Resume Next")
        if i >= 0:
            assert "Err.Number" in body[:i], "Err.Number pas na On Error Resume Next gelezen:\n" + body
            assert "Err.Number" not in body[i:], "Err.Number na On Error Resume Next gelezen (is dan altijd 0):\n" + body


def test_geen_procedure_heet_als_parameter():
    """Een procedure met dezelfde naam als een parameter of variabele elders is verwarrend en gaf in LibreOffice een crash."""
    src = open(BAS, encoding="utf-8").read()
    procs = set(n.lower() for n in _procedures(src))
    namen = set()
    for m in re.finditer(r"\b(?:Dim|Const)\s+(.+)$", src, flags=re.M):
        namen |= {w.group(1).lower() for part in m.group(1).split(",") if (w := re.match(r"\s*(\w+)", part))}
    for m in re.finditer(r"^(?:Private |Public )?(?:Sub|Function) \w+\((.*?)\)", src, flags=re.M):
        namen |= {w.group(1).lower() for part in m.group(1).split(",") if (w := re.match(r"\s*(?:ByVal |ByRef |Optional )*(\w+)", part)) and w.group(1)}
    assert procs & namen == set(), procs & namen
