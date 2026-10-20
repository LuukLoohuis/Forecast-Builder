"""Sjabloon v19 -> v20: dia 7 (verkoop en transport) met veel DAEB, en een simulatie van dia 6 en 7.

    python3 tools/pptx_v20_daeb.py INVOER_v19.pptx UITVOER_v20.pptx [--render MAP] [--simuleer doorgerekend.xlsx UIT.pptx]

* VP1_GRAFIEK en VP2_GRAFIEK: een reeks erbij op de %-as (kolom AE resp. AF van het gegevensblad, layout.PP_VP_KOL = 32): de
  markering DAEB, een gouden ruit met de reeksnaam als label links ('DAEB · 187 woningen in één keer'). Alleen met veel DAEB (meer
  dan de helft van de woningen) heeft die kolom een waarde en een naam; anders is de kop '(uit)' en haalt de macro de reeks weg.
  Met veel DAEB tonen de balken en het % alleen de koopwoningen (het Model laat de DAEB-balken weg).
* Rechts op dia 7 drie kaartjes VP_STAT1..3 (kader, LABEL, WAARDE, SUB) in de stijl van dia 8; de kaart VP_KAART en de twee
  grafieken worden smaller (zoals VO_KAART / VO_GRAFIEK).
* BT_GRAFIEK: elk rijlabel tonen (tickLblSkip = 1). De bouwtermijnen in twee groepen (koop / DAEB-fees, met kopregel en bedrag per
  fee) komen verder uit tab PowerPoint.

--simuleer: doet op de nieuwe v20 na wat de macro met dia 6 en 7 doet (teksten, BT_GRAFIEK met de tijd-as, VP1/VP2 met '(uit)'),
met de waarden uit een doorgerekend werkboek."""
import argparse
import copy
import io
import os
import re
import sys

from lxml import etree
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from forecast_builder import layout as LY  # noqa: E402
from pptx_v12_bouwtermijnen import set_text_keep_format, shapes_by_name  # noqa: E402
from pptx_v18_scenarios import _cache, render_nl  # noqa: E402
from pptx_v17_kerncijfers import nl  # noqa: E402

GOUD, GOUD_LABEL, GOUD_VLAK = "EDA100", "7A5300", "FFF6DD"
KOL_DAEB = {"VP1_GRAFIEK": get_column_letter(LY.PP_VP_KOL - 1), "VP2_GRAFIEK": get_column_letter(LY.PP_VP_KOL)}   # AE, AF
KAART_W = 6.6                   # breedte VP_KAART (inch), zoals VO_KAART
STAT_L, STAT_W = 7.25, 2.3      # kaartjes rechts, zoals VO_STAT


def _kolom(f):
    m = re.search(r"!\$?([A-Z]+)\$?\d", f or "")
    return m.group(1) if m else None


def _dia(prs, vorm):
    for i, s in enumerate(prs.slides):
        if any(sh.name == vorm for sh in s.shapes):
            return i, s
    sys.exit(f"{vorm} niet gevonden")


def markering(chart, kol):
    """kopie van de mijlpaalreeks (punt met label) als ruit in goud op de %-as, waarden in kolom `kol`"""
    cs = chart._chartSpace
    if any(_kolom(f.text) == kol for f in cs.iter(qn("c:f"))):
        return False
    mijlpaal = next(s for s in cs.xpath(".//c:plotArea/c:lineChart/c:ser") if s.find(qn("c:marker") + "/" + qn("c:symbol")) is not None
                    and s.find(qn("c:marker") + "/" + qn("c:symbol")).get("val") == "circle")
    pct_groep = next(g for g in cs.xpath(".//c:plotArea/c:lineChart") if g.find(qn("c:axId")).get("val") != mijlpaal.getparent().find(qn("c:axId")).get("val"))
    alle = cs.xpath(".//c:plotArea/*/c:ser")
    ser = copy.deepcopy(mijlpaal)
    ser.find(qn("c:idx")).set("val", str(max(int(s.find(qn("c:idx")).get("val")) for s in alle) + 1))
    ser.find(qn("c:order")).set("val", str(max(int(s.find(qn("c:order")).get("val")) for s in alle) + 1))
    f = ser.find(qn("c:val") + "/" + qn("c:numRef") + "/" + qn("c:f"))
    blad = f.text.split("!")[0]
    r2 = int(re.search(r"(\d+)$", f.text).group(1))
    f.text = f"{blad}!${kol}$2:${kol}${r2}"
    _cache(f.getparent(), [None] * (r2 - 1))
    tx = ser.find(qn("c:tx") + "/" + qn("c:strRef"))
    tx.find(qn("c:f")).text = f"{blad}!${kol}$1"
    _cache(tx, ["DAEB"], tekst=True)
    mk = ser.find(qn("c:marker"))
    mk.find(qn("c:symbol")).set("val", "diamond")
    size = mk.find(qn("c:size"))
    if size is not None:
        size.set("val", "9")
    for c in mk.iter(qn("a:srgbClr")):
        c.set("val", GOUD if c.get("val") not in ("FFFFFF",) else "FFFFFF")
    dl = ser.find(qn("c:dLbls"))
    if dl is not None:
        sp = dl.find(qn("c:spPr"))
        if sp is not None:
            vul = sp.find(qn("a:solidFill") + "/" + qn("a:srgbClr"))
            if vul is not None:
                vul.set("val", GOUD_VLAK)
            ln = sp.find(qn("a:ln") + "//" + qn("a:srgbClr"))
            if ln is not None:
                ln.set("val", GOUD)
        for c in dl.iter(qn("a:srgbClr")):
            if c.getparent().getparent().tag in (qn("a:defRPr"), qn("a:rPr")):
                c.set("val", GOUD_LABEL)
        pos = dl.find(qn("c:dLblPos"))
        if pos is not None:
            pos.set("val", "l")
    # na de laatste reeks van de %-groep (vóór dLbls/marker/axId van de groep)
    laatste = pct_groep.findall(qn("c:ser"))[-1]
    laatste.addnext(ser)
    # ingebed werkboek: kop in de nieuwe kolom
    deel = chart.part.chart_workbook
    wb = load_workbook(io.BytesIO(deel.xlsx_part.blob))
    wb.worksheets[0][f"{kol}1"] = "DAEB"
    buf = io.BytesIO()
    wb.save(buf)
    deel.update_from_xlsx_blob(buf.getvalue())
    return True


def _nieuw_id(slide, el):
    """unieke vorm-id op de dia (een kopie houdt anders de id van het origineel)"""
    ids = [int(c.get("id")) for c in slide.shapes._spTree.iter(qn("p:cNvPr"))]
    for c in el.iter(qn("p:cNvPr")):
        c.set("id", str(max(ids) + 1))
        ids.append(max(ids) + 1)


def kaartjes(prs):
    _, d7 = _dia(prs, "VP1_GRAFIEK")
    _, d8 = _dia(prs, "VO_STAT1")
    by7, by8 = shapes_by_name(d7), shapes_by_name(d8)
    if "VP_STAT1" in by7:
        return False
    by7["VP_KAART"].width = Inches(KAART_W)
    for naam in ("VP1_GRAFIEK", "VP2_GRAFIEK"):
        by7[naam].width = Inches(KAART_W - 0.2)
    boom = d7.shapes._spTree
    for n in range(1, 4):
        for deel in ("", "_LABEL", "_WAARDE"):
            el = copy.deepcopy(by8[f"VO_STAT{n}{deel}"]._element)
            _nieuw_id(d7, el)
            boom.append(el)
        by = shapes_by_name(d7)
        for deel in ("", "_LABEL", "_WAARDE"):
            sh = by[f"VO_STAT{n}{deel}"]
            sh.name = f"VP_STAT{n}{deel}"
        by = shapes_by_name(d7)
        sub = copy.deepcopy(by[f"VP_STAT{n}_LABEL"]._element)
        _nieuw_id(d7, sub)
        boom.append(sub)
        sh = [s for s in d7.shapes if s.name == f"VP_STAT{n}_LABEL"][-1]
        sh.name = f"VP_STAT{n}_SUB"
        w = shapes_by_name(d7)[f"VP_STAT{n}_WAARDE"]
        sh.top = Emu(int(w.top + w.height + Inches(0.06)))
        set_text_keep_format(sh, "toelichting")
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                r.font.bold = False
    return True


def labels_alle(chart):
    """BT_GRAFIEK: elk rijlabel tonen (met kopregels en DAEB-fees zijn er meer rijen; anders slaat de as labels over)"""
    for ax in chart._chartSpace.xpath(".//c:plotArea/c:catAx"):
        if ax.find(qn("c:tickLblSkip")) is not None:
            continue
        skip = ax.makeelement(qn("c:tickLblSkip"), {"val": "1"})
        na = ax.find(qn("c:noMultiLvlLbl"))
        if na is not None:
            na.addprevious(skip)
        else:
            ax.append(skip)


# ---- simulatie (dia 6 en 7) ----------------------------------------------------------------------------------------------------
class Gegevens:
    def __init__(self, pad):
        wb = load_workbook(pad, data_only=True)
        m, pp = wb["Model"], wb["PowerPoint"]
        self.tekst = {}
        r = LY.PP_KPI_ROW1
        while pp[f"C{r}"].value:
            vorm = pp[f"F{r}"].value
            if pp[f"B{r}"].value in (6, 7) and vorm:
                v = pp[f"D{r}"].value
                self.tekst[vorm] = "" if v is None else nl(str(v))
            r += 1
        h = lambda naam: m[f"{LY.M_HULP}{LY.H[naam]}"].value  # noqa: E731
        self.t_first, self.t_end = float(h("t_first") or 0), float(h("t_end") or 0)
        self.blok = {}
        for naam, kol, n in (("bt", LY.PP_BT_KOL, int(h("bt_n") or 0)), ("vp", LY.PP_VP_KOL, int(h("vp_n") or 0))):
            c0 = column_index_from_string(LY.PP_BLOK[naam])
            koppen = ["" if pp.cell(7, c0 + j).value is None else str(pp.cell(7, c0 + j).value) for j in range(kol)]
            rijen = [[None if v in (None, "") or (isinstance(v, str) and v.startswith("#")) else v
                      for v in (pp.cell(LY.PP_KPI_ROW1 + i, c0 + j).value for j in range(kol))] for i in range(n)]
            self.blok[naam] = (koppen, rijen)


def vul(chart, koppen, rijen):
    """zoals VulGrafiekVorm: per reeks de kolom uit de formule, kop als naam ('(uit)' = weg), waarden als cache, categorieën uit A"""
    n = len(rijen)
    for ser in chart._chartSpace.xpath(".//c:plotArea/*/c:ser"):
        f = ser.find(qn("c:val") + "/" + qn("c:numRef") + "/" + qn("c:f"))
        kol = _kolom(f.text if f is not None else "")
        if kol is None:
            continue
        j = column_index_from_string(kol) - 1
        kop = koppen[j] if j < len(koppen) else ""
        if kop.endswith("(uit)"):
            ser.getparent().remove(ser)
            continue
        blad = f.text.split("!")[0]
        f.text = f"{blad}!${kol}$2:${kol}${n + 1}"
        _cache(f.getparent(), [r[j] if isinstance(r[j], (int, float)) else None for r in rijen])
        tx = ser.find(qn("c:tx") + "/" + qn("c:strRef"))
        if tx is not None:
            tkol = _kolom(tx.find(qn("c:f")).text)
            tj = column_index_from_string(tkol) - 1 if tkol else j
            _cache(tx, [koppen[tj] if tj < len(koppen) else kop], tekst=True)
        cat = ser.find(qn("c:cat"))
        if cat is not None:
            ref = cat.find(qn("c:strRef")) if cat.find(qn("c:strRef")) is not None else cat.find(qn("c:numRef"))
            if ref is not None:
                if ref.tag == qn("c:numRef"):
                    cat.remove(ref)
                    ref = etree.SubElement(cat, qn("c:strRef"))
                    etree.SubElement(ref, qn("c:f"))
                ref.find(qn("c:f")).text = f"{blad}!$A$2:$A${n + 1}"
                _cache(ref, ["" if r[0] is None else (str(int(r[0])) if isinstance(r[0], float) and r[0] == int(r[0]) else str(r[0]))
                             for r in rijen], tekst=True)


def simuleer(prs, g):
    for i in range(len(prs.slides)):
        by = shapes_by_name(prs.slides[i])
        for vorm, tekst in g.tekst.items():
            sh = by.get(vorm)
            if sh is not None and sh.has_text_frame:
                set_text_keep_format(sh, tekst)
        if "BT_GRAFIEK" in by:
            ch = by["BT_GRAFIEK"].chart
            vul(ch, *g.blok["bt"])
            for ax in ch._chartSpace.xpath(".//c:plotArea/c:valAx"):
                sc = ax.find(qn("c:scaling"))
                for tag, v in (("c:max", g.t_end), ("c:min", g.t_first)):
                    e = sc.find(qn(tag))
                    if e is None:
                        e = etree.SubElement(sc, qn(tag))
                    e.set("val", repr(v))
            for lg in ch._chartSpace.findall(".//" + qn("c:legend")):
                lg.getparent().remove(lg)
            # zoals de macro: rijlabels passend bij het aantal rijen (inch -> pt; tekenvlak ongeveer 90% van de grafiek)
            n = max(1, len(g.blok["bt"][1]))
            letter = max(5.0, min(7.5, 1.6 * by["BT_GRAFIEK"].height / 12700 * 0.9 / n))
            for ax in ch._chartSpace.xpath(".//c:plotArea/c:catAx"):
                for d in ax.iter(qn("a:defRPr")):
                    d.set("sz", str(int(letter * 100)))
        for naam in ("VP1_GRAFIEK", "VP2_GRAFIEK"):
            if naam in by:
                vul(by[naam].chart, *g.blok["vp"])


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("invoer")
    ap.add_argument("uitvoer")
    ap.add_argument("--render", default="")
    ap.add_argument("--simuleer", nargs=2, metavar=("WERKBOEK", "UIT_PPTX"), default=None)
    args = ap.parse_args()
    prs = Presentation(args.invoer)
    _, d7 = _dia(prs, "VP1_GRAFIEK")
    by = shapes_by_name(d7)
    n = sum(markering(by[naam].chart, kol) for naam, kol in KOL_DAEB.items())
    k = kaartjes(prs)
    _, d6 = _dia(prs, "BT_GRAFIEK")
    labels_alle(shapes_by_name(d6)["BT_GRAFIEK"].chart)
    prs.save(args.uitvoer)
    print(f"Opgeslagen: {args.uitvoer}\n  DAEB-markering in {n} grafiek(en); kaartjes VP_STAT1..3: {'ja' if k else 'stonden er al'}")
    if args.render:
        render_nl(args.uitvoer, args.render)
    if args.simuleer:
        prs2 = Presentation(args.uitvoer)
        simuleer(prs2, Gegevens(args.simuleer[0]))
        prs2.save(args.simuleer[1])
        print(f"  simulatie: {args.simuleer[1]}")
        if args.render:
            render_nl(args.simuleer[1], os.path.join(args.render, "simulatie"))


if __name__ == "__main__":
    main()
