"""Lege PowerPoint-template: dezelfde dia's, opmaak en indeling als de knop 'Maak PowerPoint', maar zonder cijfers.
Elk vak zegt wat er komt en uit welk tabblad. Leest de dia-specificatie (PPT_Spec) uit de werkmap.

Gebruik: python tools/template_pptx.py <werkmap-met-cache> <uitvoer.pptx>
"""
import re
import sys
from pathlib import Path

from pptx import Presentation
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Pt

sys.path.insert(0, str(Path(__file__).resolve().parent))
from preview_pptx import GEEL, GRIJS, LICHT, LIJN, NAVY, Book, tekst, vlak  # noqa: E402

CH0, CH_ROWS = 400, 24   # zelfde als build_rdfo (grafiekblokken op Dias)

INHOUD = {
    "K_DASH": ("6 kerncijfertegels", "omzet FY, BOR, realisatie YTD, to go, woningverkopen, orderintake · tabblad Dashboard"),
    "K_WON": ("6 tegels woningen", "totaal, verkocht, jaarverwachting, cumulatief, etalage · tabblad Woningen"),
    "T_RDFO": ("Tabel RDFO", "omzet, BOR, AK-dekking, fees, EBIT per versie (Budget, Q1–Q4, upload) · tabblad RDFO"),
    "T_PROJ1": ("Tabel projecten met stoplicht", "omzet, realisatie, to go, BOR %, Δ budget, verkoop/vergunning/netcongestie · tabblad Projecten"),
    "T_PROJ2": ("Tabel projecten (vervolg)", "alleen bij meer dan 18 projecten met PowerPoint = Ja"),
    "T_PROJ3": ("Tabel projecten (vervolg)", "alleen bij meer dan 36 projecten met PowerPoint = Ja"),
    "T_OUT": ("Tabel Outlook", "4 jaar vooruit: omzet, BOR, AK, kansen/risico's, reserves, woningen, orderintake · tabblad Outlook"),
    "T_BB": ("Tabel scenario's", "5 scenario's x lopend jaar + 2 · tabblad Bandbreedte"),
    "T_KR": ("Tabel kansen en risico's", "project, omschrijving, wanneer, soort, bedrag · tabblad Kansen en risico's"),
    "T_RESS": ("Tabel reserves per soort", "commercieel, operationeel, nazorg: reserve en vrijval · tabblad Reserves"),
    "T_RES": ("Tabel reserves per project", "max. 12 projecten + totaal · tabblad Reserves"),
    "T_GEB": ("Tabel per gebied", "budget, vorige, geselecteerd, Δ, BOR · tabblad RDFO (onderaan)"),
    "T_OIQ": ("Tabel orderintake per kwartaal", "tabblad Outlook"),
    "T_OI": ("Tabel recentste orderintake", "tabblad Outlook"),
    "T_WONQ": ("Heattabel woningverkopen", "15 projecten x 8 kwartalen + totaal · tabblad Woningen"),
    "PPT_Bullets": ("Aandachtspunten", "automatisch: controle, rode stoplichten, verliesprojecten, ontbrekende invoer · tabblad Controle/PPT"),
}


def vak(s, x, y, w, h, kop_, uitleg, soort):
    sh = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Pt(x), Pt(y), Pt(w), Pt(h))
    sh.fill.solid()
    sh.fill.fore_color.rgb = LICHT
    sh.line.color.rgb = GRIJS
    sh.line.width = Pt(1)
    sh.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    sh.shadow.inherit = False
    tf = sh.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    for i, (t, sz, vet, kl) in enumerate([(soort.upper(), 9, True, GRIJS), (kop_, 15, True, NAVY), (uitleg, 10, False, GRIJS)]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = t
        p.alignment = PP_ALIGN.CENTER
        for run in p.runs:
            run.font.name, run.font.size, run.font.bold = "Arial", Pt(sz), vet
            run.font.color.rgb = kl


def main(src, out):
    bk = Book(src)
    prs = Presentation()
    prs.slide_width, prs.slide_height = Pt(960), Pt(540)
    blank = prs.slide_layouts[6]
    dias = bk.wb["Dias"]

    s = prs.slides.add_slide(blank)
    vlak(s, 0, 0, 960, 540, GEEL)
    vlak(s, 0, 0, 960, 10, NAVY)
    tekst(s, 72, 64, 600, 30, "[Regio]", 16, NAVY, True)
    tekst(s, 72, 180, 820, 130, "Prognoserapportage [Regio]", 46, NAVY, True)
    vlak(s, 74, 322, 120, 8, NAVY)
    tekst(s, 72, 346, 820, 34, "[Versie] [jaar]  ·  doeljaar [jaar]  ·  peildatum [dd-mm-jjjj]", 20, NAVY)
    tekst(s, 72, 470, 820, 20, "Template — de knop 'Maak PowerPoint' in het RDFO vult deze dia's met de cijfers", 10, NAVY)

    def nieuwe_dia(titel, kern):
        s = prs.slides.add_slide(blank)
        nr = len(prs.slides)
        vlak(s, 0, 0, 960, 6, NAVY)
        vlak(s, 36, 30, 6, 30, GEEL)
        tekst(s, 52, 28, 860, 34, titel, 24, NAVY, True)
        if kern:
            tekst(s, 52, 64, 860, 22, "[Kernboodschap — automatisch, bijv. totaal en verschil t.o.v. budget]", 12, GRIJS)
        ln = s.shapes.add_connector(1, Pt(36), Pt(510), Pt(924), Pt(510))
        ln.line.color.rgb = LIJN
        ln.line.width = Pt(0.75)
        tekst(s, 36, 516, 700, 16, "[Regio]  ·  [versie]  ·  peildatum [dd-mm-jjjj]", 9, GRIJS)
        tekst(s, 880, 516, 44, 16, str(nr), 9, NAVY, True, rechts=True)
        return s

    spec_ws, spec_r = bk.ref("PPT_Spec")
    vorige, sld = None, None
    for row in spec_ws[spec_r]:
        dia, titel, kern, soort, ref, x, y, w, h, kp, _ = (c.value for c in row)
        if dia != vorige:
            sld = nieuwe_dia(titel or "", kern or "")
            vorige = dia
        if kp and not str(kp).startswith("="):
            kp = str(kp).split(":")[0] if re.search(r"\d[.,]\d", str(kp)) else kp   # geen demo-cijfers in de kop
            tekst(sld, x, y - 20, w, 16, kp, 11, NAVY, True)
        if not h:   # tabel: hoogte naar aantal regels
            ws_, r = bk.ref(ref)
            a, b = r.split(":")
            h = min(500 - y, (ws_[b].row - ws_[a].row + 1) * 15)
        if soort == "grafiek":
            naam = dias.cell(row=CH0 + (int(ref) - 1) * CH_ROWS, column=1).value or "Grafiek"
            vak(sld, x, y, w, h, naam, "grafiek uit het RDFO (tabblad Dias), als afbeelding", "grafiek")
        else:
            kop_, uitleg = INHOUD.get(ref, (ref, ""))
            vak(sld, x, y, w, h, kop_, uitleg, {"kpi": "tegels", "opsomming": "tekst"}.get(soort, soort))
    prs.save(out)
    return len(prs.slides)


if __name__ == "__main__":
    print(main(sys.argv[1], sys.argv[2]), "dia's in de template")
