"""Schrijft de handleiding, mapping, productievragen en testresultaten als markdown naar docs/."""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import docs_content as DOC  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
DOCS = ROOT / "docs"


def cell(x):
    return str(x).replace("|", "\\|").replace("\n", " ")


def table(head, rows):
    out = ["| " + " | ".join(head) + " |", "|" + "---|" * len(head)]
    out += ["| " + " | ".join(cell(x) for x in r) + " |" for r in rows]
    return "\n".join(out)


def main(results_path=None):
    DOCS.mkdir(exist_ok=True)
    h = ["# Handleiding RDFO-prototype", "", DOC.INLEIDING, "",
         "## Elk kwartaal in 5 stappen", ""]
    for i, (t, wat, waar, let) in enumerate(DOC.STAPPEN):
        h += [f"**{i+1}. {t}** — {waar}", "", wat, "", f"*Let op:* {let}", ""]
    h += ["## De tabbladen", "", table(["Tabblad", "Wat staat er", "Wat doe je er"], DOC.TABBLADEN), ""]
    h += ["## Stoplicht invullen", ""] + [f"- {x}" for x in DOC.STOPLICHT] + [""]
    h += ["## Een vastgelegde prognose corrigeren", ""] + [f"{i+1}. {x}" for i, x in enumerate(DOC.CORRECTIE)] + [""]
    h += ["## Spelregels", ""] + [f"- {x}" for x in DOC.SPELREGELS] + [""]
    h += ["## Problemen oplossen", "", table(["Probleem", "Oplossing"], DOC.FAQ), ""]
    h += ["## Waarom zo gebouwd", ""] + [f"- **{k}** — {v}" for k, v in DOC.ONTWERPKEUZE] + [""]
    (DOCS / "Handleiding.md").write_text("\n".join(h))

    m = ["# Van bron naar rapport", "", "Elke rapportregel met bron, berekening, aanvullende invoer en open punt. Ook in de werkmap (Handleiding › bijlage B).", "",
         table(["Rapportregel", "Bron", "Berekening", "Aanvullende invoer", "Open punt"], DOC.MAPPING)]
    (DOCS / "Rapportagemapping.md").write_text("\n".join(m) + "\n")
    p = ["# Vragen vóór productie", "", "Het prototype draait op expliciete demo-aannames.", "",
         table(["Nr", "Thema", "Vraag", "Raakt"], [(i + 1,) + q for i, q in enumerate(DOC.PRODUCTIEVRAGEN)])]
    (DOCS / "Productievragen.md").write_text("\n".join(p) + "\n")
    if results_path and Path(results_path).exists():
        res = json.loads(Path(results_path).read_text())
        t = ["# Testresultaten", "",
             "Uitgevoerd met `tools/test_rdfo.py`: kopieën van de werkmap worden aangepast zoals een gebruiker dat doet, "
             "door LibreOffice herberekend en vergeleken met een onafhankelijke pandas-berekening. "
             "De macro's zelf (VBA) zijn hier niet in Excel uitgevoerd; getest is dat het VBA-project leesbaar is en de knoppen naar bestaande macro's verwijzen.", "",
             table(["Nr", "Test", "Uitkomst / bewijs", "Resultaat"], [(i + 1, a, b, c) for i, (a, b, c) in enumerate(res)])]
        (DOCS / "Testresultaten.md").write_text("\n".join(t) + "\n")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else None)
