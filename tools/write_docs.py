"""Schrijft de documentatie (handleiding, mapping, productievragen, testresultaten) als markdown naar docs/."""
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
    h = ["# Handleiding RDFO-prototype", "",
         "Bronvervanging, verversen, kwartaalvastlegging en correcties. Dezelfde tekst staat in de werkmap (tab *Handleiding*).", ""]
    h += ["## Ontwerpkeuze", ""] + [f"- **{k}** — {v}" for k, v in DOC.ONTWERPKEUZE] + [""]
    for sec, lines in DOC.HANDLEIDING:
        h += [f"## {sec}", ""] + [f"- {ln}" for ln in lines] + [""]
    (DOCS / "Handleiding.md").write_text("\n".join(h))

    m = ["# Bron-naar-rapportagemapping", "",
         "Elke rapportageregel met bron, berekening, benodigde aanvullende invoer en openstaande definitie. Ook in de werkmap (tab *Rapportagemapping*).", "",
         table(["Rapportageregel", "Bron", "Berekening", "Aanvullende invoer", "Openstaande definitie"], DOC.MAPPING)]
    (DOCS / "Rapportagemapping.md").write_text("\n".join(m) + "\n")

    p = ["# Productievragen", "", "Te beantwoorden vóór productie. Het prototype draait op expliciete demo-aannames.", "",
         table(["Nr", "Thema", "Vraag", "Raakt"], [(i + 1,) + q for i, q in enumerate(DOC.PRODUCTIEVRAGEN)])]
    (DOCS / "Productievragen.md").write_text("\n".join(p) + "\n")

    if results_path and Path(results_path).exists():
        res = json.loads(Path(results_path).read_text())
        t = ["# Testresultaten", "",
             "Uitgevoerd met `tools/test_rdfo.py`: kopieën van de werkmap worden aangepast zoals een gebruiker dat doet, "
             "door LibreOffice herberekend en vergeleken met een onafhankelijke pandas-berekening op de bronbestanden.", "",
             table(["Nr", "Test", "Uitkomst / bewijs", "Resultaat"], [(i + 1, a, b, c) for i, (a, b, c) in enumerate(res)])]
        (DOCS / "Testresultaten.md").write_text("\n".join(t) + "\n")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else None)
