"""Volledige oplevering: bouwen → scenario-tests → bouwen met testresultaten → herberekenen → cache invullen → docs.

Gebruik: python tools/release.py <pad-naar-recalc.py> [werkmap]
"""
import json
import re
import shutil
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(HERE))

recalc = sys.argv[1]
work = Path(sys.argv[2] if len(sys.argv) > 2 else "/tmp/rdfo_release")
work.mkdir(parents=True, exist_ok=True)

t = subprocess.run([sys.executable, str(HERE / "test_rdfo.py"), str(work / "tests"), recalc])
results = json.loads((work / "tests" / "testresultaten.json").read_text())

import build_rdfo  # noqa: E402
import inject_cache  # noqa: E402
import package_xlsm  # noqa: E402
import preview_pptx  # noqa: E402
import template_pptx  # noqa: E402
import write_docs  # noqa: E402

raw = work / "RDFO_raw.xlsx"
calc = work / "RDFO_calc.xlsx"
build_rdfo.build(str(raw), test_results=results)
shutil.copy(raw, calc)
info = json.loads(subprocess.run([sys.executable, recalc, str(calc), "600"], capture_output=True, text=True).stdout)
print("herberekening:", info)
# enige toegestane #N/A: bewuste gaten in de BOR%-lijn van de grafiek (versie zonder vastgelegde waarde)
fouten = {k: [loc for loc in v["locations"] if not (k == "#N/A" and re.match(r"(Outlook![P-W]|Woningen!Q|Bandbreedte!A[CE]|Dashboard![MN])\d+$", loc))]
          for k, v in info.get("error_summary", {}).items()}
assert info.get("status") in ("success", "errors_found") and not any(fouten.values()), info
cached = work / "RDFO_cache.xlsx"
n = inject_cache.main(str(raw), str(calc), str(cached))
print(n, "formulecellen van cache voorzien")
print(package_xlsm.main(str(cached), str(ROOT / "RDFO_prototype.xlsm")), "knoppen geplaatst")
# lege templateversie (zelfde formules en macro's, zonder data en historie)
raw_t, calc_t, cached_t = work / "RDFO_template_raw.xlsx", work / "RDFO_template_calc.xlsx", work / "RDFO_template_cache.xlsx"
build_rdfo.build(str(raw_t), test_results=results, template=True)
shutil.copy(raw_t, calc_t)
info_t = json.loads(subprocess.run([sys.executable, recalc, str(calc_t), "600"], capture_output=True, text=True).stdout)
fouten_t = {k: [loc for loc in v["locations"] if not (k == "#N/A" and re.match(r"(Outlook![P-W]|Woningen!Q|Bandbreedte!A[CE]|Dashboard![MN])\d+$", loc))]
            for k, v in info_t.get("error_summary", {}).items()}
assert not any(fouten_t.values()), info_t
inject_cache.main(str(raw_t), str(calc_t), str(cached_t))
print(package_xlsm.main(str(cached_t), str(ROOT / "RDFO_template.xlsm")), "knoppen geplaatst (template)")
write_docs.main(str(work / "tests" / "testresultaten.json"))
print(preview_pptx.main(str(cached), str(ROOT / "docs" / "Voorbeeld_PowerPoint.pptx"), str(work / "preview")), "dia's in voorbeeldpresentatie")
print(template_pptx.main(str(cached), str(ROOT / "docs" / "RDFO_PowerPoint_template.pptx")), "dia's in PowerPoint-template")
sys.exit(t.returncode)
