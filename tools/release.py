"""Volledige oplevering: bouwen → scenario-tests → bouwen met testresultaten → herberekenen → cache invullen → docs.

Gebruik: python tools/release.py <pad-naar-recalc.py> [werkmap]
"""
import json
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
import write_docs  # noqa: E402

raw = work / "RDFO_raw.xlsx"
calc = work / "RDFO_calc.xlsx"
build_rdfo.build(str(raw), test_results=results)
shutil.copy(raw, calc)
info = json.loads(subprocess.run([sys.executable, recalc, str(calc), "600"], capture_output=True, text=True).stdout)
print("herberekening:", info)
assert info.get("total_errors") == 0, info
n = inject_cache.main(str(raw), str(calc), str(ROOT / "RDFO_prototype.xlsx"))
print(n, "formulecellen van cache voorzien")
write_docs.main(str(work / "tests" / "testresultaten.json"))
sys.exit(t.returncode)
