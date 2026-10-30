"""Rendert gekozen tabbladen van een werkmap naar PNG (via LibreOffice → PDF → pdftoppm), voor visuele controle.

Gebruik: python tools/preview.py <werkmap> <uitvoermap> <blad1,blad2,...> [dpi]
Verborgen bladen in de lijst worden zichtbaar gemaakt; alle andere bladen worden verborgen.
"""
import re
import subprocess
import sys
import zipfile
from pathlib import Path


def main(src, outdir, sheets, dpi=80):
    out = Path(outdir)
    out.mkdir(parents=True, exist_ok=True)
    tmp = out / ("_preview" + Path(src).suffix)
    zin = zipfile.ZipFile(src)
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for n in zin.namelist():
            data = zin.read(n)
            if n == "xl/workbook.xml":
                x = data.decode()

                def fix(m):
                    tag = m.group(0)
                    name = re.search(r'name="([^"]+)"', tag).group(1)
                    tag = re.sub(r'\s+state="[^"]*"', "", tag)
                    if name not in sheets:
                        tag = tag.replace("<sheet ", '<sheet state="hidden" ', 1)
                    return tag
                x = re.sub(r"<sheet [^>]*/>", fix, x)
                x = re.sub(r'activeTab="\d+"', "", x)
                data = x.encode()
            z.writestr(n, data)
    prof = Path(outdir).resolve() / "_lo"
    subprocess.run(["soffice", "--headless", f"-env:UserInstallation=file://{prof}", "--convert-to", "pdf", "--outdir", str(out), str(tmp)],
                   capture_output=True, timeout=600)
    pdf = tmp.with_suffix(".pdf")
    subprocess.run(["pdftoppm", "-r", str(dpi), "-png", str(pdf), str(out / "p")], check=True)
    return sorted(out.glob("p-*.png"))


if __name__ == "__main__":
    print("\n".join(map(str, main(sys.argv[1], sys.argv[2], sys.argv[3].split(","), int(sys.argv[4]) if len(sys.argv) > 4 else 80))))
