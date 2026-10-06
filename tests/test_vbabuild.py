"""Tests voor forecast_builder.vbabuild (uitvoerbaar met `python3 -m pytest tests/` of direct).

Vereist: olefile, oletools. Het template en de .bas worden gezocht via de omgevingsvariabelen
VBABUILD_TEMPLATE en VBABUILD_BAS; anders de standaardpaden hieronder.
"""
import os
import random
import struct
import subprocess
import sys
import tempfile
import unittest
import zipfile

import olefile
from oletools.olevba import VBA_Parser, decompress_stream

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)

from forecast_builder import vbabuild as vb  # noqa: E402

HIER = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HIER)
TEMPLATE = os.environ.get("VBABUILD_TEMPLATE", os.path.join(REPO, "vba", "vbaProject_template.bin"))
BAS = os.environ.get("VBABUILD_BAS", os.path.join(REPO, "vba", "CashflowNaarPowerPoint_v10.bas"))
XLSM = os.environ.get("VBABUILD_XLSM", os.path.join(REPO, "out", "Cashflow_scenario.xlsm"))   # ontbreekt -> injectietest overgeslagen


def macros(data: bytes) -> dict:
    """{bestandsnaam: broncode} zoals olevba die decodeert (inclusief Attribute-regels)."""
    return {m[2]: m[3] for m in VBA_Parser("x.bin", data=data).extract_all_macros()}


class CompressionTest(unittest.TestCase):
    def cases(self):
        rnd = random.Random(12345)
        yield b""
        yield b"a"
        yield b"ab"
        yield b"abc"
        yield b"abcabcabc"
        yield b"x" * 10000
        yield b"ab" * 3000
        yield bytes(range(256)) * 40
        yield rnd.randbytes(4096)                       # niet comprimeerbaar, precies 1 chunk
        yield rnd.randbytes(4096 * 2 + 100)
        yield rnd.randbytes(3700) + b"\x01"             # niet-comprimeerbare staart > 3640
        yield rnd.randbytes(4000)
        yield b"Attribute VB_Name = \"Module1\"\r\n" * 500
        with open(BAS, "rb") as f:
            yield f.read()
        for n in (1, 2, 3, 4, 15, 16, 17, 255, 256, 257, 4095, 4096, 4097, 8191, 8192, 8193):
            yield bytes(rnd.randrange(3) for _ in range(n))  # laag alfabet: veel korte matches

    def test_roundtrip_eigen_en_oletools(self):
        for data in self.cases():
            z = vb.compress(data)
            self.assertEqual(z[0], 0x01)
            self.assertEqual(vb.decompress(z), data, "eigen decompress, len=%d" % len(data))
            self.assertEqual(bytes(decompress_stream(bytearray(z))), data, "oletools decompress, len=%d" % len(data))

    def test_decompress_template_streams_zoals_oletools(self):
        ole = olefile.OleFileIO(TEMPLATE)
        for path in ole.listdir():
            if path[0] == "VBA" and path[1] != "_VBA_PROJECT":
                raw = ole.openstream(path).read()
                self.assertEqual(vb.decompress(raw), bytes(decompress_stream(bytearray(raw))), path)

    def test_chunk_headers(self):
        z = vb.compress(b"y" * 9000)
        pos = 1
        chunks = 0
        while pos < len(z):
            hdr = struct.unpack_from("<H", z, pos)[0]
            self.assertEqual((hdr >> 12) & 7, 0b011)
            self.assertTrue(hdr & 0x8000)
            pos += (hdr & 0xFFF) + 3
            chunks += 1
        self.assertEqual(pos, len(z))
        self.assertEqual(chunks, 3)


class DirStreamTest(unittest.TestCase):
    def test_parse_serialize_identiek(self):
        ole = olefile.OleFileIO(TEMPLATE)
        raw = vb.decompress(ole.openstream("VBA/dir").read())
        recs = vb.parse_dir(raw)
        self.assertEqual(vb.serialize_dir(recs), raw)
        self.assertEqual(recs[-1].rid, vb.REC_TERMINATOR)
        mods = vb.dir_modules(recs, vb.dir_codepage(recs))
        self.assertEqual([m.name for m in mods],
                         ["ThisWorkbook", "shDashboard", "shInvoer", "shModel", "shPowerPoint", "CashflowNaarPowerPoint"])
        self.assertEqual(vb.dir_codepage(recs), "cp1252")
        self.assertTrue(all(m.is_document for m in mods[:5]))
        self.assertFalse(mods[5].is_document)


class BuildTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with open(TEMPLATE, "rb") as f:
            cls.template = f.read()
        with open(BAS, "rb") as f:
            cls.bas_text = f.read().decode("utf-8")
        cls.roundtrip = vb.build_vbaproject(cls.template, {})
        cls.nieuw = vb.build_vbaproject(cls.template, {"CashflowNaarPowerPoint": cls.bas_text})

    def check_cfb(self, data: bytes, verwacht: dict):
        """Leest terug met olefile (strikt) en vergelijkt elke stream byte-voor-byte."""
        ole = olefile.OleFileIO(data, raise_defects=olefile.DEFECT_INCORRECT)
        self.assertEqual(len(data) % 512, 0)
        self.assertEqual(data[:8], b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1")
        self.assertEqual(struct.unpack_from("<HHH", data, 24), (0x3E, 3, 0xFFFE))
        paden = {"/".join(p) for p in ole.listdir()}
        self.assertEqual(paden, set(verwacht))
        for pad, inhoud in verwacht.items():
            self.assertEqual(ole.get_size(pad), len(inhoud), pad)
            self.assertEqual(ole.openstream(pad).read(), inhoud, pad)
        self.assertEqual(ole.listdir(streams=False, storages=True), [["VBA"]])
        for e in ole.direntries:
            if e is not None and e.entry_type != 0:
                self.assertEqual(e.color, 1)
        return ole

    def verwachte_streams(self, modules: dict):
        ole = olefile.OleFileIO(self.template)
        recs = vb.parse_dir(vb.decompress(ole.openstream("VBA/dir").read()))
        mods = vb.dir_modules(recs, "cp1252")
        verwacht = {
            "PROJECT": ole.openstream("PROJECT").read(),
            "PROJECTwm": ole.openstream("PROJECTwm").read(),
            "VBA/_VBA_PROJECT": bytes.fromhex("cc61ffff000000"),
        }
        for m in mods:
            raw = ole.openstream("VBA/" + m.stream_name).read()
            if m.name in modules:
                src = modules[m.name].replace("\r\n", "\n").replace("\n", "\r\n").encode("cp1252")
            else:
                src = vb.decompress(raw[m.text_offset:])
            verwacht["VBA/" + m.stream_name] = vb.compress(src)
            for r in m.records:
                if r.rid == vb.REC_MODULEOFFSET:
                    r.data = b"\x00\x00\x00\x00"
        verwacht["VBA/dir"] = vb.compress(vb.serialize_dir(recs))
        return verwacht

    def test_rondreis_zonder_wijzigingen(self):
        orig = olefile.OleFileIO(self.template)
        ole = self.check_cfb(self.roundtrip, self.verwachte_streams({}))
        self.assertEqual([p for p in ole.listdir(streams=True, storages=True)],
                         [p for p in orig.listdir(streams=True, storages=True) if not p[-1].startswith("__SRP_")])
        self.assertEqual(macros(self.roundtrip), macros(self.template))
        self.assertEqual(ole.openstream("VBA/_VBA_PROJECT").read(), bytes.fromhex("cc61ffff000000"))
        # alle MODULEOFFSET = 0
        recs = vb.parse_dir(vb.decompress(ole.openstream("VBA/dir").read()))
        offsets = [struct.unpack("<I", r.data)[0] for r in recs if r.rid == vb.REC_MODULEOFFSET]
        self.assertEqual(len(offsets), 6)
        self.assertEqual(set(offsets), {0})
        # alle andere records byte-voor-byte gelijk aan het origineel
        orig_recs = vb.parse_dir(vb.decompress(orig.openstream("VBA/dir").read()))
        self.assertEqual([(r.rid, r.data) for r in recs if r.rid != vb.REC_MODULEOFFSET],
                         [(r.rid, r.data) for r in orig_recs if r.rid != vb.REC_MODULEOFFSET])

    def test_rondreis_mini_stream_en_root(self):
        ole = olefile.OleFileIO(self.roundtrip, raise_defects=olefile.DEFECT_INCORRECT)
        root = ole.direntries[0]
        self.assertEqual(root.name, "Root Entry")
        self.assertEqual(root.entry_type, 5)
        mini_total = sum(-(-ole.get_size("/".join(p)) // 64) * 64 for p in ole.listdir()
                         if 0 < ole.get_size("/".join(p)) < 4096)
        self.assertEqual(root.size, mini_total)
        self.assertNotEqual(root.isectStart, 0xFFFFFFFE)
        groot = [p for p in ole.listdir() if ole.get_size("/".join(p)) >= 4096]
        self.assertEqual(groot, [["VBA", "CashflowNaarPowerPoint"]])

    def test_nieuwe_module(self):
        self.check_cfb(self.nieuw, self.verwachte_streams({"CashflowNaarPowerPoint": self.bas_text}))
        m = macros(self.nieuw)
        verwacht = self.bas_text.replace("\r\n", "\n").replace("\n", "\r\n")
        self.assertEqual(m["CashflowNaarPowerPoint.bas"], verwacht)
        self.assertIn("\r\n", m["CashflowNaarPowerPoint.bas"])
        self.assertNotIn("\n\n", m["CashflowNaarPowerPoint.bas"].replace("\r\n", ""))
        self.assertTrue(m["CashflowNaarPowerPoint.bas"].startswith('Attribute VB_Name = "CashflowNaarPowerPoint"\r\n'))
        # ruwe bytes in cp1252: de é uit de bron staat er als één byte in
        ole = olefile.OleFileIO(self.nieuw)
        src = vb.decompress(ole.openstream("VBA/CashflowNaarPowerPoint").read())
        self.assertEqual(src, verwacht.encode("cp1252"))
        # documentmodules ongewijzigd
        orig = macros(self.template)
        for k in orig:
            if k != "CashflowNaarPowerPoint.bas":
                self.assertEqual(m[k], orig[k], k)

    def test_onbekende_module(self):
        with self.assertRaises(KeyError):
            vb.build_vbaproject(self.template, {"Bestaatniet": "Option Explicit"})

    def test_cli(self):
        with tempfile.TemporaryDirectory() as d:
            out = os.path.join(d, "vbaProject.bin")
            r = subprocess.run([sys.executable, "-m", "forecast_builder.vbabuild", "--template", TEMPLATE,
                                "--module", "CashflowNaarPowerPoint=" + BAS, "--out", out],
                               cwd=ROOT, capture_output=True, text=True)
            self.assertEqual(r.returncode, 0, r.stderr)
            with open(out, "rb") as f:
                self.assertEqual(f.read(), self.nieuw)

    def test_vergelijk_met_excel_origineel(self):
        """Excel's eigen bestand heeft dezelfde streamgroottes: onze compressor is even compact."""
        orig = olefile.OleFileIO(self.template)
        ole = olefile.OleFileIO(self.roundtrip)
        for p in orig.listdir():
            if p == ["VBA", "dir"]:
                continue
            self.assertEqual(ole.get_size("/".join(p)), orig.get_size("/".join(p)), p)


@unittest.skipUnless(os.path.exists(XLSM), "geen Cashflow_scenario_v2.xlsm")
class InjectTest(unittest.TestCase):
    def test_injecteer_en_open_met_libreoffice(self):
        with open(TEMPLATE, "rb") as f:
            template = f.read()
        with open(BAS, "rb") as f:
            bas = f.read().decode("utf-8")
        nieuw = vb.build_vbaproject(template, {"CashflowNaarPowerPoint": bas})
        with tempfile.TemporaryDirectory() as d:
            pad = os.path.join(d, "test_v10.xlsm")
            with zipfile.ZipFile(XLSM) as zin, zipfile.ZipFile(pad, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = nieuw if item.filename == "xl/vbaProject.bin" else zin.read(item.filename)
                    zout.writestr(item, data)
            with zipfile.ZipFile(pad) as z:
                self.assertEqual(z.read("xl/vbaProject.bin"), nieuw)
                self.assertIsNone(z.testzip())
            self.assertEqual(macros(nieuw), macros(zipfile.ZipFile(pad).read("xl/vbaProject.bin")))
            try:
                r = subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", d, pad],
                                   capture_output=True, text=True, timeout=300)
            except FileNotFoundError:
                self.skipTest("soffice niet gevonden")
            self.assertEqual(r.returncode, 0, r.stderr)
            self.assertTrue(os.path.exists(os.path.join(d, "test_v10.pdf")), r.stdout + r.stderr)


if __name__ == "__main__":
    unittest.main(verbosity=2)
