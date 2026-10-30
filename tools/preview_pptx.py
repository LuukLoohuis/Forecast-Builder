"""Maakt een voorbeeldpresentatie zoals de knop 'Maak PowerPoint' die maakt: zelfde dia-specificatie (PPT_Spec),
bewerkbare tabellen, kerncijfertegels en grafieken als afbeelding (via LibreOffice gerenderd).

Gebruik: python tools/preview_pptx.py <werkmap-met-cache> <uitvoer.pptx> <werkmap-voor-tijdelijke-bestanden>
"""
import datetime as dt
import re
import sys
import warnings
from pathlib import Path

from openpyxl import load_workbook
from PIL import Image, ImageChops
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Pt

sys.path.insert(0, str(Path(__file__).resolve().parent))
import preview  # noqa: E402

warnings.filterwarnings("ignore")
NAVY, GEEL, GRIJS, LICHT, LIJN, WIT = (RGBColor(20, 40, 75), RGBColor(245, 184, 0), RGBColor(107, 115, 133),
                                       RGBColor(243, 245, 249), RGBColor(217, 222, 232), RGBColor(255, 255, 255))
SL = {3: RGBColor(0x2E, 0x9E, 0x44), 2: RGBColor(0xF3, 0x9C, 0x12), 1: RGBColor(0xD6, 0x3B, 0x2F), 0: RGBColor(0xC5, 0xCA, 0xD3)}
ROOD, GROEN, GEELLICHT = RGBColor(0xC0, 0x39, 0x2B), RGBColor(0x3B, 0x7A, 0x1E), RGBColor(0xFF, 0xF4, 0xCC)


# --------------------------------------------------------------------------- getalnotatie (zoals Excel, Nederlandse scheidingstekens)
def _sections(nf):
    out, cur, q = [], "", False
    for ch in nf:
        if ch == '"':
            q = not q
        if ch == ";" and not q:
            out.append(cur)
            cur = ""
        else:
            cur += ch
    return out + [cur]


def fmt(v, nf):
    """-> (tekst, rood?)"""
    if v is None or v == "":
        return "", False
    if isinstance(v, str):
        return v, False
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime("%m-%Y" if nf.startswith("mm-yyyy") else "%d-%m-%y"), False
    if nf == '"●";"●";"○"':
        return ("○" if v == 0 else "●"), False
    if (not nf or nf == "General") and float(v).is_integer() and 1900 <= v <= 2100:
        return str(int(v)), False
    if not nf or nf == "General":
        return (f"{v:,.0f}" if abs(v) >= 100 else f"{v:,.2f}".rstrip("0").rstrip(".")).replace(",", "X").replace(".", ",").replace("X", "."), False
    secs = _sections(nf)
    sec = secs[0]
    if v < 0 and len(secs) > 1:
        sec, v2 = secs[1], -v
    elif v == 0 and len(secs) > 2:
        sec, v2 = secs[2], v
    else:
        v2 = v
    red = "[Red]" in sec
    sec = sec.replace("[Red]", "")
    lits = re.findall(r'"([^"]*)"', sec)
    core = re.sub(r'"[^"]*"', "\x00", sec)
    m = re.search(r"[#0][#0,]*(\.0+)?%?", core)
    if not m:
        return "".join(lits) if lits else "", red
    num = m.group(0)
    dec = len(m.group(1)) - 1 if m.group(1) else 0
    val = v2 * 100 if num.endswith("%") else v2
    s = f"{val:,.{dec}f}" if "," in num else f"{val:.{dec}f}"
    s = s.replace(",", "X").replace(".", ",").replace("X", ".") + ("%" if num.endswith("%") else "")
    pre, post = core[:m.start()], core[m.end():]
    it = iter(lits)
    pre = "".join(next(it) if ch == "\x00" else ch for ch in pre)
    post = "".join(next(it) if ch == "\x00" else ch for ch in post)
    return pre + s + post, red


# --------------------------------------------------------------------------- vormen
def vlak(sld, x, y, w, h, kleur):
    sh = sld.shapes.add_shape(MSO_SHAPE.RECTANGLE, Pt(x), Pt(y), Pt(w), Pt(h))
    sh.fill.solid()
    sh.fill.fore_color.rgb = kleur
    sh.line.fill.background()
    sh.shadow.inherit = False
    return sh


def tekst(sld, x, y, w, h, t, grootte, kleur, vet=False, rechts=False):
    tb = sld.shapes.add_textbox(Pt(x), Pt(y), Pt(w), Pt(h))
    tf = tb.text_frame
    tf.word_wrap = True
    for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, m, 0)
    p = tf.paragraphs[0]
    p.text = str(t)
    if rechts:
        p.alignment = PP_ALIGN.RIGHT
    if p.runs:
        f = p.runs[0].font
        f.name, f.size, f.bold = "Arial", Pt(grootte), vet
        f.color.rgb = kleur
    return tb


class Book:
    def __init__(self, path):
        self.wb = load_workbook(path, data_only=True)

    def ref(self, name):
        sheet, r = self.wb.defined_names[name].attr_text.split("!")
        return self.wb[sheet], r.replace("$", "")

    def val(self, name):
        ws, r = self.ref(name)
        v = ws[r].value
        return "" if v is None else v

    def text(self, name):
        ws, r = self.ref(name)
        return fmt(ws[r].value, ws[r].number_format)[0]


def trim(path):
    im = Image.open(path).convert("RGB")
    box = ImageChops.difference(im, Image.new("RGB", im.size, (255, 255, 255))).getbbox()
    (im.crop(box) if box else im).save(path)


def main(src, out, workdir):
    bk = Book(src)
    wb = bk.wb
    prs = Presentation()
    prs.slide_width, prs.slide_height = Pt(960), Pt(540)
    blank = prs.slide_layouts[6]
    sel = str(bk.val("Sel_Versie"))

    # grafieken: Dias-grafiekblokken staan als laatste pagina's in de pdf
    pngs = preview.main(src, Path(workdir) / "dias", ["Dias"], dpi=160)
    chart_png = {}
    total = int(str(bk.val("PPT_ChartN")) or 0)
    for i, p in enumerate(pngs[-total:]):
        trim(p)
        chart_png[i + 1] = p

    # titeldia (geel vlak, donkerblauwe letters)
    s = prs.slides.add_slide(blank)
    vlak(s, 0, 0, 300, 540, GEEL)
    vlak(s, 300, 0, 8, 540, NAVY)
    tekst(s, 40, 60, 230, 30, bk.val("P_Regio"), 14, NAVY, True)
    tekst(s, 40, 470, 230, 40, "Voorbeeld — gemaakt met tools/preview_pptx.py", 9, NAVY)
    tekst(s, 360, 190, 560, 110, bk.val("PPT_Titel"), 38, NAVY, True)
    vlak(s, 362, 304, 90, 6, GEEL)
    tekst(s, 360, 324, 560, 30, bk.val("PPT_Sub"), 16, NAVY)
    tekst(s, 360, 356, 560, 20, "Bron: RDFO_prototype.xlsm", 10, GRIJS)

    def nieuwe_dia(titel, kern):
        s = prs.slides.add_slide(blank)
        nr = len(prs.slides)
        vlak(s, 0, 0, 960, 6, NAVY)
        vlak(s, 36, 30, 6, 30, GEEL)
        tekst(s, 52, 28, 860, 34, titel, 24, NAVY, True)
        if kern:
            tekst(s, 52, 64, 860, 22, kern, 12, GRIJS)
        ln = s.shapes.add_connector(1, Pt(36), Pt(510), Pt(924), Pt(510))
        ln.line.color.rgb = LIJN
        ln.line.width = Pt(0.75)
        tekst(s, 36, 516, 700, 16, bk.val("PPT_Footer"), 9, GRIJS)
        tekst(s, 880, 516, 44, 16, str(nr), 9, NAVY, True, rechts=True)
        return s

    def kop(s, t, x, y, w):
        if t:
            tekst(s, x, y - 20, w, 16, t, 11, NAVY, True)

    def tegels(s, name, x, y, w, h):
        ws, r = bk.ref(name)
        r1 = ws[r.split(":")[0]].row
        tw = (w - 5 * 8) / 6
        for k in range(6):
            tx = x + k * (tw + 8)
            vlak(s, tx, y, tw, h, LICHT)
            vlak(s, tx, y, tw, 3, GEEL)
            lab, val, sub = (ws.cell(row=r1 + k, column=c) for c in (2, 3, 4))
            tekst(s, tx + 6, y + 8, tw - 12, 14, lab.value or "", 8.5, NAVY, True)
            tekst(s, tx + 6, y + 24, tw - 12, 28, fmt(val.value, val.number_format)[0], 19, NAVY, True)
            tekst(s, tx + 6, y + h - 18, tw - 12, 14, sub.value or "", 8, GRIJS)

    def tabel(s, name, x, y, w, t):
        ws, r = bk.ref(name)
        a, b = r.split(":")
        r1, c1 = ws[a].row, ws[a].column
        r2, c2 = ws[b].row, ws[b].column
        cols = [c for c in range(c1, c2 + 1) if isinstance(ws.cell(row=r1 - 1, column=c).value, (int, float)) and ws.cell(row=r1 - 1, column=c).value > 0]
        rows = [rr for rr in range(r1, r2 + 1) if rr == r1 or ws.cell(row=rr, column=c1).value not in (None, "")]
        som = sum(ws.cell(row=r1 - 1, column=c).value for c in cols)
        codes = {c: ws.cell(row=r1 - 2, column=c).value for c in cols}
        heads = {c: str(ws.cell(row=r1, column=c).value or "") for c in cols}
        grootte = 8 if len(rows) > 15 else 9
        kop(s, t, x, y, w)
        gf = s.shapes.add_table(len(rows), len(cols), Pt(x), Pt(y), Pt(w), Pt(len(rows) * (grootte + 6)))
        tbl = gf.table
        tblPr = tbl._tbl.tblPr
        style = tblPr.find(qn("a:tableStyleId"))
        if style is None:
            style = tblPr.makeelement(qn("a:tableStyleId"), {})
            tblPr.append(style)
        style.text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"
        for j, c in enumerate(cols):
            tbl.columns[j].width = Pt(w * ws.cell(row=r1 - 1, column=c).value / som)
        for i, rr in enumerate(rows):
            tbl.rows[i].height = Pt(grootte + 6)
            for j, c in enumerate(cols):
                cel = ws.cell(row=rr, column=c)
                txt, red = fmt(cel.value, cel.number_format)
                if i == 0:
                    txt = txt.replace("\n", " ")
                dot = cel.number_format == '"●";"●";"○"' and isinstance(cel.value, (int, float))
                tc = tbl.cell(i, j)
                tc.margin_left = tc.margin_right = Pt(3)
                tc.margin_top = tc.margin_bottom = Pt(1)
                tc.vertical_anchor = MSO_ANCHOR.MIDDLE
                tf = tc.text_frame
                tf.word_wrap = True
                p = tf.paragraphs[0]
                p.text = txt
                al = cel.alignment.horizontal
                p.alignment = PP_ALIGN.LEFT if j == 0 or al == "left" else (PP_ALIGN.CENTER if (dot or i == 0 or al == "center") else PP_ALIGN.RIGHT)
                kleur = WIT if i == 0 else NAVY
                if i > 0:
                    rgb = cel.font.color.rgb if cel.font and cel.font.color and isinstance(cel.font.color.rgb, str) else None
                    if rgb and rgb[-6:] not in ("000000",):
                        kleur = RGBColor.from_string(rgb[-6:])
                    if red:
                        kleur = ROOD
                    if dot:
                        kleur = SL[int(cel.value)]
                    if "Δ" in heads[c] and isinstance(cel.value, (int, float)) and "%" not in cel.number_format:
                        kleur = ROOD if cel.value < 0 else (GROEN if cel.value > 0 else kleur)
                for run in p.runs:
                    run.font.name = "Arial"
                    run.font.size = Pt(grootte + 5 if dot else grootte)
                    run.font.bold = True if i == 0 else bool(cel.font.b)
                    run.font.italic = bool(cel.font.i) if i else False
                    run.font.color.rgb = kleur
                tc.fill.solid()
                if i == 0:
                    tc.fill.fore_color.rgb = NAVY
                elif codes.get(c) == sel:
                    tc.fill.fore_color.rgb = GEELLICHT
                elif cel.fill and cel.fill.fill_type == "solid" and cel.fill.start_color.rgb not in (None, "00000000", "FFFFFFFF"):
                    tc.fill.fore_color.rgb = RGBColor.from_string(str(cel.fill.start_color.rgb)[-6:])
                else:
                    tc.fill.fore_color.rgb = WIT

    def grafiek(s, idx, x, y, w, h, t):
        kop(s, t, x, y, w)
        png = chart_png.get(idx)
        if not png:
            return
        iw, ih = Image.open(png).size
        if iw / ih > w / h:
            pw, ph = w, w * ih / iw
        else:
            pw, ph = h * iw / ih, h
        s.shapes.add_picture(str(png), Pt(x + (w - pw) / 2), Pt(y), Pt(pw), Pt(ph))

    def opsomming(s, name, x, y, w, h):
        ws, r = bk.ref(name)
        items = [c.value for row in ws[r] for c in row if c.value not in (None, "")]
        vlak(s, x, y, w, h, LICHT)
        tb = s.shapes.add_textbox(Pt(x + 24), Pt(y + 20), Pt(w - 48), Pt(h - 40))
        tf = tb.text_frame
        tf.word_wrap = True
        for i, t in enumerate(items or ["Geen bijzonderheden."]):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = "■  " + str(t)
            p.space_after = Pt(14)
            for run in p.runs:
                run.font.name, run.font.size = "Arial", Pt(16)
                run.font.color.rgb = NAVY

    spec_ws, spec_r = bk.ref("PPT_Spec")
    vorige, sld = None, None
    for row in spec_ws[spec_r]:
        dia, titel, kern, soort, ref, x, y, w, h, kp, actief = (c.value for c in row)
        if actief != 1:
            continue
        if dia != vorige:
            sld = nieuwe_dia(titel or "", kern or "")
            vorige = dia
        if soort == "kpi":
            tegels(sld, ref, x, y, w, h)
        elif soort == "tabel":
            tabel(sld, ref, x, y, w, kp or "")
        elif soort == "grafiek":
            grafiek(sld, int(ref), x, y, w, h, kp or "")
        elif soort == "opsomming":
            opsomming(sld, ref, x, y, w, h)
    prs.save(out)
    return len(prs.slides)


if __name__ == "__main__":
    print(main(sys.argv[1], sys.argv[2], sys.argv[3]), "dia's")
