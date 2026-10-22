import sys; sys.path.insert(0, sys.path[0])
from hstyle import *
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side, Font
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.chart import BarChart, Reference

SWKC, WBC = PETROL, "C98A00"
wb = Workbook()
v = wb.active; v.title = "Vergelijking"

def name(n, ref): wb.defined_names[n] = DefinedName(n, attr_text=ref)
def ja_nee(ws, cells):
    dv = DataValidation(type="list", formula1='"ja,nee"', allow_blank=False); ws.add_data_validation(dv)
    for c in cells: dv.add(c)
def heads(ws, r, labels, align=None):
    for i, h in enumerate(labels):
        c = ws.cell(r, 2 + i, h); c.font = font(size=9, bold=True, color=NAVY)
        c.alignment = Alignment(horizontal=(align or {}).get(i, "left" if i in (0, len(labels) - 1) else "right"), wrap_text=True)
        c.border = Border(bottom=Side(style="medium", color=NAVY))
def vcenter(ws, r1, r2, c2):
    for row in ws.iter_rows(min_row=r1, max_row=r2, min_col=2, max_col=c2):
        for c in row:
            a = c.alignment; c.alignment = Alignment(horizontal=a.horizontal, vertical="center", indent=a.indent, wrap_text=a.wrap_text)
        ws.row_dimensions[row[0].row].height = ws.row_dimensions[row[0].row].height or 19
def gold_total(ws, r, c2, label, col, formula):
    for cc in range(2, c2 + 1):
        ws.cell(r, cc).fill = fill(GOLD)
        ws.cell(r, cc).border = Border(top=Side(style="medium", color=NAVY), bottom=Side(style="medium", color=NAVY))
    ws.cell(r, 2, label).font = font(size=11, bold=True, color=NAVY)
    c = ws.cell(r, col, formula); c.font = font(size=11, bold=True, color=NAVY); c.number_format = EUR
    ws.row_dimensions[r].height = 24

# =============== Vergelijking ===============
for col, w in zip("BCDEFGH", [40, 16, 16, 20, 3, 36, 42]): v.column_dimensions[col].width = w
brand(v, 8, "Vergelijking SWK  vs.  Woningborg  |  Garantie- en Waarborgregeling",
      "Blauwe cijfers op geel = invoer  ·  Tarieven: 'Tarieven SWK' (rekenmodel 2026) en 'Tarieven Woningborg' (tarievenbrief 2027)")

tile(v, 5, 2, 2, "SWK  (EXCL. BTW)", "=SWK_Excl", EUR, SWKC, SWKC)
tile(v, 5, 3, 4, "WONINGBORG  (EXCL. BTW)", "=WB_Excl", EUR, WBC, WBC)
tile(v, 5, 5, 6, "VERSCHIL", "=ABS(WB_Excl-SWK_Excl)", EUR)
tile(v, 5, 7, 8, "VOORDELIGST", '=IF(SWK_Excl=WB_Excl,"Gelijk",IF(SWK_Excl<WB_Excl,"SWK","Woningborg"))', "General")
v["E8"] = '=IF(MAX(SWK_Excl,WB_Excl)=0,"",TEXT(ABS(WB_Excl-SWK_Excl)/MAX(SWK_Excl,WB_Excl),"0.0%")&" goedkoper")'
v["E8"].font = font(size=9, italic=True, color=MUTED); v["E8"].alignment = Alignment(horizontal="center")
v.conditional_formatting.add("G6", FormulaRule(formula=['G6="SWK"'], font=Font(color=SWKC, bold=True)))
v.conditional_formatting.add("G6", FormulaRule(formula=['G6="Woningborg"'], font=Font(color=WBC, bold=True)))

# Project
bar(v, 10, 2, 4, "PROJECT"); bar(v, 10, 7, 8, "KEUZES")
v["B11"] = "Projectnaam"; v.merge_cells("C11:D11"); inp(v["C11"]); v["C11"] = "Voorbeeldproject"
v["D11"].fill = fill(INPUT); v["D11"].border = BOTTOM
v["B12"] = "Aantal woningen"; inp(v["D12"], "#,##0"); v["D12"] = 1
v["B13"] = "Aanneemsom (excl. btw)"; inp(v["D13"], EUR); v["D13"] = 420000
v["B14"] = "Koopsom grond (excl. btw)"; inp(v["D14"], EUR); v["D14"] = 80000
v["B15"] = '="Btw op koop-/aanneemsom ("&TEXT(C15,"0%")&")"'; inp(v["C15"], "0%", center=True); v["C15"] = 0.21
v["D15"] = "=(D13+D14)*C15"; calc(v["D15"], EUR)
v["B16"] = "Aanneemsom (incl. btw)"; v["D16"] = "=D13*(1+C15)"; calc(v["D16"], EUR, bold=True)
v["B17"] = "Koop- en aanneemsom (incl. btw)"; v["D17"] = "=(D13+D14)*(1+C15)"; calc(v["D17"], EUR, bold=True)
v["B18"] = "Aantal bouwlagen boven straatpeil"; inp(v["D18"], "#,##0"); v["D18"] = 4
for r in (11, 12, 13, 14, 15, 18): v.cell(r, 2).font = font()
for r in (16, 17): v.cell(r, 2).font = font(bold=True, color=NAVY)
name("AantalWoningen", "Vergelijking!$D$12"); name("AanneemsomExcl", "Vergelijking!$D$13")
name("GrondExcl", "Vergelijking!$D$14"); name("BtwKoopsom", "Vergelijking!$C$15")
name("AanneemsomIncl", "Vergelijking!$D$16"); name("KoopAanneemIncl", "Vergelijking!$D$17")
name("Bouwlagen", "Vergelijking!$D$18")

# Keuzes
def sub(r, text, color):
    c = v.cell(r, 7, text); c.font = font(size=9, bold=True, italic=True, color=color)
sub(11, "SWK", SWKC)
v["G12"] = "Verlengde garantie zakelijk"; inp(v["H12"], center=True); v["H12"] = "nee"
v["G13"] = "Prognose bouwtijd > 3 jaar"; inp(v["H13"], center=True); v["H13"] = "nee"
v["G14"] = "Hoogbouwopslag (> 13 bouwlagen)"; v["H14"] = '=IF(Bouwlagen>13,"ja","nee")'
sub(15, "Woningborg", WBC)
v["G16"] = "Regeling / certificaat"; inp(v["H16"]); v["H16"] = "Particulier nieuwbouw – verkoop met grond"
v["G17"] = "Hoogbouwtarief (vanaf 13 lagen)"; v["H17"] = '=IF(Bouwlagen>=13,"ja","nee")'
v["G18"] = "Grondslag premie"; v["H18"] = "=WB_GrondslagSoort"
for r in (12, 13, 14, 16, 17, 18): v.cell(r, 7).font = font()
for c in ("H14", "H17", "H18"): calc(v[c], color=MUTED); v[c].alignment = Alignment(horizontal="center")
ja_nee(v, ["H12", "H13"])
name("KeuzeZakelijk", "Vergelijking!$H$12"); name("KeuzeBouwtijd", "Vergelijking!$H$13")
name("KeuzeHoogbouw", "Vergelijking!$H$14"); name("WB_Regeling", "Vergelijking!$H$16")
name("WB_Hoog", "Vergelijking!$H$17")
dvr = DataValidation(type="list", formula1="='Tarieven Woningborg'!$B$12:$B$19", allow_blank=False); v.add_data_validation(dvr); dvr.add("H16")

# Kostenvergelijking
bar(v, 20, 2, 8, "KOSTENVERGELIJKING")
heads(v, 21, ["Kostenpost", "SWK", "Woningborg", "Verschil (WB − SWK)", "", "Voordeligst", "Toelichting"], {5: "center"})
v["C21"].fill = fill(SWKC); v["D21"].fill = fill(WBC); v["C21"].font = v["D21"].font = font(size=9, bold=True, color="FFFFFF")
posts = [(22, "Provisie / premie", "SWK_Premie", "WB_Premie", "SWK: Sectie I + II  ·  WB: tarief × grondslag regeling"),
         (23, "Opslagen", "SWK_Opslag", "WB_Opslag", "WB: hoogbouw zit in het tarief zelf"),
         (24, "Overige / bijkomende kosten", "SWK_Overig", "WB_Overig", "SWK: bureaukosten  ·  WB: bezoeken, metingen, diensten")]
for r, lab, s_, w_, tip in posts:
    v.cell(r, 2, lab).font = font()
    calc(v.cell(r, 3, f"={s_}"), EUR, color="008000"); calc(v.cell(r, 4, f"={w_}"), EUR, color="008000")
    calc(v.cell(r, 5, f"=D{r}-C{r}"), EUR, bg=None)
    c = v.cell(r, 7, f'=IF(C{r}=D{r},"–",IF(C{r}<D{r},"SWK","Woningborg"))'); c.font = font(size=9, bold=True); c.alignment = Alignment(horizontal="center")
    note(v.cell(r, 8), tip)
gold_total(v, 26, 8, "Totaal excl. btw", 3, "=SUM(C22:C24)")
for col, f in (("D", "=SUM(D22:D24)"), ("E", "=D26-C26")):
    v[f"{col}26"] = f; v[f"{col}26"].font = font(size=11, bold=True, color=NAVY); v[f"{col}26"].number_format = EUR
v["G26"] = '=IF(C26=D26,"–",IF(C26<D26,"SWK","Woningborg"))'; v["G26"].font = font(size=11, bold=True); v["G26"].alignment = Alignment(horizontal="center")
note(v["H26"], "Relevante vergelijking als de btw aftrekbaar is")
rows = [(27, "Btw over tarieven", "=SWK_Btw", "=WB_Btw", EUR, "Percentages: zie de tariefbladen"),
        (28, "Totaal incl. btw", "=SWK_Incl", "=WB_Incl", EUR, ""),
        (29, "Kosten per woning (excl. btw)", "=IFERROR(C26/AantalWoningen,0)", "=IFERROR(D26/AantalWoningen,0)", EUR, ""),
        (30, "In % van aanneemsom incl. btw", "=IFERROR(C26/AanneemsomIncl,0)", "=IFERROR(D26/AanneemsomIncl,0)", "0.00%", "Zelfde noemer voor beide, zodat ze vergelijkbaar zijn")]
for r, lab, fs, fw, fmt, tip in rows:
    b = r == 28
    v.cell(r, 2, lab).font = font(bold=b, color=NAVY if b else "000000")
    calc(v.cell(r, 3, fs), fmt, bold=b); calc(v.cell(r, 4, fw), fmt, bold=b)
    calc(v.cell(r, 5, f"=D{r}-C{r}"), fmt, bold=b, bg=None)
    note(v.cell(r, 8), tip)
v["G28"] = '=IF(C28=D28,"–",IF(C28<D28,"SWK","Woningborg"))'; v["G28"].font = font(bold=True); v["G28"].alignment = Alignment(horizontal="center")
for rng, f0 in (("G22:G28", "G22"),):
    v.conditional_formatting.add(rng, FormulaRule(formula=[f'{f0}="SWK"'], font=Font(color="FFFFFF", bold=True), fill=fill(SWKC)))
    v.conditional_formatting.add(rng, FormulaRule(formula=[f'{f0}="Woningborg"'], font=Font(color="FFFFFF", bold=True), fill=fill(WBC)))
v.conditional_formatting.add("E22:E30", FormulaRule(formula=["E22<0"], font=Font(color="38761D")))
note(v["B31"], "Positief verschil = Woningborg duurder; negatief = Woningborg goedkoper.")

# Controles / aandachtspunten
bar(v, 33, 2, 4, "CONTROLES"); bar(v, 33, 7, 8, "LET OP: VERSCHILLEN IN REKENWIJZE")
checks = [(34, "Invoer compleet", '=IF(AND(D13>0,D12>0,D18>0),"OK","AANVULLEN")'),
          (35, "Keuzes geldig", '=IF(AND(OR(H12="ja",H12="nee"),OR(H13="ja",H13="nee"),ISNUMBER(MATCH(H16,\'Tarieven Woningborg\'!$B$12:$B$19,0))),"OK","CONTROLEREN")'),
          (36, "Tarieven SWK bevestigd", "='Tarieven SWK'!H5"),
          (37, "Tarieven Woningborg bevestigd", "='Tarieven Woningborg'!H5"),
          (38, "Koopsom grond ingevuld (bij 'met grond')", '=IF(AND(WB_GrondslagSoort="Koop- en aanneemsom incl. btw",D14<=0),"CONTROLEREN","OK")')]
for r, lab, f in checks:
    v.cell(r, 2, lab).font = font(size=9, color=MUTED)
    c = v.cell(r, 4, f); c.font = font(size=9, bold=True, color=NAVY); c.alignment = Alignment(horizontal="center")
    for cc in range(2, 5): v.cell(r, cc).border = BOTTOM
v.conditional_formatting.add("D34:D38", FormulaRule(formula=['D34="OK"'], fill=fill(OKF)))
v.conditional_formatting.add("D34:D38", FormulaRule(formula=['D34<>"OK"'], fill=fill(BADF)))
pts = ["• SWK: twee provisies + opslagen over de aanneemsom. WB: één tarief over de koop- én aanneemsom.",
       "• WB rekent btw over de premie; het SWK-rekenmodel noemt geen btw (staat op 0%, controleren).",
       "• Hoogbouw: SWK-opslag bij > 13 lagen, WB-hoogbouwtarief vanaf 13 lagen.",
       "• Tariefjaar verschilt: SWK 2026 (rekenmodel), Woningborg 2027 (tarievenbrief).",
       "• WB-bouwplaatsbezoeken/geluidsmetingen alleen bij transformatie en houtbouw (op offerte)."]
for i, t in enumerate(pts):
    v.merge_cells(start_row=34 + i, start_column=7, end_row=34 + i, end_column=8); note(v.cell(34 + i, 7), t)

ch = BarChart(); ch.type = "bar"; ch.grouping = "clustered"; ch.title = "Kosten per post (excl. btw)"
ch.add_data(Reference(v, min_col=3, max_col=4, min_row=21, max_row=24), titles_from_data=True)
ch.set_categories(Reference(v, min_col=2, min_row=22, max_row=24))
for s, col in zip(ch.series, (SWKC, WBC)):
    s.graphicalProperties.solidFill = col; s.graphicalProperties.line.solidFill = col
ch.x_axis.scaling.orientation = "maxMin"; ch.y_axis.numFmt = '€ #,##0'; ch.y_axis.majorGridlines = None
ch.legend.position = "b"; ch.height = 7; ch.width = 24
v.add_chart(ch, "B41")
vcenter(v, 4, 38, 8)
v.freeze_panes = "A10"
print_setup(v, "A1:H56", "Heijmans Vastgoed | Vergelijking SWK – Woningborg"); v.page_setup.fitToHeight = 1
v.sheet_properties.tabColor = GOLD

# =============== Tarieven SWK ===============
COND = '"Altijd,Hoogbouw,Zakelijk,Bouwtijd > 3 jaar,Nooit"'
def active(ref):
    return (f'IF({ref}="Altijd",1,IF({ref}="Hoogbouw",--(KeuzeHoogbouw="ja"),IF({ref}="Zakelijk",--(KeuzeZakelijk="ja"),'
            f'IF({ref}="Bouwtijd > 3 jaar",--(KeuzeBouwtijd="ja"),0))))')

def header_block(s, title, color, src, ok, btw, btwnote, c2=9):
    s.sheet_properties.tabColor = color
    brand(s, c2, f"Tarieven {title}  |  Garantie- en Waarborgregeling",
          "Blauwe cijfers op geel = tarief/keuze aanpassen  ·  Lichtblauw = berekening (rekent met de invoer op 'Vergelijking')")
    s["B5"] = "Bron / tarieflijst"; inp(s["C5"]); s["C5"] = src; s.merge_cells("C5:F5")
    for cc in range(4, 7): s.cell(5, cc).fill = fill(INPUT); s.cell(5, cc).border = BOTTOM
    s["G5"] = "Bevestigd?"; s["G5"].font = font(bold=True); s["G5"].alignment = Alignment(horizontal="right")
    s["I5"] = ok; inp(s["I5"], center=True)
    s["H5"] = '=IF(I5="ja","OK","CONTROLEREN")'; calc(s["H5"]); s["H5"].alignment = Alignment(horizontal="center")
    ja_nee(s, ["I5"])
    s.conditional_formatting.add("H5", FormulaRule(formula=['H5="OK"'], fill=fill(OKF)))
    s.conditional_formatting.add("H5", FormulaRule(formula=['H5<>"OK"'], fill=fill(BADF)))
    s["B6"] = "Btw over tarieven"; inp(s["C6"], "0%", center=True); s["C6"] = btw
    note(s["D6"], btwnote)

s = wb.create_sheet("Tarieven SWK")
for col, w in zip("BCDEFGHI", [36, 12, 14, 22, 20, 15, 15, 48]): s.column_dimensions[col].width = w
header_block(s, "SWK", SWKC, "SWK Rekenmodel 2026 – Berekening Garantie- en Waarborgregeling", "ja", 0,
             "Het SWK-rekenmodel noemt geen btw over de provisie; controleren en zo nodig op 21% zetten")
bar(s, 9, 2, 9, "BASISPROVISIES", SWKC)
heads(s, 10, ["Kostenpost", "Tarief", "Grondslag %", "", "Grondslag (€)", "Bedrag (€)", "", "Toelichting"])
s["B11"] = "(grondslag = % van de aanneemsom incl. btw)"; s["B11"].font = font(size=8, italic=True, color=MUTED)
for i, (lab, pct, factor, tip) in enumerate([("Provisie Sectie I (Afbouw)", 0.005198, 0.2, "Over 20% van de aanneemsom incl. btw"),
                                             ("Provisie Sectie II (Kwaliteit)", 0.00038, 1, "Over de aanneemsom incl. btw")]):
    r = 12 + i
    s.cell(r, 2, lab).font = font()
    inp(s.cell(r, 3, pct), PCT4); inp(s.cell(r, 4, factor), "0%")
    calc(s.cell(r, 6, f"=AanneemsomIncl*D{r}"), EUR, bg=None, color=MUTED)
    calc(s.cell(r, 7, f"=F{r}*C{r}"), EUR)
    note(s.cell(r, 9), tip)
bar(s, 15, 2, 9, "OPSLAGEN", SWKC)
heads(s, 16, ["Omschrijving", "Tarief", "Berekend over", "Van toepassing bij", "Grondslag (€)", "Actief", "Bedrag (€)", "Toelichting"])
dvo = DataValidation(type="list", formula1='"Sectie I,Sectie II,Aanneemsom incl. btw"'); s.add_data_validation(dvo)
dvc = DataValidation(type="list", formula1=COND); s.add_data_validation(dvc)
opslag = [("Opslag hoogbouw", 0.303, "Sectie II", "Hoogbouw", "> 13 bouwlagen (volgt uit aantal bouwlagen)"),
          ("Opslag verlengde garantie zakelijk", 0.1, "Sectie II", "Zakelijk", "Alleen bij zakelijke opdrachtgever"),
          ("Opslag prognose bouwtijd > 3 jaar", 0.1, "Sectie I", "Bouwtijd > 3 jaar", "Start bouw tot 1e oplevering")]
for i in range(6):
    r = 17 + i
    lab, pct, over, cond, tip = opslag[i] if i < len(opslag) else (None,) * 5
    inp(s.cell(r, 2, lab)); s.cell(r, 2).font = font(color="0000FF")
    inp(s.cell(r, 3, pct), "0.0%"); inp(s.cell(r, 4, over)); inp(s.cell(r, 5, cond))
    dvo.add(f"D{r}"); dvc.add(f"E{r}")
    calc(s.cell(r, 6, f'=IF(D{r}="Sectie I",$G$12,IF(D{r}="Sectie II",$G$13,IF(D{r}="Aanneemsom incl. btw",AanneemsomIncl,0)))'), EUR, bg=None, color=MUTED)
    c = s.cell(r, 7, f'=IF(B{r}="","",IF({active(f"E{r}")}=1,"ja","nee"))'); calc(c, bg=None, color=MUTED); c.alignment = Alignment(horizontal="center")
    calc(s.cell(r, 8, f'=IF(G{r}="ja",F{r}*C{r},0)'), EUR)
    note(s.cell(r, 9), tip or "")
s["G23"] = "Totaal"; s["G23"].font = font(bold=True); s["G23"].alignment = Alignment(horizontal="right")
s["H23"] = "=SUM(H17:H22)"; calc(s["H23"], EUR, bold=True)
bar(s, 25, 2, 9, "OVERIGE KOSTEN", SWKC)
heads(s, 26, ["Omschrijving", "Bedrag", "Per", "Van toepassing bij", "Aantal", "Actief", "Bedrag (€)", "Toelichting"])
dvp = DataValidation(type="list", formula1='"woning,project"'); s.add_data_validation(dvp)
for i in range(6):
    r = 27 + i
    lab, bed, per, cond, tip = ("Bureaukosten, voorschot", 750, "woning", "Altijd", "Voorschot per woning") if i == 0 else (None,) * 5
    inp(s.cell(r, 2, lab)); s.cell(r, 2).font = font(color="0000FF")
    inp(s.cell(r, 3, bed), EUR); inp(s.cell(r, 4, per)); inp(s.cell(r, 5, cond))
    dvp.add(f"D{r}"); dvc.add(f"E{r}")
    calc(s.cell(r, 6, f'=IF(D{r}="woning",AantalWoningen,IF(D{r}="project",1,0))'), '#,##0', bg=None, color=MUTED)
    c = s.cell(r, 7, f'=IF(B{r}="","",IF({active(f"E{r}")}=1,"ja","nee"))'); calc(c, bg=None, color=MUTED); c.alignment = Alignment(horizontal="center")
    calc(s.cell(r, 8, f'=IF(G{r}="ja",F{r}*C{r},0)'), EUR)
    note(s.cell(r, 9), tip or "")
s["G33"] = "Totaal"; s["G33"].font = font(bold=True); s["G33"].alignment = Alignment(horizontal="right")
s["H33"] = "=SUM(H27:H32)"; calc(s["H33"], EUR, bold=True)
gold_total(s, 35, 9, "Totaal kosten SWK (excl. btw)", 8, "=G12+G13+H23+H33")
s["B36"] = "Btw over tarieven"; s["H36"] = "=H35*C6"; calc(s["H36"], EUR)
s["B37"] = "Totaal kosten SWK (incl. btw)"; s["B37"].font = font(bold=True, color=NAVY); s["H37"] = "=H35+H36"; calc(s["H37"], EUR, bold=True)
for n, ref in (("SWK_Premie", "($G$12+$G$13)"), ("SWK_Opslag", "$H$23"), ("SWK_Overig", "$H$33"),
               ("SWK_Excl", "$H$35"), ("SWK_Btw", "$H$36"), ("SWK_Incl", "$H$37")):
    name(n, f"'Tarieven SWK'!{ref}" if not ref.startswith("(") else f"('Tarieven SWK'!$G$12+'Tarieven SWK'!$G$13)")
vcenter(s, 4, 37, 9); s.freeze_panes = "A9"
print_setup(s, "A1:I37", "Heijmans Vastgoed | Tarieven SWK"); s.page_setup.fitToHeight = 1

# =============== Tarieven Woningborg ===============
w = wb.create_sheet("Tarieven Woningborg")
for col, wd in zip("BCDEFGHI", [46, 14, 14, 30, 16, 14, 16, 40]): w.column_dimensions[col].width = wd
header_block(w, "Woningborg", WBC, "Tarievenbrief Woningborg Verzekeren 2027 (18-09-2026), Heijmans Vastgoed – inschrijfnr. 4889", "ja", 0.21,
             "Volgens de tarievenbrief: 'Over alle tarieven wordt btw in rekening gebracht'")
w.merge_cells("B7:I7"); w["B7"] = ("Geldig voor plannen geaccepteerd op of na 1 januari 2027. Particulier: grondslag incl. meer-/minderwerk en btw. "
                                   "Zakelijk: grondslag excl. btw.")
w["B7"].font = font(size=9, italic=True, color=MUTED)
bar(w, 9, 2, 9, "TARIEFMATRIX 2027  (% van de grondslag)", WBC)
heads(w, 10, ["Regeling / certificaat", "Tot 13 lagen", "Vanaf 13 lagen", "Grondslag", "", "", "", "Toelichting"])
w["B11"] = "(kies de regeling op het tabblad 'Vergelijking')"; w["B11"].font = font(size=8, italic=True, color=MUTED)
GR = ["Koop- en aanneemsom incl. btw", "Aanneemsom incl. btw", "Koop- en aanneemsom excl. btw", "Aanneemsom excl. btw"]
matrix = [("Particulier nieuwbouw – verkoop met grond", 0.00227, 0.0025, GR[0], "Ook bij afgekochte erfpachtcanon"),
          ("Particulier nieuwbouw – verkoop zonder grond", 0.00321, 0.00354, GR[1], "Grond door derde verkocht / erfpacht zonder afkoop"),
          ("Kavelbouw", 0.0038, 0.0038, GR[1], ""),
          ("Transformatie particulier", 0.00321, 0.00321, GR[0], "Ook zonder grond over koop- en aanneemsom; + bijkomende kosten"),
          ("Zakelijk nieuwbouw – insolventie + herstel", 0.00568, 0.00757, GR[3], ""),
          ("Zakelijk nieuwbouw – alleen insolventie", 0.00343, 0.00458, GR[3], ""),
          ("Zakelijk transformatie – insolventie + herstel", 0.00903, 0.00903, GR[2], "Gebouw valt WEL onder dekking; + bijkomende kosten"),
          ("Zakelijk transformatie – alleen insolventie", 0.00343, 0.00343, GR[3], "Gebouw valt NIET onder dekking")]
dvg = DataValidation(type="list", formula1='"' + ",".join(GR) + '"'); w.add_data_validation(dvg)
for i, (lab, a, b, g, tip) in enumerate(matrix):
    r = 12 + i
    w.cell(r, 2, lab).font = font()
    inp(w.cell(r, 3, a), "0.000%"); inp(w.cell(r, 4, b), "0.000%"); inp(w.cell(r, 5, g)); w.cell(r, 5).alignment = Alignment(indent=1); dvg.add(f"E{r}")
    note(w.cell(r, 9), tip)
w.conditional_formatting.add("B12:E19", FormulaRule(formula=["$B12=WB_Regeling"], font=Font(bold=True), fill=fill("FDE9B8")))

bar(w, 21, 2, 9, "PREMIEBEREKENING", WBC)
calc_rows = [(22, "Gekozen regeling", "=WB_Regeling", "General"),
             (23, "Hoogbouwtarief (vanaf 13 lagen)", "=WB_Hoog", "General"),
             (24, "Tarief", '=IFERROR(IF(WB_Hoog="ja",INDEX($D$12:$D$19,MATCH(WB_Regeling,$B$12:$B$19,0)),INDEX($C$12:$C$19,MATCH(WB_Regeling,$B$12:$B$19,0))),0)', "0.000%"),
             (25, "Soort grondslag", '=IFERROR(INDEX($E$12:$E$19,MATCH(WB_Regeling,$B$12:$B$19,0)),"")', "General"),
             (26, "Grondslag (€)", f'=IF(D25="{GR[0]}",KoopAanneemIncl,IF(D25="{GR[1]}",AanneemsomIncl,IF(D25="{GR[2]}",AanneemsomExcl+GrondExcl,IF(D25="{GR[3]}",AanneemsomExcl,0))))', EUR)]
for r, lab, f, fmt in calc_rows:
    w.cell(r, 2, lab).font = font()
    w.merge_cells(start_row=r, start_column=4, end_row=r, end_column=6)
    c = w.cell(r, 4, f); calc(c, fmt, color=MUTED, bg=None); c.alignment = Alignment(horizontal="right")
w["B27"] = "Premie"; w["B27"].font = font(bold=True)
w["H27"] = "=D26*D24"; calc(w["H27"], EUR, bold=True)
name("WB_GrondslagSoort", "'Tarieven Woningborg'!$D$25")

bar(w, 29, 2, 9, "BIJKOMENDE KOSTEN EN EXTRA DIENSTVERLENING", WBC)
heads(w, 30, ["Omschrijving", "Bedrag", "Aantal", "Wanneer", "", "", "Totaal (€)", "Toelichting"])
extra = [("Bouwplaatsbezoek (per bezoek)", 1041, 0, "Transformatie / houtbouw", "Op offerte na eerste opname"),
         ("Geluidsmeting (per meting)", 911, 0, "Transformatie", "Op offerte na eerste opname"),
         ("Opstellen depotovereenkomst", 1500, 0, "Op verzoek", "Eenmalig"),
         ("Opstellen afbouwverklaring", 500, 0, "Op verzoek", "Eenmalig"),
         ("Laten vervallen van plan na risicobeoordeling", 2500, 0, "Op verzoek", "Eenmalig"),
         ("Planacceptatie bij < 70% verkocht", 1500, 0, "Op verzoek", "Eenmalig"),
         ("Aanpassing standaard termijnregeling", 1250, 0, "Op verzoek", "Eenmalig"),
         ("Jaarlijkse bijdrage (per onderneming)", 644, 0, "Per jaar", "Niet projectgebonden; standaard niet meegeteld"),
         ("Inschrijfgeld (per onderneming)", 4018, 0, "Eenmalig", "Binnen concern € 1.133 per extra inschrijving")]
for i, (lab, bed, n, when, tip) in enumerate(extra):
    r = 31 + i
    w.cell(r, 2, lab).font = font()
    inp(w.cell(r, 3, bed), EUR); inp(w.cell(r, 4, n), "#,##0", center=True)
    note(w.cell(r, 5), when)
    calc(w.cell(r, 8, f"=C{r}*D{r}"), EUR)
    note(w.cell(r, 9), tip)
w["G40"] = "Totaal"; w["G40"].font = font(bold=True); w["G40"].alignment = Alignment(horizontal="right")
w["H40"] = "=SUM(H31:H39)"; calc(w["H40"], EUR, bold=True)
note(w["B41"], "Alle bedragen excl. btw. Vul bij 'Aantal' in wat voor dit project verwacht wordt.")

gold_total(w, 43, 9, "Totaal kosten Woningborg (excl. btw)", 8, "=H27+H40")
w["B44"] = "Btw over tarieven"; w["H44"] = "=H43*C6"; calc(w["H44"], EUR)
w["B45"] = "Totaal kosten Woningborg (incl. btw)"; w["B45"].font = font(bold=True, color=NAVY); w["H45"] = "=H43+H44"; calc(w["H45"], EUR, bold=True)
for n, ref in (("WB_Premie", "$H$27"), ("WB_Overig", "$H$40"), ("WB_Excl", "$H$43"), ("WB_Btw", "$H$44"), ("WB_Incl", "$H$45")):
    name(n, f"'Tarieven Woningborg'!{ref}")
wb.defined_names["WB_Opslag"] = DefinedName("WB_Opslag", attr_text="0")
vcenter(w, 4, 45, 9); w.freeze_panes = "A9"
print_setup(w, "A1:I45", "Heijmans Vastgoed | Tarieven Woningborg 2027"); w.page_setup.fitToHeight = 1

# =============== Toelichting ===============
t = wb.create_sheet("Toelichting")
for col, wd in zip("BC", [4, 120]): t.column_dimensions[col].width = wd
brand(t, 3, "Toelichting", "Hoe werkt deze vergelijking?")
lines = [("1", "Vul op 'Vergelijking' het project in: aantal woningen, aanneemsom en koopsom grond (excl. btw), btw-% en aantal bouwlagen."),
         ("2", "Kies de SWK-keuzes (zakelijk, bouwtijd) en de Woningborg-regeling. Hoogbouw volgt automatisch uit het aantal bouwlagen."),
         ("3", "Vul op 'Tarieven Woningborg' bij 'Aantal' de verwachte bijkomende kosten in (bezoeken, metingen, extra diensten)."),
         ("", ""),
         ("", "Verschil in rekenwijze"),
         ("•", "SWK: provisie Sectie I (0,5198% over 20% van de aanneemsom) + Sectie II (0,0380% over de aanneemsom) + opslagen + € 750 bureaukosten per woning."),
         ("•", "Woningborg: één tarief per regeling over de koop- en aanneemsom (met grond) of de aanneemsom (zonder grond / kavelbouw). Hoogbouw (≥ 13 lagen) = hoger tarief, geen opslag."),
         ("•", "Woningborg particulier rekent over de som incl. btw en meer-/minderwerk, zakelijk over de som excl. btw. Over de premie zelf komt btw."),
         ("•", "De SWK-provisie staat zonder btw (0%), zoals in het SWK-rekenmodel. Controleer of SWK btw rekent en pas het percentage op 'Tarieven SWK' zo nodig aan."),
         ("•", "Tariefjaren: SWK 2026 (rekenmodel), Woningborg 2027 (tarievenbrief 18-09-2026). Werk de SWK-tarieven bij zodra die voor 2027 bekend zijn."),
         ("•", "Niet meegenomen: houtbouwtoeslag van Woningborg (na technische beoordeling) en SWK-beoordelingen (transformatie, hoogbouw > 25 lagen, afwijkende termijnregeling).")]
for i, (a, b) in enumerate(lines):
    t.cell(5 + i, 2, a).font = font(bold=True, color=NAVY)
    t.cell(5 + i, 3, b).font = font(bold=(a == "" and b != ""), color=NAVY if a == "" else "000000")
    t.row_dimensions[5 + i].height = 20
wb.save(sys.argv[1])
