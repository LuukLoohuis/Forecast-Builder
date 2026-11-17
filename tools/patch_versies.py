"""Versies uit geplakte exports: per versie (Budget, Q1-Q4) een plakblad ZPMW03 en Woningetalage.
Rekenblad VERSIE_REKEN zet ze om naar de lange tabel die de rapporten lezen (L_-namen).
usage: python3 -I patch_versies.py in.xlsm out.xlsm"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from xlsxlib import Book, skeleton, q
from openpyxl.utils import get_column_letter as CL, column_index_from_string as CI

VERS = ['Budget', 'Q1', 'Q2', 'Q3', 'Q4']
VROW = {v: 6 + i for i, v in enumerate(VERS)}          # rij op tabblad Versies
NP = 100                                              # projecten (V_WBS)
RNG = '$A$1:$GZ$600'
ZKEYS = [('BLJ', '"Bedrijfsopbrengsten LJ"'), ('RLJ', '"Bruto (omzet) result. LJ"'), ('KLJ', '"Kosten LJ"'),
         ('BX', '"OP.BOB "&{X}'), ('RX', '"OP.BOR "&{X}'), ('KX', '"Kosten "&{X}'),
         ('BX1', '"OP.BOB "&({X}+1)'), ('RX1', '"OP.BOR "&({X}+1)'), ('KX1', '"Kosten "&({X}+1)'),
         ('BX2', '"OP.BOB "&({X}+2)'), ('RX2', '"OP.BOR "&({X}+2)'), ('KX2', '"Kosten "&({X}+2)'),
         ('KQ1', '"Kosten "&{E}&"-Q1"'), ('KQ2', '"Kosten "&{E}&"-Q2"'), ('KQ3', '"Kosten "&{E}&"-Q3"'), ('KQ4', '"Kosten "&{E}&"-Q4"')]
EKEYS = [('HJ', '"Huidig jaar"'), ('VJR', '"Volgend jaar"'), ('TMV', '"t/m vorig jr Totaal"')] + \
        [(f'M{m:02d}', f'"hjr {m:02d}"') for m in range(1, 13)] + [(f'VQ{n}', f'"Volgend jr Q{n}"') for n in range(1, 5)]
SLOTS = [('Omzet FY', 0, 'd0_omz'), ('BOR FY', 0, 'd0_bor'), ('Omzet actual YTD', 0, 'omz_ytd'), ('BOR actual YTD', 0, 'bor_ytd'),
         ('AK-dekking FY', 0, 'd0_ak'), ('Woningverkopen FY', 0, 'w_fy0'), ('Woningen verkocht t/m vorig jaar', 0, 'w_tmvj'),
         ('Woningverkopen Q1', 0, 'w_q1'), ('Woningverkopen Q2', 0, 'w_q2'), ('Woningverkopen Q3', 0, 'w_q3'), ('Woningverkopen Q4', 0, 'w_q4'),
         ('Woningverkopen actual YTD', 0, 'w_ytd'), ('Omzet FY', 1, 'd1_omz'), ('BOR FY', 1, 'd1_bor'), ('AK-dekking FY', 1, 'd1_ak'),
         ('Woningverkopen FY', 1, 'w_fy1')]
NS = len(SLOTS)
L0 = 6; LN = L0 + len(VERS) * NP * NS - 1            # lange tabel rijen
W0 = 6                                                # brede hulptabel rijen 6..505

def zs(v): return f'{v} ZPMW03'
def es(v): return f'{v} Woningetalage'

def build(bk, st):
    # ---------- plakbladen
    for v in VERS:
        for nm, txt in ((zs(v), f'Plak hier de ZPMW03-export van {v} (hele blad, vanaf A1).'),
                        (es(v), f'Plak hier de Woningetalage-export van {v} (hele blad, vanaf A1).')):
            s = bk.add_sheet(nm, skeleton(tab='FFBFBFBF'))
            s.set('A1', text=txt)
    # ---------- Versies (zichtbaar)
    V = bk.add_sheet('Versies', skeleton(
        cols='<cols><col min="1" max="1" width="2" customWidth="1"/><col min="2" max="2" width="12" customWidth="1"/>'
             '<col min="3" max="4" width="13" customWidth="1"/><col min="5" max="7" width="10" customWidth="1"/>'
             '<col min="8" max="9" width="14" customWidth="1"/><col min="10" max="10" width="16" customWidth="1"/>'
             '<col min="11" max="14" width="13" customWidth="1"/><col min="15" max="15" width="40" customWidth="1"/></cols>',
        tab='FF2E75B6', hidden_grid=True))
    V.set('B1', text='Versies: Budget, Q1, Q2, Q3, Q4', style=st['title'])
    V.set('B2', style=st['text'], text="Plak per versie de ZPMW03 en de Woningetalage in de grijze tabbladen achteraan (bijv. 'Q1 ZPMW03'). Klaar: de RDFO, Projecten en Woningen tonen de versie vanzelf.")
    V.set('B3', style=st['text'], text="Budget = de prognose die vorig jaar is afgegeven voor dit jaar (bijv. de export van Q3 2025 voor budget 2026). Peildatum leeg = standaard.")
    hdr = ['Versie', 'Peildatum export', 'Gebruikt', 'Exportjaar', 'Versiejaar', 'Verschuiving', 'Peilmaand',
           'Projecten ZPMW03', 'Projecten Woningetalage', 'AK-dekking officieel (EUR, optioneel)', 'Omzet FY (EUR)', 'BOR FY (EUR)', 'AK-dekking FY (EUR)', 'Opmerking']
    for i, h in enumerate(hdr): V.set(f'{CL(2+i)}5', text=h, style=st['hdr'])
    for i, v in enumerate(VERS):
        r = VROW[v]
        V.set(f'B{r}', text=v, style=st['label'])
        V.set(f'C{r}', style=st['inp_date'])
        dflt = 'DATE(P_VJ-1,9,30)' if v == 'Budget' else f'DATE(P_VJ,{3*i+1},0)'
        V.set(f'D{r}', formula=f'IF(ISNUMBER(C{r}),C{r},{dflt})', style=st['date'])
        V.set(f'E{r}', formula=f'YEAR(D{r})', style=st['cell'])
        V.set(f'F{r}', formula=f'E{r}+{1 if v=="Budget" else 0}', style=st['cell'])
        V.set(f'G{r}', formula=f'F{r}-E{r}', style=st['cell'])
        V.set(f'H{r}', formula=f'MONTH(D{r})', style=st['cell'])
        a, b = W0 + i*NP, W0 + i*NP + NP - 1
        V.set(f'I{r}', formula=f'IF(VERSIE_REKEN!$CE${r}=0,"leeg",COUNTIF(VERSIE_REKEN!$AA${a}:$AA${b},">0"))', style=st['cell'])
        V.set(f'J{r}', formula=f'IF(VERSIE_REKEN!$CF${r}=0,"leeg",COUNTIF(VERSIE_REKEN!$AB${a}:$AB${b},">0"))', style=st['cell'])
        V.set(f'K{r}', style=st['inp_num'])
        for col, k in (('L', 'Omzet FY'), ('M', 'BOR FY'), ('N', 'AK-dekking FY')):
            V.set(f'{col}{r}', style=st['num'], formula=
                  f'IF(VERSIE_REKEN!$CE${r}=0,"",SUMIFS(L_Waarde,L_Versie,$B{r},L_Versiejaar,$F{r},L_Doeljaar,$F{r},L_Kengetal,"{k}"))')
        V.set(f'O{r}', style=st['text'], is_str=True, formula=
              f'IF(AND(VERSIE_REKEN!$CE${r}=0,VERSIE_REKEN!$CF${r}=0),"",IF(VERSIE_REKEN!$CE${r}=0,"ZPMW03 ontbreekt",IF(VERSIE_REKEN!$CG${r}=0,"kolom WBS-element niet gevonden",IF($F{r}<>P_VJ,"versiejaar "&$F{r}&" (rapport toont "&P_VJ&")",IF(K{r}<>"","AK: officieel bedrag telt","")))))')
    V.set('B12', style=st['text'], text="AK-dekking per versie = kosten / 1,05 x AK %. Vul 'AK-dekking officieel' in om het gerapporteerde bedrag te tonen.")
    V.set('B13', style=st['text'], text="Projecten die niet op de Projectlijst staan of op Nee staan, tellen niet mee. Zie Handleiding.")
    V.root.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}dimension').set('ref', 'A1:O13')

    # ---------- VERSIE_REKEN (verborgen)
    R = bk.add_sheet('VERSIE_REKEN', skeleton(), state='hidden')
    R.set('A1', text='VERSIE_REKEN: versies uit de plakbladen (niet wijzigen). Lange tabel A:V = L_-namen; X:CD per versie x project; CE: kolomnummers per versie.')
    for i, h in enumerate(['ImportID','Bronbestand','Importdatum','Status','Versiejaar','Versie','Revisie','Peildatum','Doeljaar','WBS',
                           'Projectdefinitie','Projectnaam','Type','IFRS','Kengetal','Waarde','Eenheid','Toelichting','Sleutel','MaxRev','Actief','Dubbel']):
        R.set(f'{CL(1+i)}5', text=h)
    # kolomnummers per versie (rij 6..10)
    zcol = {k: CL(CI('CH') + j) for j, (k, _) in enumerate(ZKEYS)}
    ecol = {k: CL(CI('CY') + j) for j, (k, _) in enumerate(EKEYS)}
    R.set('CE5', text='hdrZ'); R.set('CF5', text='hdrE'); R.set('CG5', text='WBS'); R.set('CX5', text='Stufe3')
    for k, c in zcol.items(): R.set(f'{c}5', text=k)
    for k, c in ecol.items(): R.set(f'{c}5', text=k)
    for v in VERS:
        r = VROW[v]; X = f'Versies!$E${r}'; E = f'Versies!$F${r}'
        Z, Wn = q(zs(v)), q(es(v))
        R.set(f'CE{r}', formula=''.join(f'IF(COUNTIF({Z}!${k}:${k},"WBS-element")>0,{k},' for k in range(1, 16)) + '0' + ')'*15)
        R.set(f'CF{r}', formula=''.join(f'IF(COUNTIF({Wn}!${k}:${k},"Stufe 3")>0,{k},' for k in range(1, 16)) + '0' + ')'*15)
        look = lambda sh, hdrc, name: f'IF(${hdrc}${r}=0,0,IFERROR(MATCH({name},INDEX({sh}!$A$1:$GZ$15,${hdrc}${r},0),0),0))'
        R.set(f'CG{r}', formula=look(Z, 'CE', '"WBS-element"'))
        R.set(f'CX{r}', formula=look(Wn, 'CF', '"Stufe 3"'))
        for k, expr in ZKEYS: R.set(f'{zcol[k]}{r}', formula=look(Z, 'CE', expr.format(X=X, E=E)))
        for k, expr in EKEYS: R.set(f'{ecol[k]}{r}', formula=look(Wn, 'CF', expr.format(X=X, E=E)))
    # brede tabel
    zraw = {k: CL(CI('AC') + j) for j, (k, _) in enumerate(ZKEYS)}        # AC..AR
    eraw = {k: CL(CI('AS') + j) for j, (k, _) in enumerate(EKEYS)}        # AS..BK
    dcol = {key: CL(CI('BM') + j) for j, (_, _, key) in enumerate(SLOTS)} # BM..CB
    for vi, v in enumerate(VERS):
        r = VROW[v]; Z, Wn = q(zs(v)), q(es(v)); off = f'Versies!$G${r}'; pm = f'Versies!$H${r}'
        for p in range(1, NP + 1):
            w = W0 + vi*NP + p - 1
            R.set(f'X{w}', text=v); R.set(f'Y{w}', formula=str(p))
            R.set(f'Z{w}', is_str=True, formula=f'IF(INDEX(V_WBS,{p})="","",INDEX(V_WBS,{p})&"")')
            R.set(f'AA{w}', formula=f'IF(OR($Z{w}="",$CG${r}=0),0,COUNTIF(INDEX({Z}!{RNG},0,$CG${r}),$Z{w}))')
            R.set(f'AB{w}', formula=f'IF(OR($Z{w}="",$CX${r}=0),0,COUNTIF(INDEX({Wn}!{RNG},0,$CX${r}),$Z{w}))')
            for k, c in zraw.items():
                R.set(f'{c}{w}', formula=f'IF(OR($AA{w}=0,${zcol[k]}${r}=0),"",SUMIFS(INDEX({Z}!{RNG},0,${zcol[k]}${r}),INDEX({Z}!{RNG},0,$CG${r}),$Z{w}))')
            for k, c in eraw.items():
                R.set(f'{c}{w}', formula=f'IF(OR($AB{w}=0,${ecol[k]}${r}=0),"",SUMIFS(INDEX({Wn}!{RNG},0,${ecol[k]}${r}),INDEX({Wn}!{RNG},0,$CX${r}),$Z{w}))')
            R.set(f'BL{w}', formula=f'IF($Z{w}="","",INDEX(VERW_Project!$N$6:$N$105,{p}))')
            g = lambda k: f'N({zraw[k]}{w})'; h = lambda k: f'N({eraw[k]}{w})'
            aZ, aE = f'$AA{w}=0', f'$AB{w}=0'
            months = '+'.join(f'IF({pm}>={m},{h(f"M{m:02d}")},0)' for m in range(1, 13))
            F = {
             'd0_omz': f'IF({aZ},"",IF({off}=0,-({g("BLJ")}+{g("BX")}),-{g("BX1")}))',
             'd0_bor': f'IF({aZ},"",IF({off}=0,-({g("RLJ")}+{g("RX")}),-{g("RX1")}))',
             'omz_ytd': f'IF(OR({aZ},{off}<>0),"",-{g("BLJ")})',
             'bor_ytd': f'IF(OR({aZ},{off}<>0),"",-{g("RLJ")})',
             'd0_ak': f'IF(OR({aZ},$BL{w}=""),"",IF({off}=0,{g("KLJ")}+{g("KX")},{g("KX1")})/P_AKFactor*$BL{w})',
             'w_fy0': f'IF({aE},"",IF({off}=0,{h("HJ")},{h("VJR")}))',
             'w_tmvj': f'IF(OR({aE},{off}<>0),"",{h("TMV")})',
             'w_ytd': f'IF(OR({aE},{off}<>0),"",{months})',
             'd1_omz': f'IF({aZ},"",IF({off}=0,-{g("BX1")},-{g("BX2")}))',
             'd1_bor': f'IF({aZ},"",IF({off}=0,-{g("RX1")},-{g("RX2")}))',
             'd1_ak': f'IF(OR({aZ},$BL{w}=""),"",IF({off}=0,{g("KX1")},{g("KX2")})/P_AKFactor*$BL{w})',
             'w_fy1': f'IF(OR({aE},{off}<>0),"",{h("VJR")})',
            }
            for n in range(1, 5):
                m = [h(f'M{3*(n-1)+j:02d}') for j in (1, 2, 3)]
                F[f'w_q{n}'] = f'IF({aE},"",IF({off}=0,{"+".join(m)},{h(f"VQ{n}")}))'
            for key, c in dcol.items(): R.set(f'{c}{w}', formula=F[key])
    # lange tabel
    row = L0
    for vi, v in enumerate(VERS):
        r = VROW[v]
        for p in range(1, NP + 1):
            w = W0 + vi*NP + p - 1
            for (kg, dj, key) in SLOTS:
                R.set(f'A{row}', text='PLAK'); R.set(f'B{row}', text=zs(v) if key[0] != 'w' else es(v))
                R.set(f'D{row}', text='Vastgesteld'); R.set(f'F{row}', text=v); R.set(f'G{row}', formula='0')
                R.set(f'E{row}', formula=f'Versies!$F${r}')
                R.set(f'H{row}', formula=f'Versies!$D${r}')
                R.set(f'I{row}', formula=f'Versies!$F${r}+{dj}')
                R.set(f'J{row}', is_str=True, formula=f'$Z${w}')
                R.set(f'M{row}', is_str=True, formula=f'IF($Z${w}="","",INDEX(V_Type,{p})&"")')
                R.set(f'N{row}', is_str=True, formula=f'IF($Z${w}="","",INDEX(V_IFRS,{p})&"")')
                R.set(f'P{row}', formula=f'${dcol[key]}${w}')
                R.set(f'O{row}', is_str=True, formula=f'IF(OR($Z${w}="",P{row}=""),"","{kg}")')
                R.set(f'S{row}', is_str=True, formula=f'IF(O{row}="","",J{row}&"|"&E{row}&"|"&F{row}&"|"&I{row}&"|"&O{row})')
                R.set(f'T{row}', formula='0'); R.set(f'U{row}', formula=f'IF(O{row}="",0,1)'); R.set(f'V{row}', formula='0')
                row += 1
    assert row - 1 == LN
    R.root.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}dimension').set('ref', f'A1:DQ{LN}')
    # ---------- namen naar VERSIE_REKEN
    for nm, c in (('L_ImportID','A'),('L_Status','D'),('L_Versiejaar','E'),('L_Versie','F'),('L_Revisie','G'),('L_Doeljaar','I'),
                  ('L_WBS','J'),('L_Type','M'),('L_IFRS','N'),('L_Kengetal','O'),('L_Waarde','P'),('L_Toelichting','R'),
                  ('L_Sleutel','S'),('L_Actief','U'),('L_Dubbel','V')):
        bk.set_name(nm, f'VERSIE_REKEN!${c}${L0}:${c}${LN}')
    # ---------- officiële AK per versie in CALC_Versies (AB5:AB9)
    cv = bk.sheet('CALC_Versies')
    for i, v in enumerate(VERS):
        r = 5 + i; el = cv.cell(f'AB{r}'); f = el.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}f').text
        cv.set(f'AB{r}', style=cv.style(f'AB{r}'), formula=f'IF(AND(ISNUMBER(Versies!$K${VROW[v]}),Versies!$F${VROW[v]}=P_VJ,Sel_Doeljaar=P_VJ),Versies!$K${VROW[v]},{f})')

def styles(bk):
    inv = bk.sheet('Invoer'); his = bk.sheet('Historie')
    return dict(title=inv.style('A5'), text=inv.style('A6'), label=inv.style('C9'), inp_txt=inv.style('D14'),
                inp_date=inv.style('D10'), inp_num=inv.style('D9'), hdr=his.style('B5'), cell=his.style('A6'),
                num=his.style('P6'), date=his.style('C6'), help=inv.style('E9'))

if __name__ == '__main__':
    bk = Book(sys.argv[1]); build(bk, styles(bk)); bk.save(sys.argv[2])
