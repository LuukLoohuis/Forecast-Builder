"""Versies uit geplakte ZPMW03-exports (Budget, Q1-Q4). Rekenblad VERSIE_REKEN zet ze om naar de lange tabel
die de rapporten lezen (L_-namen). Woningen komen alleen uit de huidige Woningetalage.
usage: python3 -I patch_versies.py in.xlsm out.xlsm"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from xlsxlib import Book, skeleton, q
from openpyxl.utils import get_column_letter as CL, column_index_from_string as CI

VERS = ['Budget', 'Q1', 'Q2', 'Q3', 'Q4']
VROW = {v: 6 + i for i, v in enumerate(VERS)}
NP = 150; NSRC = 600; RNG = '$A$1:$GZ$600'
ZKEYS = [('BLJ', '"Bedrijfsopbrengsten LJ"'), ('RLJ', '"Bruto (omzet) result. LJ"'), ('KLJ', '"Kosten LJ"'),
         ('BX', '"OP.BOB "&{X}'), ('RX', '"OP.BOR "&{X}'), ('KX', '"Kosten "&{X}'),
         ('BX1', '"OP.BOB "&({X}+1)'), ('RX1', '"OP.BOR "&({X}+1)'), ('KX1', '"Kosten "&({X}+1)'),
         ('BX2', '"OP.BOB "&({X}+2)'), ('RX2', '"OP.BOR "&({X}+2)'), ('KX2', '"Kosten "&({X}+2)'),
         ('KQ1', '"Kosten "&{E}&"-Q1"'), ('KQ2', '"Kosten "&{E}&"-Q2"'), ('KQ3', '"Kosten "&{E}&"-Q3"'), ('KQ4', '"Kosten "&{E}&"-Q4"')]
SLOTS = [('Omzet FY', 0, 'd0_omz'), ('BOR FY', 0, 'd0_bor'), ('Omzet actual YTD', 0, 'omz_ytd'), ('BOR actual YTD', 0, 'bor_ytd'),
         ('AK-dekking FY', 0, 'd0_ak'), ('Omzet FY', 1, 'd1_omz'), ('BOR FY', 1, 'd1_bor'), ('AK-dekking FY', 1, 'd1_ak')]
NS = len(SLOTS)
L0 = 6; LN = L0 + len(VERS) * NP * NS - 1
W0 = 6
PL = 'Projectlijst!$A$8:$A$107'; PAN = 'Projectlijst!$AN$8:$AN$107'
def zs(v): return f'{v} ZPMW03'

def build(bk, st):
    for v in VERS:
        s = bk.add_sheet(zs(v), skeleton(tab='FFFEB900'))
        s.set('A1', text=f'Plak hier de ZPMW03-export van {v} (hele blad, vanaf A1).')
    V = bk.add_sheet('Versies', skeleton(
        cols='<cols><col min="1" max="1" width="2" customWidth="1"/><col min="2" max="2" width="12" customWidth="1"/>'
             '<col min="3" max="4" width="13" customWidth="1"/><col min="5" max="7" width="10" customWidth="1"/>'
             '<col min="8" max="9" width="14" customWidth="1"/><col min="10" max="10" width="16" customWidth="1"/>'
             '<col min="11" max="14" width="13" customWidth="1"/><col min="15" max="15" width="48" customWidth="1"/></cols>',
        tab='FF111F39', hidden_grid=True))
    V.set('B1', text='Versies: Budget, Q1, Q2, Q3, Q4', style=st['title'])
    V.set('B2', style=st['text'], text="Plak per versie de ZPMW03-export in de gele tabbladen (bijv. 'Q1 ZPMW03'). De rapporten tonen de versie meteen.")
    V.set('B3', style=st['text'], text="Budget = de prognose die vorig jaar voor dit jaar is afgegeven (bijv. de export van Q3 2025 voor budget 2026). Peildatum leeg = standaard.")
    hdr = ['Versie', 'Peildatum export', 'Gebruikt', 'Exportjaar', 'Versiejaar', 'Verschuiving', 'Peilmaand',
           'Projecten', 'Niet op Projectlijst', 'AK-dekking officieel (EUR, optioneel)', 'Omzet FY (EUR)', 'BOR FY (EUR)', 'AK-dekking FY (EUR)', 'Status']
    for i, h in enumerate(hdr): V.set(f'{CL(2+i)}5', text=h, style=st['hdr'])
    for i, v in enumerate(VERS):
        r = VROW[v]; a, b = W0 + i*NP, W0 + i*NP + NP - 1
        V.set(f'B{r}', text=v, style=st['label']); V.set(f'C{r}', style=st['inp_date'])
        dflt = 'DATE(P_VJ-1,9,30)' if v == 'Budget' else f'DATE(P_VJ,{3*i+1},0)'
        V.set(f'D{r}', formula=f'IF(ISNUMBER(C{r}),C{r},{dflt})', style=st['date'])
        V.set(f'E{r}', formula=f'YEAR(D{r})', style=st['cell'])
        V.set(f'F{r}', formula=f'E{r}+{1 if v=="Budget" else 0}', style=st['cell'])
        V.set(f'G{r}', formula=f'F{r}-E{r}', style=st['cell'])
        V.set(f'H{r}', formula=f'MONTH(D{r})', style=st['cell'])
        V.set(f'I{r}', formula=f'IF(VERSIE_REKEN!$CE${r}=0,"leeg",COUNTIF(VERSIE_REKEN!$AA${a}:$AA${b},">0"))', style=st['cell'])
        V.set(f'J{r}', formula=f'IF(VERSIE_REKEN!$CE${r}=0,"",SUMPRODUCT((VERSIE_REKEN!$Z${a}:$Z${b}<>"")*(COUNTIF({PL},VERSIE_REKEN!$Z${a}:$Z${b})=0)))', style=st['cell'])
        V.set(f'K{r}', style=st['inp_num'])
        for col, k in (('L', 'Omzet FY'), ('M', 'BOR FY'), ('N', 'AK-dekking FY')):
            V.set(f'{col}{r}', style=st['num'], formula=
                  f'IF(VERSIE_REKEN!$CE${r}=0,"",SUMIFS(L_Waarde,L_Versie,$B{r},L_Versiejaar,$F{r},L_Doeljaar,$F{r},L_Kengetal,"{k}"))')
        V.set(f'O{r}', style=st['text'], is_str=True, formula=
              f'IF(VERSIE_REKEN!$CE${r}=0,"– nog niet geplakt",IF(VERSIE_REKEN!$CG${r}=0,"⚠ kolom WBS-element niet gevonden","✔ geplakt")'
              f'&IF($F{r}<>P_VJ," · versiejaar "&$F{r}&" (rapport toont "&P_VJ&")","")&IF(N(J{r})>0," · "&J{r}&" niet op Projectlijst (tellen als Opstal)","")&IF(K{r}<>""," · officieel AK telt",""))')
    V.set('B12', style=st['text'], text="AK-dekking per versie = kosten / factor (Invoer) x AK %. Vul 'AK-dekking officieel' in om het gerapporteerde bedrag te tonen.")
    V.set('B13', style=st['text'], text="Woningaantallen komen alleen uit de huidige Woningetalage. Projecten op 'Buiten scope' of buiten het deelgebiedfilter tellen niet mee.")
    V.set('B15', style=st['text'], text='Berekening: tabblad VERSIE_REKEN →')
    V.root.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}dimension').set('ref', 'A1:O15')

    R = bk.add_sheet('VERSIE_REKEN', skeleton(tab='FFA4B3BB'))
    R.set('A1', text='VERSIE_REKEN: versies uit de plakbladen (niet wijzigen). A:V = lange tabel (L_-namen) · X:CD = per versie x project · CE: kolomnummers · EA: WBS-lijst per versie.')
    R.set('A2', text='Omzet FY = −(Bedrijfsopbrengsten LJ + OP.BOB jaar); volgend jaar −OP.BOB volgend jaar. BOR idem met Bruto result. LJ / OP.BOR. AK = kosten / factor x AK %.')
    for i, h in enumerate(['ImportID','Bronbestand','Importdatum','Status','Versiejaar','Versie','Revisie','Peildatum','Doeljaar','WBS',
                           'Projectdefinitie','Projectnaam','Type','IFRS','Kengetal','Waarde','Eenheid','Toelichting','Sleutel','MaxRev','Actief','Dubbel']):
        R.set(f'{CL(1+i)}5', text=h)
    zcol = {k: CL(CI('CH') + j) for j, (k, _) in enumerate(ZKEYS)}
    R.set('CE5', text='kopregel'); R.set('CG5', text='WBS'); R.set('DR5', text='IFRS'); R.set('DS5', text='# WBS')
    for k, c in zcol.items(): R.set(f'{c}5', text=k)
    filt = 'OR(Sel_Deelgebied="Alle",Sel_Deelgebied="",Sel_Deelgebied="(geen)")'
    for vi, v in enumerate(VERS):
        r = VROW[v]; X = f'Versies!$E${r}'; E = f'Versies!$F${r}'; Z = q(zs(v))
        R.set(f'CE{r}', formula=''.join(f'IF(COUNTIF({Z}!${k}:${k},"WBS-element")>0,{k},' for k in range(1, 16)) + '0' + ')'*15)
        look = lambda name: f'IF($CE${r}=0,0,IFERROR(MATCH({name},INDEX({Z}!$A$1:$GZ$15,$CE${r},0),0),0))'
        R.set(f'CG{r}', formula=look('"WBS-element"'))
        R.set(f'DR{r}', formula=look('"IFRS J/N"'))
        for k, expr in ZKEYS: R.set(f'{zcol[k]}{r}', formula=look(expr.format(X=X, E=E)))
        cz, rz = CL(CI('EA') + 2*vi), CL(CI('EA') + 2*vi + 1)
        R.set(f'{cz}5', text=f'{v} WBS'); R.set(f'{rz}5', text='nr')
        for i in range(1, NSRC + 1):
            row = 5 + i
            R.set(f'{cz}{row}', is_str=True, formula=f'IF(OR($CE${r}=0,$CG${r}=0,{i}<=$CE${r}),"",TRIM(INDEX({Z}!{RNG},{i},$CG${r})&""))')
            R.set(f'{rz}{row}', formula=(f'IF(OR(LEN({cz}{row})<6,ISERROR(FIND(".",{cz}{row})),COUNTIF(${cz}$6:{cz}{row},{cz}{row})>1,'
                                        f'COUNTIFS({PL},{cz}{row},{PAN},"Nee")>0,AND(COUNTIF({PL},{cz}{row})=0,NOT({filt}))),"",MAX(${rz}$5:{rz}{row-1})+1)'))
        R.set(f'DS{r}', formula=f'N(MAX(${rz}$6:${rz}${5+NSRC}))')
    zraw = {k: CL(CI('AC') + j) for j, (k, _) in enumerate(ZKEYS)}
    dcol = {key: CL(CI('BM') + j) for j, (_, _, key) in enumerate(SLOTS)}
    for vi, v in enumerate(VERS):
        r = VROW[v]; Z = q(zs(v)); off = f'Versies!$G${r}'
        cz, rz = CL(CI('EA') + 2*vi), CL(CI('EA') + 2*vi + 1)
        for p in range(1, NP + 1):
            w = W0 + vi*NP + p - 1
            R.set(f'X{w}', text=v); R.set(f'Y{w}', formula=str(p))
            R.set(f'Z{w}', is_str=True, formula=f'IFERROR(INDEX(${cz}$6:${cz}${5+NSRC},MATCH({p},${rz}$6:${rz}${5+NSRC},0)),"")')
            R.set(f'AA{w}', formula=f'IF(OR($Z{w}="",$CG${r}=0),0,COUNTIF(INDEX({Z}!{RNG},0,$CG${r}),$Z{w}))')
            for k, c in zraw.items():
                R.set(f'{c}{w}', formula=f'IF(OR($AA{w}=0,${zcol[k]}${r}=0),"",SUMIFS(INDEX({Z}!{RNG},0,${zcol[k]}${r}),INDEX({Z}!{RNG},0,$CG${r}),$Z{w}))')
            plm = f'MATCH($Z{w},{PL},0)'
            R.set(f'BL{w}', formula=(f'IF($Z{w}="","",IF(COUNTIFS({PL},$Z{w},Projectlijst!$G$8:$G$107,"Nee")>0,"",'
                                     f'IF(IFERROR(INDEX(Projectlijst!$I$8:$I$107,{plm})&"","")<>"",VALUE(INDEX(Projectlijst!$I$8:$I$107,{plm})&""),'
                                     f'IFERROR(INDEX(Bron_AK!$A$2:$Z$3000,MATCH($Z{w},INDEX(Bron_AK!$A$2:$Z$3000,0,AK_KolWBS),0),AK_KolPct),""))))'))
            R.set(f'CD{w}', is_str=True, formula=f'IF($Z{w}="","",IFERROR(INDEX(Projectlijst!$C$8:$C$107,{plm})&"",""))')
            g = lambda k: f'N({zraw[k]}{w})'; aZ = f'$AA{w}=0'
            F = {
             'd0_omz': f'IF({aZ},"",IF({off}=0,-({g("BLJ")}+{g("BX")}),-{g("BX1")}))',
             'd0_bor': f'IF({aZ},"",IF({off}=0,-({g("RLJ")}+{g("RX")}),-{g("RX1")}))',
             'omz_ytd': f'IF(OR({aZ},{off}<>0),"",-{g("BLJ")})',
             'bor_ytd': f'IF(OR({aZ},{off}<>0),"",-{g("RLJ")})',
             'd0_ak': f'IF(OR({aZ},$BL{w}=""),"",IF({off}=0,{g("KLJ")}+{g("KX")},{g("KX1")})/P_AKFactor*$BL{w})',
             'd1_omz': f'IF({aZ},"",IF({off}=0,-{g("BX1")},-{g("BX2")}))',
             'd1_bor': f'IF({aZ},"",IF({off}=0,-{g("RX1")},-{g("RX2")}))',
             'd1_ak': f'IF(OR({aZ},$BL{w}=""),"",IF({off}=0,{g("KX1")},{g("KX2")})/P_AKFactor*$BL{w})',
            }
            for key, c in dcol.items(): R.set(f'{c}{w}', formula=F[key])
    row = L0
    for vi, v in enumerate(VERS):
        r = VROW[v]
        for p in range(1, NP + 1):
            w = W0 + vi*NP + p - 1
            for (kg, dj, key) in SLOTS:
                R.set(f'A{row}', text='PLAK'); R.set(f'B{row}', text=zs(v))
                R.set(f'D{row}', is_str=True, formula=f'IF(O{row}="","","Vastgesteld")'); R.set(f'F{row}', text=v); R.set(f'G{row}', formula='0')
                R.set(f'E{row}', formula=f'Versies!$F${r}'); R.set(f'H{row}', formula=f'Versies!$D${r}')
                R.set(f'I{row}', formula=f'Versies!$F${r}+{dj}')
                R.set(f'J{row}', is_str=True, formula=f'IF(O{row}="","",$Z${w})')
                R.set(f'M{row}', is_str=True, formula=f'IF($Z${w}="","",IF($CD${w}<>"",$CD${w},"Opstal"))')
                R.set(f'N{row}', is_str=True, formula=f'IF($Z${w}="","",IFERROR(INDEX(V_IFRS,MATCH($Z${w},V_WBS,0))&"",IFERROR(INDEX({q(zs(v))}!{RNG},MATCH($Z${w},INDEX({q(zs(v))}!{RNG},0,$CG${r}),0),$DR${r})&"","J")))')
                R.set(f'P{row}', formula=f'${dcol[key]}${w}')
                R.set(f'O{row}', is_str=True, formula=f'IF(OR($Z${w}="",P{row}=""),"","{kg}")')
                R.set(f'S{row}', is_str=True, formula=f'IF(O{row}="","",J{row}&"|"&E{row}&"|"&F{row}&"|"&I{row}&"|"&O{row})')
                R.set(f'T{row}', formula='0'); R.set(f'U{row}', formula=f'IF(O{row}="",0,1)'); R.set(f'V{row}', formula='0')
                row += 1
    assert row - 1 == LN
    R.root.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}dimension').set('ref', f'A1:{CL(CI("EA")+2*len(VERS)-1)}{max(LN, 5+NSRC)}')
    for nm, c in (('L_ImportID','A'),('L_Status','D'),('L_Versiejaar','E'),('L_Versie','F'),('L_Revisie','G'),('L_Doeljaar','I'),
                  ('L_WBS','J'),('L_Type','M'),('L_IFRS','N'),('L_Kengetal','O'),('L_Waarde','P'),('L_Toelichting','R'),
                  ('L_Sleutel','S'),('L_Actief','U'),('L_Dubbel','V')):
        bk.set_name(nm, f'VERSIE_REKEN!${c}${L0}:${c}${LN}')
    cv = bk.sheet('CALC_Versies')
    for i, v in enumerate(VERS):
        r = 5 + i; f = cv.cell(f'AB{r}').find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}f').text
        cv.set(f'AB{r}', style=cv.style(f'AB{r}'), formula=f'IF(AND(ISNUMBER(Versies!$K${VROW[v]}),Versies!$F${VROW[v]}=P_VJ,Sel_Doeljaar=P_VJ),Versies!$K${VROW[v]},{f})')

def styles(bk):
    inv = bk.sheet('Invoer'); his = bk.sheet('Historie')
    return dict(title=inv.style('A5'), text=inv.style('A6'), label=inv.style('C9'), inp_txt=inv.style('D14'),
                inp_date=inv.style('D10'), inp_num=inv.style('D9'), hdr=his.style('B5'), cell=his.style('A6'),
                num=his.style('P6'), date=his.style('C6'), help=inv.style('E9'))

if __name__ == '__main__':
    bk = Book(sys.argv[1]); build(bk, styles(bk)); bk.save(sys.argv[2])
