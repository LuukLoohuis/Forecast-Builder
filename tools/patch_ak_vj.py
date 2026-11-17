"""AK-dekking volgende jaren: 2027 per kwartaal + 2028-2030 per jaar, in tabblad AK-dekking en Outlook (per project).
Draai na patch_ak.py. usage: python3 -I patch_ak_vj.py in.xlsm out.xlsm"""
import sys, zipfile, re
sys.path.insert(0, __import__('os').path.dirname(__file__))
from patch_ak import Sheet, sheet_paths, N
from lxml import etree

def add_merge(s, ref):
    mc = s.root.find(N+'mergeCells')
    if mc is None: return
    if any(m.get('ref')==ref for m in mc): return
    e = etree.SubElement(mc, N+'mergeCell'); e.set('ref', ref)
    mc.set('count', str(len(mc)))

def set_cols(s, specs):
    cols = s.root.find(N+'cols')
    for (mn, mx, width, hidden) in specs:
        e = etree.SubElement(cols, N+'col'); e.set('min', str(mn)); e.set('max', str(mx))
        e.set('width', str(width)); e.set('customWidth', '1')
        if hidden: e.set('hidden', '1')
    # keep sorted
    items = sorted(cols, key=lambda c: int(c.get('min')))
    for c in list(cols): cols.remove(c)
    for c in items: cols.append(c)

STGQ = ['AS','AT','AU','AV']          # Kosten RJ+1 Q1..Q4 in STG_ZPMW03
QAK  = ['AI','AJ','AK','AL']          # AK RJ+1 per kwartaal
YR   = [('AM','S'),('AN','U'),('AO','W'),('AP','Y')]   # jaar RJ+1..RJ+4 (bestaande berekening)
QK   = ['AR','AS','AT','AU']          # hulp: kosten RJ+1 per kwartaal
RANK = 'AV'

def patch_akdek(s):
    hdr = s.style('S20'); dat = s.style('F21'); tot = s.style('S122')
    s.set('AI19', text='AK-DEKKING VOLGENDE JAREN (kosten / factor x AK %)', style=s.style('R19'))
    for c in ['AJ','AK','AL','AM','AN','AO','AP']: s.set(f'{c}19', style=s.style('R19'))
    add_merge(s, 'AI19:AP19')
    for q in range(4):
        s.set(f'{QAK[q]}20', formula=f'(P_RJ+1)&" Q{q+1}"', style=hdr, is_str=True)
        s.set(f'{QK[q]}20', formula=f'"Kosten"&CHAR(10)&(P_RJ+1)&" Q{q+1}"', style=hdr, is_str=True)
    for i,(c,_) in enumerate(YR):
        s.set(f'{c}20', formula=f'"Jaar"&CHAR(10)&(P_RJ+{i+1})', style=hdr, is_str=True)
    s.set(f'{RANK}20', text='Volgnr Outlook', style=hdr)
    for r in range(21, 121):
        for q in range(4):
            s.set(f'{QK[q]}{r}', style=s.style(f'S{r}'), formula=
                  f'IF(C{r}<>"Ja","",SUMIFS(STG_ZPMW03!${STGQ[q]}$6:${STGQ[q]}$505,STG_ZPMW03!$B$6:$B$505,$A{r}))')
            s.set(f'{QAK[q]}{r}', style=dat, formula=
                  f'IF(OR(C{r}<>"Ja",D{r}=""),"",N({QK[q]}{r})/P_AKFactor*D{r})')
        for c, src in YR:
            s.set(f'{c}{r}', style=dat, formula=f'IF(OR(C{r}<>"Ja",D{r}=""),"",N({src}{r}))')
        s.set(f'{RANK}{r}', style=s.style(f'S{r}'), formula=
              f'IF(OR(C{r}<>"Ja",D{r}=""),"",IF(ABS(N(AM{r}))+ABS(N(AN{r}))+ABS(N(AO{r}))+ABS(N(AP{r}))+ABS(N(AI{r}))+ABS(N(AJ{r}))+ABS(N(AK{r}))+ABS(N(AL{r}))<0.5,"",MAX(${RANK}$20:{RANK}{r-1})+1))')
    for c in QAK + [y for y,_ in YR] + QK:
        s.set(f'{c}122', style=tot, formula=f'SUM({c}21:{c}120)')
    # samenvatting bovenaan (rij 5-9)
    for q in range(4):
        s.set(f'{QAK[q]}5', formula=f'(P_RJ+1)&" Q{q+1}"', style=s.style('F5'), is_str=True)
    for i,(c,_) in enumerate(YR):
        s.set(f'{c}5', formula=f'"Jaar "&(P_RJ+{i+1})', style=s.style('F5'), is_str=True)
    for c in QAK + [y for y,_ in YR]:
        s.set(f'{c}6', text='prognose', style=s.style('F6'))
        s.set(f'{c}7', formula=f'SUM(${c}$21:${c}$120)', style=s.style('F7'))
        s.set(f'{c}9', formula=f'{c}7', style=s.style('F9'))
    s.set('AQ6', is_str=True, style=s.style('F6'), formula=
          'IF(ABS(SUM(AI7:AL7)-AM7)>1,"Let op: som kwartalen "&(P_RJ+1)&" wijkt "&FIXED(SUM(AI7:AL7)-AM7,0)&" af van het jaar (ZPMW03)","")')
    dim = s.root.find(N+'dimension')
    if dim is not None: dim.set('ref', 'A1:AV122')
    set_cols(s, [(34,34,3,False),(35,42,12,False),(43,43,3,False),(44,48,12,True)])

def patch_outlook(s):
    t0 = 194
    s.set(f'B{t0}', style=s.style('B90'), is_str=True, formula=
          '"AK-dekking per project "&(P_RJ+1)&" t/m "&(P_RJ+4)&" ("&P_Eenheid&"; kosten / "&FIXED(P_AKFactor,2)&" x AK %, alle projecten met AK meenemen = Ja)"')
    h = t0 + 1
    hs = s.style('B91')
    s.set(f'B{h}', text='Project', style=hs)
    heads = [('E','(P_RJ+1)&" Q1"'),('F','(P_RJ+1)&" Q2"'),('G','(P_RJ+1)&" Q3"'),('H','(P_RJ+1)&" Q4"'),
             ('L','"Jaar "&(P_RJ+1)'),('M','"Jaar "&(P_RJ+2)'),('N','"Jaar "&(P_RJ+3)'),('O','"Jaar "&(P_RJ+4)')]
    for c,f in heads: s.set(f'{c}{h}', formula=f, style=hs, is_str=True)
    src = {'E':'AI','F':'AJ','G':'AK','H':'AL','L':'AM','M':'AN','N':'AO','O':'AP'}
    ts, ns, ix = s.style('D92'), s.style('G92'), s.style('M92')
    first, last = h+1, h+100
    for r in range(first, last+1):
        k = r - h
        s.set(f'X{r}', style=ix, formula=f"IFERROR(MATCH({k},'AK-dekking'!$AV$21:$AV$120,0),0)")
        s.set(f'B{r}', style=ts, is_str=True, formula=f"IF($X{r}=0,\"\",INDEX('AK-dekking'!$B$21:$B$120,$X{r}))")
        for c, a in src.items():
            s.set(f'{c}{r}', style=ns, formula=f"IF($X{r}=0,\"\",N(INDEX('AK-dekking'!${a}$21:${a}$120,$X{r}))/P_Factor)")
    tr = last + 1
    s.set(f'B{tr}', text='Totaal', style=s.style('B9') or hs)
    for c in src:
        s.set(f'{c}{tr}', style=s.style('E9') or ns, formula=f'SUM({c}{first}:{c}{last})')
    s.set(f'B{tr+1}', style=ts, is_str=True, formula=
          f'IF(ABS(L{tr}-N(E18))+ABS(M{tr}-N(F18))+ABS(N{tr}-N(G18))+ABS(O{tr}-N(H18))<1,"Sluit aan op de regel AK-dekking hierboven.","Wijkt af van de regel AK-dekking hierboven: daar telt het filter (type/project) mee.")')
    dim = s.root.find(N+'dimension')
    if dim is not None: dim.set('ref', f'A1:X{tr+1}')
    set_cols(s, [(24,24,6,True)])

def main(src, dst):
    zin = zipfile.ZipFile(src); paths = sheet_paths(zin)
    targets = {paths['AK-dekking']: patch_akdek, paths['Outlook']: patch_outlook}
    zout = zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename in targets:
            s = Sheet(data); targets[item.filename](s); data = s.xml()
        zout.writestr(item, data)
    zout.close()

if __name__ == '__main__': main(sys.argv[1], sys.argv[2])
