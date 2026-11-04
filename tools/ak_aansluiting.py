"""Aansluiting AK-dekking: rekent een (oud) RDFO-bestand met de tabbladen ZPMW03, AK Percentages en AK realisatie door
volgens de methodes van het nieuwe bestand en de oude RDFO. Gebruik: python3 -I tools/ak_aansluiting.py <bestand.xlsx>"""
import sys, re
from openpyxl import load_workbook
f=sys.argv[1]
wv=load_workbook(f,data_only=True)
z=wv["ZPMW03"]
hdr={z.cell(row=1,column=c).value:c for c in range(1,z.max_column+1)}
def col(n): return hdr.get(n)
bn=[k for k in hdr if k and "Bedrijfsn" in str(k)]
print("bedrijfsnr-kolom:",bn, "omschr:", [k for k in hdr if k and "mschrijving" in str(k)][:2])
pct={}
ap=wv["AK Percentages"]
for r in range(2,ap.max_row+1):
    w=ap.cell(row=r,column=4).value
    if w: pct[str(w).strip()]=ap.cell(row=r,column=8).value or 0
geb={}
a=wv["AK realisatie"]
for r in range(2,a.max_row+1):
    v=a.cell(row=r,column=7).value; t=str(a.cell(row=r,column=15).value or ""); w=a.cell(row=r,column=10).value
    if not isinstance(v,(int,float)): continue
    m=re.search(r"G\.[0-9.]+",t); key=(str(w).strip() if w else (m.group(0).rstrip(".") if m else ""))
    geb[key]=geb.get(key,0)-v
rows=[]
for r in range(2,z.max_row+1):
    w=z.cell(row=r,column=col("WBS-element")).value
    if not w: continue
    w=str(w).strip()
    g=lambda n: z.cell(row=r,column=col(n)).value if col(n) else None
    num=lambda x: x if isinstance(x,(int,float)) else 0
    lj=num(g("Kosten LJ")); q=[num(g(f"Kosten 2026-Q{i}")) for i in range(1,5)]; k26=num(g("Kosten 2026")); k27=num(g("Kosten 2027"))
    p=pct.get(w,None)
    b=str(g(bn[0])) if bn else ""
    rows.append(dict(w=w,b=b,naam=g("Omschrijving"),lj=lj,q=q,k26=k26,k27=k27,p=p,geb=geb.get(w,0)))
F=1.05
def tot(sel, key): return sum(key(x) for x in sel)
sel=[x for x in rows if x["p"]]
print("projecten:",len(rows),"met AK%:",len(sel),"zonder AK%:",[x["w"] for x in rows if x["p"] is None][:8])
nieuw=lambda x: x["geb"]+ x["q"][3]/F*x["p"]           # geboekt t/m Q3 + kosten Q4
kosten=lambda x: (x["lj"]+x["k26"])/F*x["p"]           # Kosten LJ + Kosten 2026 (restant)
for lab,s in (("alle",sel),("bedrijf 1001",[x for x in sel if x["b"]=="1001"])):
    print(f"{lab:14s} n={len(s):3d}  geboekt={tot(s,lambda x:x['geb']):>12,.0f}  nieuw(geb+Q4)={tot(s,nieuw):>12,.0f}  kostenmethode={tot(s,kosten):>12,.0f}  Q4-deel={tot(s,lambda x:x['q'][3]/F*x['p']):>12,.0f}  LJ-deel={tot(s,lambda x:x['lj']/F*x['p']):>12,.0f}")
print("geboekt zonder bekend project:", round(sum(v for k,v in geb.items() if k not in {x['w'] for x in rows})))
print("Grootste negatieve bijdragen (nieuwe methode):")
for x in sorted(sel,key=nieuw)[:6]:
    print(f"  {x['w']:18s} {str(x['naam'])[:28]:28s} geboekt={x['geb']:>10,.0f} kostenQ4={x['q'][3]:>13,.0f} AK%={x['p']} -> {nieuw(x):>10,.0f}")
print("negatieve kosten Q4 totaal:", round(sum(x['q'][3] for x in sel if x['q'][3]<0)), " positieve:", round(sum(x['q'][3] for x in sel if x['q'][3]>0)))
zonder=[x for x in rows if x["p"] is None]
print("zonder AK%:",len(zonder)," kosten LJ:",round(tot(zonder,lambda x:x['lj']))," kosten Q4:",round(tot(zonder,lambda x:x['q'][3])),
      " indicatief AK bij 5%:",round(tot(zonder,lambda x:(x['lj']+x['q'][3])/F*0.05)))
print("waarvan bedrijf:",sorted({x['b'] for x in zonder}))
print("voorbeeld WBS zonder %:",[ (x['w'],str(x['naam'])[:20]) for x in zonder if abs(x['lj'])+abs(x['q'][3])>1e6][:8])
akw=sorted(pct)[:5]; print("AK%-dump WBS voorbeelden:",akw)
onb=sum(v for k,v in geb.items() if k not in {x['w'] for x in rows})
v1=tot(sel,nieuw)+onb
v2=tot(sel,lambda x: x['geb']+max(0,x['q'][3])/F*x['p'])+onb
v3=tot(sel,lambda x: x['geb']+x['q'][3]/F*x['p'] if x['w']!="G.800044.8.1015" else x['geb'])+onb
v4=tot(sel,lambda x: max(0,x['lj']+x['k26'])/F*x['p'])
v5=tot(sel,lambda x: (max(0,x['lj'])+max(0,x['q'][3]))/F*x['p'])
print(f"VARIANTEN (incl. {onb:,.0f} zonder project waar van toepassing):")
print(f"  nieuw bestand (geboekt + kosten Q4):            {v1:>12,.0f}")
print(f"  idem, negatieve kosten Q4 = 0:                  {v2:>12,.0f}")
print(f"  idem, alleen Zuiderveld Spoorzicht eruit:       {v3:>12,.0f}")
print(f"  kostenmethode (LJ+2026) per project min. 0:     {v4:>12,.0f}")
print(f"  kostenmethode, negatieve delen = 0:             {v5:>12,.0f}")
