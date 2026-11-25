"""Controleert werkbladen op de elementvolgorde die Excel eist (Excel weigert anders het blad)."""
import sys, zipfile, re
from lxml import etree
NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
ORDER = ['sheetPr','dimension','sheetViews','sheetFormatPr','cols','sheetData','sheetCalcPr','sheetProtection','protectedRanges','scenarios',
         'autoFilter','sortState','dataConsolidate','customSheetViews','mergeCells','phoneticPr','conditionalFormatting','dataValidations',
         'hyperlinks','printOptions','pageMargins','pageSetup','headerFooter','rowBreaks','colBreaks','customProperties','cellWatches',
         'ignoredErrors','smartTags','drawing','legacyDrawing','legacyDrawingHF','picture','oleObjects','controls','webPublishItems','tableParts','extLst']
PR = ['tabColor', 'outlinePr', 'pageSetUpPr']
z = zipfile.ZipFile(sys.argv[1]); fout = 0
for n in z.namelist():
    if not re.match(r'xl/worksheets/sheet\d+\.xml$', n): continue
    root = etree.fromstring(z.read(n))
    tags = [etree.QName(c).localname for c in root]
    idx = [ORDER.index(t) if t in ORDER else -1 for t in tags]
    if idx != sorted(idx) or -1 in idx: print(n, 'volgorde:', tags); fout += 1
    pr = root.find(NS+'sheetPr')
    if pr is not None:
        t2 = [etree.QName(c).localname for c in pr]; i2 = [PR.index(t) for t in t2]
        if i2 != sorted(i2): print(n, 'sheetPr:', t2); fout += 1
    sd = root.find(NS+'sheetData'); last = 0
    for row in sd:
        r = int(row.get('r'))
        if r <= last: print(n, 'rijvolgorde', r); fout += 1; break
        last = r; lc = 0
        for c in row:
            m = re.match(r'([A-Z]+)', c.get('r')); v = 0
            for ch in m.group(1): v = v*26 + ord(ch) - 64
            if v <= lc: print(n, 'celvolgorde', c.get('r')); fout += 1; break
            lc = v
            kids = [etree.QName(k).localname for k in c]
            if kids not in ([], ['f'], ['v'], ['f', 'v'], ['is']): print(n, 'cel', c.get('r'), kids); fout += 1; break
    for cf in root.findall(NS+'conditionalFormatting'):
        for ru in cf: 
            if ru.get('priority') is None: print(n, 'cf zonder prioriteit'); fout += 1
print('fouten:', fout)
