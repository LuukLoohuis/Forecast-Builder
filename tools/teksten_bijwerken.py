"""Werkt Start-stappen en Handleiding bij (zonder knoppen te verplaatsen). usage: teksten_bijwerken.py in out"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from xlsxlib import Book
import patch_eenvoud as pe
bk = Book(sys.argv[1])
pe.handleiding(bk)
s = bk.sheet('Start')
steps = [('1. Plakken', 'Gele tabbladen: ZPMW03 in het tabblad van het kwartaal (Budget, Q1–Q4), Woningetalage, AK-percentages, AK-realisatie, Mijlpalen.'),
         ('2. Projectlijst', "Knop 'Nieuwe projecten toevoegen'. Oranje = soort kiezen (Opstal / Deelneming / Buiten scope), deelgebied optioneel."),
         ('3. Invoer', "Zet 'Versie van de upload' op het huidige kwartaal (bijv. Q3) en controleer de peildatum. Kolom Status zegt wat ontbreekt."),
         ('4. Bekijken', 'Dashboard: kies kwartaal en deelgebied. Elk rapport heeft een link naar de berekening.'),
         ('5. Presenteren', "Knop 'Maak PowerPoint'.")]
for i, (a, b) in enumerate(steps):
    s.set(f'B{15+i}', text=a, style=s.style(f'B{15+i}')); s.set(f'C{15+i}', text=b, style=s.style(f'C{15+i}'))
s.set('C5', style=s.style('C5'), is_str=True, formula='"Versies: "&"Budget "&IF(N(Versies!$I$6)>0,"✔","–")&"  Q1 "&IF(N(Versies!$I$7)>0,"✔","–")&"  Q2 "&IF(N(Versies!$I$8)>0,"✔","–")&"  Q3 "&IF(N(Versies!$I$9)>0,"✔","–")&"  Q4 "&IF(N(Versies!$I$10)>0,"✔","–")&"  (★ huidig: "&Invoer!$D$14&")"')
bk.save(sys.argv[2])
