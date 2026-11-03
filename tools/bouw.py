"""Alles bouwen: het lege sjabloon en het demo-werkbestand (het demo-FO omgezet, zoals Oud FO inlezen dat doet).

    python tools/bouw.py                      -> output/FO_werkbestand_sjabloon.xlsm en output/FO_Parkzoom_werkbestand.xlsm
    python tools/bouw.py --oud FO.xlsx --project "Naam" --out werkbestand.xlsm      (een ander oud FO omzetten)
"""
from __future__ import annotations

import argparse
import os
import subprocess
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import omzetten  # noqa: E402
import standaard_fo  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BUILD = os.path.join(ROOT, 'tools', 'build_werkbestand.py')


def bouwen(std, out, extra):
    subprocess.run([sys.executable, BUILD, '--fo', std, '--out', out,
                    '--pptx-out', os.path.join(ROOT, 'output', 'Kwartaal_Template_cashflow_v21.pptx')] + extra, check=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--oud', help='oud FO om om te zetten (zonder: sjabloon + demo)')
    ap.add_argument('--project')
    ap.add_argument('--out')
    a = ap.parse_args()
    tmp = tempfile.mkdtemp()
    if a.oud:
        std = os.path.join(tmp, 'standaard.xlsx')
        r = omzetten.omzetten(a.oud, std)
        print('\n'.join('  ' + x for x in r['log']))
        extra = ['--oud', a.oud, '--project', a.project or 'Project']
        if r['oplevering']:
            extra += ['--oplevering', '%d-%d' % tuple(r['oplevering'])]
        if r.get('startbouw'):
            extra += ['--startbouw', '%d-%d' % tuple(r['startbouw'])]
        bouwen(std, a.out, extra)
        return
    print('== Leeg sjabloon')
    std = os.path.join(tmp, 'sjabloon.xlsx')
    standaard_fo.maak(std, leeg=True)
    bouwen(std, os.path.join(ROOT, 'output', 'FO_werkbestand_sjabloon.xlsm'), ['--project', ''])
    print('== Demo (het demo-FO omgezet)')
    std = os.path.join(tmp, 'demo.xlsx')
    r = omzetten.omzetten(standaard_fo.DEMO, std)
    print('\n'.join('  ' + x for x in r['log']))
    bouwen(std, os.path.join(ROOT, 'output', 'FO_Parkzoom_werkbestand.xlsm'),
           ['--oud', standaard_fo.DEMO, '--project', 'Parkzoom', '--oplevering', '2028-3', '--demo'])


if __name__ == '__main__':
    main()
