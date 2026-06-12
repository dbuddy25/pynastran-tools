#!/usr/bin/env python3
"""Build a minimal share bundle for a colleague who will run the tools from
source (Python install, no exe).

Run:  python make_bundle.py        (or: py make_bundle.py)
Produces: structures_tools_share.zip  in this folder.

Includes only what's needed to run the suite + ESE tool: the launcher scripts,
the three code packages, requirements, the .bat launchers, and SETUP.md.
Excludes everything else — .venv, build/, dist/, *.spec, *.ico, PDFs, reference
docs, and dev/diagnostic scripts.
"""
import os
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))

# Read the suite version so the zip name makes the version obvious when sharing.
_ver = {}
with open(os.path.join(HERE, '_version.py')) as _f:
    exec(_f.read(), _ver)
VERSION = _ver.get('__version__', 'unknown')
VTAG = VERSION.replace('.', '-')   # dashes, not dots, for a clean filename

OUT = os.path.join(HERE, f'structures_tools_v{VTAG}.zip')
TOP = f'structures_tools_v{VTAG}'   # top-level folder name inside the zip

ROOT_FILES = [
    'structures_tools.py', 'run_ese.py',
    '_version.py', '_build.py', 'requirements.txt',
    'launch_ese.bat', 'launch_structures_tools.bat', 'SETUP.md',
]
DIRS = ['preprocessing', 'postprocessing', 'calculators']
SKIP_DIRS = {'__pycache__', '.git', '.venv', 'build', 'dist'}
SKIP_EXT = {'.pyc', '.pyo'}
SKIP_FILES = {'.DS_Store', 'Thumbs.db', 'desktop.ini'}

if os.path.exists(OUT):
    os.remove(OUT)

count = 0
with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
    for f in ROOT_FILES:
        p = os.path.join(HERE, f)
        if os.path.exists(p):
            z.write(p, os.path.join(TOP, f))
            count += 1
        else:
            print('  WARNING missing:', f)
    for d in DIRS:
        droot = os.path.join(HERE, d)
        if not os.path.isdir(droot):
            print('  WARNING missing dir:', d)
            continue
        for dp, dns, fns in os.walk(droot):
            dns[:] = [x for x in dns if x not in SKIP_DIRS]
            for fn in fns:
                if fn in SKIP_FILES or os.path.splitext(fn)[1] in SKIP_EXT:
                    continue
                ap = os.path.join(dp, fn)
                rel = os.path.relpath(ap, HERE)
                z.write(ap, os.path.join(TOP, rel))
                count += 1

size_mb = os.path.getsize(OUT) / 1e6
print(f"\nWrote {OUT}")
print(f"  {count} files, {size_mb:.1f} MB")
print("\nSend that zip. Colleague unzips, installs Python 3.12 from python.org,")
print("then double-clicks launch_ese.bat (or launch_structures_tools.bat).")
