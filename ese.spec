# -*- mode: python ; coding: utf-8 -*-
# ESE-ONLY PyInstaller build — packages just the strain-energy tool.
# Build on Windows with:  pyinstaller ese.spec
# Output: dist\ese_breakdown\ese_breakdown.exe  (one-folder)
#
# Why this is simpler than the full suite: the entry point run_ese.py imports
# energy_breakdown as a FLAT module (not the `modules.` package), so it avoids
# the package-resolution problem that greys out the suite's postprocessing tools.
# It also needs no scipy/matplotlib, so the build is much smaller.

import os
from PyInstaller.utils.hooks import collect_all

datas = []
binaries = []
hiddenimports = []

# Only the ESE tool's real dependencies. NOT scipy/matplotlib (unused here);
# pyNastran still pulls scipy automatically if it actually needs it.
for pkg in ('customtkinter', 'tksheet', 'pyNastran', 'openpyxl'):
    d, b, h = collect_all(pkg)
    datas += d
    binaries += b
    hiddenimports += h

# energy_breakdown.py lives in postprocessing/modules but is imported FLAT by
# run_ese.py; bdf_utils/_version live at the repo root. Flat hidden imports
# resolve reliably (same as mass_scale in the suite build).
hiddenimports += [
    'energy_breakdown', 'bdf_utils', '_version', '_build',
    'matplotlib.backends.backend_tkagg',  # harmless if matplotlib absent
]

block_cipher = None

a = Analysis(
    ['run_ese.py'],
    pathex=[
        SPECPATH,
        os.path.join(SPECPATH, 'postprocessing', 'modules'),
        os.path.join(SPECPATH, 'postprocessing'),
    ],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    # Same None-stdout/stderr guard the suite needs (customtkinter writes a
    # warning to stderr at import; windowed builds have stderr = None).
    runtime_hooks=[os.path.join(SPECPATH, 'rthook_stdio.py')],
    excludes=['PyQt5', 'PySide2', 'PyQt6', 'PySide6', 'IPython', 'pytest'],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# ONE-FILE build: bundle binaries + data INTO the exe (no _internal folder).
# Produces a single dist\ese_breakdown.exe. Trade-off vs one-folder: each launch
# silently unpacks to a temp dir first (a few seconds here since there's no
# scipy/matplotlib), and single self-extracting exes draw more antivirus
# attention. For ESE-only that's an acceptable price for one portable file.
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='ese_breakdown',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,            # avoid antivirus false positives
    runtime_tmpdir=None,
    console=False,        # GUI app — set True to debug a crash
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
