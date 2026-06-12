# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec for the full Structures Tools suite.
# Build on Windows with:  pyinstaller structures_tools.spec
# Output: dist\structures_tools\structures_tools.exe  (one-folder)
#
# Why a spec and not a one-line command: customtkinter, tksheet, matplotlib,
# scipy and pyNastran ship data files / dynamically-imported submodules that
# PyInstaller misses by default. collect_all() below grabs them so the frozen
# app doesn't die with "ModuleNotFoundError" or a missing theme/JSON at runtime.

import os
from PyInstaller.utils.hooks import collect_all

datas = []
binaries = []
hiddenimports = []

# Packages PyInstaller under-collects — pull data files + submodules explicitly.
for pkg in ('customtkinter', 'tksheet', 'pyNastran', 'matplotlib',
            'scipy', 'openpyxl'):
    d, b, h = collect_all(pkg)
    datas += d
    binaries += b
    hiddenimports += h

# The launcher inserts preprocessing/, postprocessing/, calculators/ onto
# sys.path and imports modules by BARE name (e.g. `from mass_scale import ...`,
# `from modules.meff import ...`). Name them so PyInstaller's static analysis
# can't miss any. _build_info is intentionally absent (created at freeze, and
# imported under try/except), so it is NOT listed here.
hiddenimports += [
    'mass_scale', 'renumber_includes', 'thermal_cte', 'miles_equation',
    'bdf_utils', '_version', '_build',
    'modal_effective_mass',
    'modules.meff', 'modules.energy_breakdown', 'modules.cbush_forces',
    'modules.mass_breakdown', 'modules.asd_overlay', 'modules.asd_common',
    'modules.response_limiting', 'modules.random_vibe_env',
    # matplotlib's Tk backend is loaded lazily by the plotting tools.
    'matplotlib.backends.backend_tkagg',
]

block_cipher = None

a = Analysis(
    ['structures_tools.py'],
    pathex=[
        SPECPATH,
        os.path.join(SPECPATH, 'preprocessing'),
        os.path.join(SPECPATH, 'postprocessing'),
        os.path.join(SPECPATH, 'calculators'),
    ],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    # Guard None stdout/stderr in windowed builds (fixes customtkinter crash).
    runtime_hooks=[os.path.join(SPECPATH, 'rthook_stdio.py')],
    excludes=['PyQt5', 'PySide2', 'PyQt6', 'PySide6', 'IPython', 'pytest'],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='structures_tools',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,            # UPX can trip antivirus on work machines; leave off.
    console=False,        # GUI app — no console window. Set True to debug crashes.
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon='app.ico',     # uncomment if you add an icon file
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='structures_tools',
)
