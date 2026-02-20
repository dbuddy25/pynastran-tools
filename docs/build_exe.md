# Building Nastran Tools as a Standalone Executable (Windows)

## Prerequisites

- Python 3.9+ (64-bit)
- All runtime dependencies installed in the build environment

## 1. Install Dependencies

```
pip install -r requirements.txt
pip install pyinstaller
```

## 2. Build with PyInstaller

From the project root (`pynastran/`):

```
pyinstaller --noconfirm --windowed --name "NastranTools" ^
    --add-data "preprocessing;preprocessing" ^
    --add-data "postprocessing;postprocessing" ^
    --hidden-import customtkinter ^
    --hidden-import tksheet ^
    --hidden-import pyNastran ^
    --hidden-import pyNastran.bdf ^
    --hidden-import pyNastran.bdf.bdf ^
    --hidden-import pyNastran.op2 ^
    --hidden-import pyNastran.op2.op2 ^
    --hidden-import scipy ^
    --hidden-import scipy.linalg ^
    --hidden-import numpy ^
    --collect-all customtkinter ^
    nastran_tools.py
```

### What the flags do

| Flag | Purpose |
|------|---------|
| `--windowed` | No console window (GUI app) |
| `--add-data` | Bundles `preprocessing/` and `postprocessing/` packages |
| `--hidden-import` | Forces inclusion of dynamically imported modules |
| `--collect-all customtkinter` | Bundles customtkinter's themes, assets, and JSON files — without this the GUI will fail at runtime |

## 3. Output

```
dist\
  NastranTools\
    NastranTools.exe
    _internal\
```

Run `dist\NastranTools\NastranTools.exe` to verify it launches.

## 4. Single-File Mode (Optional)

Add `--onefile` for a single `.exe`. Startup is slower (unpacks to a temp dir) but distribution is simpler:

```
pyinstaller --noconfirm --windowed --onefile --name "NastranTools" ^
    --add-data "preprocessing;preprocessing" ^
    --add-data "postprocessing;postprocessing" ^
    --hidden-import customtkinter ^
    --hidden-import tksheet ^
    --hidden-import pyNastran ^
    --hidden-import pyNastran.bdf ^
    --hidden-import pyNastran.bdf.bdf ^
    --hidden-import pyNastran.op2 ^
    --hidden-import pyNastran.op2.op2 ^
    --hidden-import scipy ^
    --hidden-import scipy.linalg ^
    --hidden-import numpy ^
    --collect-all customtkinter ^
    nastran_tools.py
```

## Troubleshooting

### customtkinter theme errors

If the app crashes with a missing theme/JSON error, `--collect-all customtkinter` wasn't applied. Verify it's in the command.

### "No module named 'modules.meff'"

The app uses `sys.path` manipulation to import from `preprocessing/` and `postprocessing/`. If PyInstaller doesn't resolve these, add explicit hidden imports:

```
--hidden-import mass_scale
--hidden-import renumber_includes
--hidden-import bdf_utils
--hidden-import modules.meff
--hidden-import modal_effective_mass
```

### numpy/scipy version conflicts

PyInstaller works best when numpy and scipy are installed from PyPI wheels (not conda). If you hit DLL errors, try building in a clean venv:

```
python -m venv build_env
build_env\Scripts\activate
pip install -r requirements.txt
pip install pyinstaller
```

### Debug a failed launch

Build without `--windowed` to see console errors:

```
pyinstaller --noconfirm --name "NastranTools" ...
```

Then run the exe from a terminal to see tracebacks.
