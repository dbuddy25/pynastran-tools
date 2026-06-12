# Running Structures Tools from Python (no exe)

For a machine that has **Python 3** installed. No PyInstaller build, no giant exe.

## Quick start (Windows)

1. Copy/unzip this project folder onto your machine.
2. Double-click one of:
   - **`launch_ese.bat`** — just the ESE (strain energy) tool. Lighter.
   - **`launch_structures_tools.bat`** — the full Structures Tools suite.

   The first run installs the needed Python packages, then launches. Later runs
   just launch.

That's it. If a window pops up saying a package is missing, run the matching
`.bat` again (it installs them) or use the manual commands below.

## Updating to a newer version

When you get a new bundle, just **unzip it over your existing folder and replace
everything**. Your work isn't stored in here (OP2/BDF files and exports live
wherever you save them), so overwriting is safe. The launchers re-check
dependencies on every run, so any newly required packages install automatically.

## Manual (any OS / terminal)

ESE tool only:
```
pip install customtkinter tksheet numpy pyNastran openpyxl
python run_ese.py
```

Full suite:
```
pip install -r requirements.txt
python structures_tools.py
```

## Notes

- **No Python yet?** Install Python 3.10–3.12 from <https://python.org> and tick
  **"Add Python to PATH"** during setup. Verify with `python --version`.
- **Dependencies** — ESE needs only `customtkinter tksheet numpy pyNastran
  openpyxl` (no scipy/matplotlib). The full suite adds `scipy` and `matplotlib`
  (see `requirements.txt`).
- **No console window?** Rename `run_ese.py` to `run_ese.pyw`, or launch with
  `pythonw run_ese.py`.
- **pip blocked by your network?** Use your organization's internal package
  mirror, or `pip install --index-url <your-mirror> ...`.
