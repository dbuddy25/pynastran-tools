# Building the Structures Tools Windows executable

This produces a standalone `structures_tools.exe` that runs the **whole suite**
on a Windows machine with no Python install required. You hand a colleague a
folder; they double-click the `.exe`.

> **Why this has to run on Windows:** PyInstaller cannot cross-compile. A Windows
> `.exe` must be built *on* Windows. Do these steps on the work machine (or a
> Windows VM), not on a Mac.

---

## 1. Prerequisites (one time)

1. Install **Python 3.10–3.12 (64-bit)** from python.org. During install, tick
   **"Add Python to PATH."**
   - Match the bitness to the machine (64-bit is normal). Avoid 3.13+ until you've
     confirmed pyNastran/scipy wheels exist for it.
2. Open a fresh **Command Prompt** (or PowerShell) so the new PATH takes effect.

## 2. Get the code onto the Windows machine

Copy the whole project folder over (git clone, zip, or a network share). You need
the full tree — `structures_tools.py`, `structures_tools.spec`, `preprocessing/`,
`postprocessing/`, `calculators/`, `bdf_utils.py`, `_version.py`, `_build.py`,
`requirements.txt`.

## 3. Create a clean virtual environment

A venv keeps the bundle small and avoids dragging in unrelated packages.

```bat
cd path\to\pynastran
python -m venv .venv
.venv\Scripts\activate
python -m pip install --upgrade pip
```

You should see `(.venv)` at the start of the prompt.

## 4. Install dependencies + PyInstaller

```bat
pip install -r requirements.txt
pip install pyinstaller
```

(`pyvista` in requirements.txt is optional 3D preview — if it fails to install,
edit it out of requirements.txt and re-run; the suite runs without it.)

## 5. Build

```bat
pyinstaller structures_tools.spec
```

First build takes a few minutes (it's collecting scipy, matplotlib, pyNastran).
When it finishes you'll have:

```
dist\structures_tools\structures_tools.exe   <-- the app
dist\structures_tools\...                     <-- supporting files (DLLs, data)
```

## 6. Test it

```bat
dist\structures_tools\structures_tools.exe
```

The launcher window should open with all tools. Open the **ESE Breakdown** tool,
load an OP2 (or punch), and confirm results appear.

## 7. Hand it to a colleague

The app is the **entire `dist\structures_tools\` folder**, not just the `.exe` —
the exe needs the sibling files next to it. So:

1. Zip the `dist\structures_tools\` folder.
2. Send the zip. They unzip anywhere and run `structures_tools.exe` inside.

No Python needed on their end.

---

## Troubleshooting

**The window flashes and closes / nothing happens.**
Temporarily switch to a console build to see the error: open
`structures_tools.spec`, change `console=False` to `console=True`, rebuild, and
run the exe from a Command Prompt. The traceback tells you what's missing.

**`ModuleNotFoundError: No module named 'X'` at runtime.**
Add `'X'` to the `hiddenimports` list in `structures_tools.spec` and rebuild.
(This is the usual fix for a lazily-imported module PyInstaller didn't see.)

**A theme / data file is missing (customtkinter or matplotlib error).**
The `collect_all(...)` loop at the top of the spec already covers these; make sure
the package name is in that loop, then rebuild.

**Crash on launch: `'NoneType' object has no attribute 'write'` (customtkinter).**
Already handled. In a windowed build `sys.stdout`/`sys.stderr` are `None`, and
customtkinter writes a font warning to stderr at import, crashing the app. The
runtime hook `rthook_stdio.py` (registered via `runtime_hooks` in the spec) gives
those streams a dummy sink before any import. Keep `rthook_stdio.py` next to the
spec; if you ever see this error, confirm the `runtime_hooks=[...]` line is intact.

**Antivirus quarantines the exe.**
Common with PyInstaller one-folder/one-file binaries — it's a false positive. The
spec sets `upx=False`, which reduces this. If IT still flags it, they may need to
allow-list the file, or you sign it with a code-signing certificate.

**You want a single .exe instead of a folder.**
One-file is more convenient to share but starts slower (it unpacks to a temp dir
each launch — noticeable with scipy/pyNastran) and is more AV-prone. To switch,
replace the `EXE(...)` + `COLLECT(...)` block at the bottom of the spec with a
single one-file `EXE`:

```python
exe = EXE(
    pyz, a.scripts, a.binaries, a.datas, [],
    name='structures_tools',
    debug=False, strip=False, upx=False,
    console=False, runtime_tmpdir=None,
)
```

(Delete the `COLLECT(...)` call.) Rebuild — you'll get a single
`dist\structures_tools.exe`.

---

## Optional: stamp the build SHA

`_build.py` looks for a `_build_info.py` baked at freeze time so the About box can
show which commit was built. To include it, create `_build_info.py` next to
`structures_tools.py` before building:

```bat
git rev-parse --short HEAD > tmp.txt
set /p SHA=<tmp.txt
echo __build__ = "%SHA%" > _build_info.py
del tmp.txt
```

Then add `'_build_info'` to `hiddenimports` in the spec and rebuild. This is
purely cosmetic — the app runs fine without it (it falls back to the live git SHA,
or "unknown").
