# Structures Tools

A suite of structural-analysis tools with a unified desktop GUI — Nastran FEA
pre/post-processing (built on [pyNastran](https://github.com/SteveDoyle2/pyNastran))
plus hand-calculation utilities. Built with customtkinter + tksheet.

Runs from source on any machine with **Python 3** — no build step, no installer.

## Quick start

**Windows (easiest):** install Python 3.12 from [python.org](https://python.org)
(tick *Add Python to PATH*), then double-click:

- **`launch_ese.bat`** — just the ESE (strain energy) tool. Lighter.
- **`launch_structures_tools.bat`** — the full suite.

The first run installs the needed packages, then launches. Later runs just
launch. See [SETUP.md](SETUP.md) for details and troubleshooting.

**Any OS / terminal:**

```bash
pip install -r requirements.txt
python structures_tools.py        # full suite
python run_ese.py                 # ESE tool only
```

## Tools

**Pre-processing** (BDF manipulation)
- **Mass Scale** — per-include-file mass scaling (density, NSM, CONM2 mass/inertia)
- **Renumber** — per-include-file ID renumbering with validation
- **Thermal CTE** — set/verify a uniform CTE (incl. RBE2 ALPHA/TREF) across a model
- **Partition** — split a model into include files

**Post-processing** (OP2 results)
- **ESE Breakdown** — element strain-energy %ESE by group (Property ID, Include
  File, or **element-ID range** — CSV-importable, no BDF required). Matches Femap;
  reads OP2 or punch.
- **Modal Effective Mass** — MEFFMASS fractions from OP2, with Excel export
- **CBUSH Forces** — CBUSH force extraction by joint
- **Mass Breakdown** — mass properties by group (incl. DMIG), GPWG validation
- **ASD Overlay** — acceleration spectral density plotting/overlay
- **Response Limiting** — response-limiting / force-limited vibration plots
- **Random Vibe Environment** — random-vibration ASD environment builder

**Calculators**
- **Miles Equation** — Miles' equation random-vibration response

Most modules also run standalone, e.g. `python postprocessing/modules/meff.py`.

## Dependencies

Installed automatically by the launchers / `requirements.txt`:

- **pyNastran** (pinned to 1.4.1) — BDF/OP2 reader
- **customtkinter**, **tksheet** — GUI
- **numpy**, **scipy** — numerics
- **matplotlib** — plotting (ASD / response / random-vibe tools)
- **openpyxl** — Excel export

The ESE tool alone needs only `customtkinter tksheet numpy pyNastran openpyxl`
(no scipy/matplotlib).

## Sharing a build

To hand the tools to a colleague, build a minimal zip (only the runtime files —
no `.venv`, build artifacts, or dev files):

```bash
python make_bundle.py        # -> structures_tools_v<version>.zip
```

They unzip it, install Python, and double-click a `launch_*.bat`. See
[SETUP.md](SETUP.md). To update, send a new bundle and unzip over the old folder.

## Project structure

```
structures_tools.py              # Unified GUI launcher
run_ese.py                       # ESE-only launcher
launch_ese.bat / launch_structures_tools.bat   # one-click launchers
make_bundle.py                   # build a share zip
preprocessing/                   # BDF tools + bdf_utils.py (shared helpers)
postprocessing/
  ├── modal_effective_mass.py    # MEFFMASS CLI
  └── modules/                   # ESE, CBUSH, mass, ASD, response, random-vibe, meff
calculators/                     # Miles equation
docs/                            # implementation notes
```

## License

MIT — see [LICENSE](LICENSE).
