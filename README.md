# pynastran-tools

Python GUI and CLI tools for Nastran FEA model manipulation, built on [pyNastran](https://github.com/SteveDoyle2/pyNastran).

## Tools

### Mass Scaling (`preprocessing/mass_scale.py`)

GUI tool for per-include-file mass scaling. Reads a BDF/DAT with includes, shows mass breakdown by file, lets you apply scale factors (material density, NSM, CONM2 mass/inertia), preview live, and write the scaled model preserving include structure.

### Include File Renumbering (`preprocessing/renumber_includes.py`)

GUI tool for per-include-file ID renumbering. Scans a BDF/DAT to catalog entity types and ID ranges per file, lets you set new ranges (Simple or Advanced mode), validates, and renumbers all cards — including contact, loads, BCs, and case control — writing per-file output.

### Modal Effective Mass Report (`postprocessing/modal_effective_mass.py`)

CLI tool for MEFFMASS fraction reports from SOL 103 OP2 files. Prints per-mode fractions with cumulative sums for each direction (Tx–Rz). Supports comparison of two OP2 files with mode matching by number and by cosine similarity. Optional Excel export.

### Nastran Post-Processing GUI (`postprocessing/nastran_tool.py`)

Tab-based GUI host for post-processing modules. Currently includes the MEFF module for interactive effective mass fraction viewing and comparison, with Excel export.

## Installation

```bash
pip install -r requirements.txt
```

### Dependencies

- **pyNastran** — BDF/OP2 reader/writer
- **customtkinter** — modern tkinter widgets (mass_scale, renumber_includes)
- **tksheet** — spreadsheet widget (mass_scale, renumber_includes)
- **numpy**, **scipy** — numerical computation (modal_effective_mass, nastran_tool)
- **openpyxl** — Excel export (optional, for modal_effective_mass)

## Usage

```bash
# Mass scaling GUI
python preprocessing/mass_scale.py

# Include file renumbering GUI
python preprocessing/renumber_includes.py

# Modal effective mass report (CLI)
python postprocessing/modal_effective_mass.py model.op2
python postprocessing/modal_effective_mass.py model.op2 --xlsx output.xlsx
python postprocessing/modal_effective_mass.py baseline.op2 --compare updated.op2

# Post-processing GUI
python postprocessing/nastran_tool.py
```

## Project Structure

```
├── preprocessing/
│   ├── bdf_utils.py              # Shared utilities (IncludeFileParser, CARD_ENTITY_MAP, make_model)
│   ├── mass_scale.py             # Mass scaling GUI
│   └── renumber_includes.py      # Include file renumbering GUI
├── postprocessing/
│   ├── modal_effective_mass.py   # MEFFMASS CLI tool
│   ├── nastran_tool.py           # Tab-based post-processing GUI
│   └── modules/
│       ├── __init__.py
│       └── meff.py               # MEFF module (comparison logic, Excel helpers, GUI tab)
└── docs/
    ├── bdf_utils.md
    ├── mass_scaling_implementation.md
    ├── renumber_includes_tool.md
    └── modal_effective_mass_guide.md
```

## License

MIT — see [LICENSE](LICENSE).
