# pynastran-tools

Python GUI and CLI tools for Nastran FEA model manipulation, built on [pyNastran](https://github.com/SteveDoyle2/pyNastran).

## Quick Start

```bash
pip install -r requirements.txt
python nastran_tools.py
```

## Tools

### Unified GUI (`nastran_tools.py`)

Single-window application with sidebar navigation grouping all tools:

- **Pre-Processing**
  - **Mass Scale** — per-include-file mass scaling (material density, NSM, CONM2 mass/inertia)
  - **Renumber** — per-include-file ID renumbering with validation
- **Post-Processing**
  - **MEFF Viewer** — modal effective mass fractions from OP2, with comparison and Excel export

### Standalone tools

Each tool also works standalone:

```bash
# Mass scaling GUI
python preprocessing/mass_scale.py

# Include file renumbering GUI
python preprocessing/renumber_includes.py

# MEFF Viewer GUI (standalone)
python postprocessing/modules/meff.py

# Modal effective mass report (CLI)
python postprocessing/modal_effective_mass.py model.op2
python postprocessing/modal_effective_mass.py model.op2 --xlsx output.xlsx
python postprocessing/modal_effective_mass.py baseline.op2 --compare updated.op2
```

## Installation

```bash
pip install -r requirements.txt
```

### Dependencies

- **pyNastran** — BDF/OP2 reader/writer
- **customtkinter** — modern tkinter widgets
- **tksheet** — spreadsheet widget
- **numpy**, **scipy** — numerical computation (modal effective mass)
- **openpyxl** — Excel export (optional)

## Project Structure

```
├── nastran_tools.py                # Unified GUI entry point (sidebar navigation)
├── preprocessing/
│   ├── __init__.py
│   ├── bdf_utils.py                # Shared utilities (IncludeFileParser, CARD_ENTITY_MAP, make_model)
│   ├── mass_scale.py               # Mass scaling tool (CTkFrame, standalone or embedded)
│   └── renumber_includes.py        # Include file renumbering tool (CTkFrame, standalone or embedded)
├── postprocessing/
│   ├── __init__.py
│   ├── modal_effective_mass.py     # MEFFMASS CLI tool
│   └── modules/
│       ├── __init__.py
│       └── meff.py                 # MEFF module (CustomTkinter + tksheet, comparison logic, Excel helpers)
└── docs/
    ├── bdf_utils.md
    ├── mass_scaling_implementation.md
    ├── renumber_includes_tool.md
    └── modal_effective_mass_guide.md
```

## License

MIT — see [LICENSE](LICENSE).
