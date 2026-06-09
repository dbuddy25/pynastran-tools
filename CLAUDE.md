# CLAUDE.md — Structures Tools

## Project Overview

Suite of structural analysis tools with a unified desktop GUI. Includes Nastran FEA preprocessing/postprocessing tools (built on pyNastran for BDF/OP2 file handling) and hand calculation utilities. Uses customtkinter for the interface.

## Tech Stack

- **Language**: Python
- **GUI**: customtkinter + tksheet
- **FEA Library**: pyNastran (BDF reading, OP2 result extraction)
- **Excel Export**: openpyxl
- **Version**: see `_version.py` (single source of truth)

## Architecture

```
structures_tools.py          # Unified launcher GUI
preprocessing/            # BDF manipulation tools
  ├── mass_scale.py       # Mass scaling utility
  └── renumber.py         # ID renumbering
postprocessing/           # Results processing
  ├── meff.py             # Modal effective mass viewer
  ├── energy_breakdown.py # Strain energy breakdown
  └── cbush_forces.py     # CBUSH force extraction
bdf_utils.py              # Shared BDF helper functions
docs/                     # Implementation documentation
references/               # Reference materials
skills/pynastran-api/     # Claude skill with pyNastran API reference
```

## Critical Rules

- **No Claude/Anthropic attribution in commits.** No `Co-Authored-By` lines, no AI mentions.
- **Background threading**: Long operations (OP2 loads) must use `_run_in_background(label, work_fn, done_fn)` pattern to keep GUI responsive.
- **Dark mode**: Raw `tk.Listbox` doesn't inherit CTk dark theme — must pass explicit colors: bg `#2b2b2b`, fg `#dce4ee`, selectbackground `#1f6aa5`.
- **Excel exports**: Fixed row positions — always emit all preamble rows even if empty, so header rows are always in the same place.

## Key Patterns

- `_matrix_to_dense()` and `DIRECTIONS` constants live in `modules/meff.py`; CLI imports from there
- tksheet column formatting: `set_all_column_widths()` + `align_columns(..., align="center", align_header=True)`
- See `project-tools-patterns.md` for detailed architecture notes

## Skills

- **pynastran-api**: Local skill at `skills/pynastran-api/SKILL.md` — covers pyNastran v1.4 API for BDF/OP2 reading, result access, MEFFMASS matrices, mass properties, card creation.

## Development Workflow

- Develop and test on macOS
- Keep commits focused on single logical changes
- Run with `python structures_tools.py` for the unified launcher
