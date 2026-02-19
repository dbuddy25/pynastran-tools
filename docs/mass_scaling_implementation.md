# BDF Mass Scaling Tool — Implementation

GUI tool available standalone or in the unified app.

```
pip install customtkinter tksheet
python nastran_tools.py          # unified app (Pre-Processing > Mass Scale)
python preprocessing/mass_scale.py  # standalone
```

## What It Does

- Opens a BDF/DAT file (with INCLUDE files)
- Shows mass breakdown by include file in a **tksheet** spreadsheet (5 columns)
- Uses `IncludeFileParser` from `bdf_utils.py` to reliably map cards to their source include files
- Displays WTMASS value as an informational label (not as table columns)
- **Divide by 386.1 toggle** — checkbox to divide all displayed masses by 386.1 (display only, does not affect written BDF). Useful for IPS unit conversion (slinch to lbf).
- Editable scale factor per file — preview updates live
- Writes scaled model back, preserving include structure
- Modern UI via **CustomTkinter** (works on macOS and Windows)

## Dependencies

- `customtkinter` — modern-looking tkinter widgets
- `tksheet` — spreadsheet widget (replaces hand-built scrollable grid)
- `pyNastran` — BDF/OP2 reader
- `bdf_utils` — shared utilities module (local)

## Table Columns

| # | Column | Editable | Notes |
|---|--------|----------|-------|
| 0 | File Name | No | Include file basename |
| 1 | Original Mass | No | Scientific notation, affected by /386.1 toggle |
| 2 | Scale Factor | **Yes** | Default "1.0000" |
| 3 | Scaled Mass | No | = Original x Scale, affected by /386.1 toggle |
| 4 | Delta | No | Percentage change, e.g. "+0%" |

Last row is a bold TOTAL row (highlighted, not editable).

## What Gets Scaled

| Card type | Field scaled | Skipped when |
|---|---|---|
| MAT1, MAT8, MAT9 | `rho` (density) | rho is 0 or None |
| PSHELL, PCOMP, PBAR, PBARL, PBEAM, PBEAML, PROD | `nsm` | nsm is 0 or None |
| CONROD | `nsm` | nsm is 0 or None |
| CONM2 | `mass` and all 6 `I` components | never |
| CONM1 | `mass_matrix` | never |
| CMASS1, CMASS2 | `mass` | never |

## Source

See `preprocessing/mass_scale.py` (imports `IncludeFileParser` and `make_model` from `bdf_utils.py`).

---

## How It Works

### Loading

1. Click **Open BDF**
2. Tries `read_bdf(path, save_file_structure=True)` first — if that fails, falls back to `read_bdf(path)` without structure tracking
3. `cross_reference()` links cards so `elem.Mass()` works
4. Reads `PARAM,WTMASS` (defaults to 1.0 if absent) — displayed as info label
5. `_build_ifile_lookup()` uses `IncludeFileParser` to parse raw BDF text and map each card ID to its source file index
6. `_compute_groups()` iterates all elements/masses and groups by file index
7. `_populate_sheet()` fills tksheet with one row per include file plus TOTAL row

### /386.1 Toggle

- Checkbox in the info bar: "Divide displayed masses by 386.1"
- When checked, all displayed mass values (Original Mass, Scaled Mass, summary) are divided by 386.1
- **Does not affect scale factors or written output** — purely a display conversion
- Toggling calls `_refresh_display()` which recomputes all displayed values

### Contact Card Handling

Cards like BCPROPS, BGPARM, BCTPARM, etc. are disabled via `make_model(_CARDS_TO_SKIP)`. Disabled cards are stored as rejected card text and written back out unchanged.

### Live Preview

Editing a scale factor in the sheet triggers `<<SheetModified>>` -> `_refresh_display()`. This reads all scale factors from column 2, computes `scaled = original x scale` for each row, applies the /386.1 divisor for display, and updates columns 1, 3, 4 plus the TOTAL row and summary label.

### Writing

1. All scale factors validated as valid floats (read from sheet column 2)
2. `SaveModeDialog` (CTkToplevel) — choose: add suffix (default `_scaled`), output directory, or overwrite
3. **Capture** original card values (rho, nsm, mass, inertia) into a dict
4. **Apply** scale factors in-place to model cards
5. `uncross_reference()` then `write_bdf()` / `write_bdfs()`
6. **Restore** originals from snapshot
7. `cross_reference()` to return model to working state

### Edge Cases

- **`save_file_structure=True` fails**: automatically retries without it; file grouping still works because `IncludeFileParser` reads raw BDF text independently
- **No includes**: all cards grouped as main file
- **Zero-density materials**: skipped (no mass created where none existed)
- **Empty include files**: shown with mass=0, scale entry read-only
- **Invalid scale factor**: preview treats as 1.0; write rejects with error
- **`write_bdfs` unavailable**: falls back to single consolidated file
- **/386.1 toggle**: display-only, never affects written BDF values
