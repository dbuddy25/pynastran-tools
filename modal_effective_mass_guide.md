# Modal Effective Mass Fraction Report

A Python CLI tool that reads a Nastran SOL 103 (normal modes) OP2 file and prints a modal effective mass fraction table. Optionally compares two OP2 files to track how modes shift between design iterations. Reads **exact MEFFMASS data** written by Nastran (via `MEFFMASS(PLOT)` case control) — no approximation. Works with **NX Nastran**, **MSC Nastran**, and **Optistruct** OP2 files.

## Requirements

```
pip install pyNastran numpy
```

**openpyxl** is only needed for `--xlsx` Excel export:

```
pip install openpyxl
```

Python 3.9 or later.

## Usage

### Print to console

```
python modal_effective_mass.py model.op2
```

### Export to Excel

```
python modal_effective_mass.py model.op2 --xlsx output.xlsx
```

### Compare two OP2 files

```
python modal_effective_mass.py baseline.op2 --compare updated.op2
```

This prints four tables: File A fractions, File B fractions, comparison by mode number, and comparison by MEFF similarity match.

### Compare with Excel export

```
python modal_effective_mass.py baseline.op2 -c updated.op2 -x compare.xlsx
```

Produces a 4-sheet Excel workbook (see Excel export section below).

## Output

### Single file

A single table with one row per mode. For each of six directions (Tx–Rz) two columns are shown: the per-mode mass fraction and its cumulative sum.

```
  Mode    Freq Tx Frac  Tx Sum Ty Frac  Ty Sum Tz Frac  Tz Sum Rx Frac  Rx Sum Ry Frac  Ry Sum Rz Frac  Rz Sum
----------------------------------------------------------------------------------------------------------------
     1    12.3    0.45    0.45    0.00    0.00    0.00    0.00    0.00    0.00    0.00    0.00    0.00    0.00
     2    18.8    0.00    0.45    0.39    0.39    0.00    0.00    0.00    0.00    0.00    0.00    0.00    0.00
     3    25.1    0.00    0.45    0.00    0.39    0.51    0.51    0.00    0.00    0.00    0.00    0.00    0.00
```

Fraction values are decimals (0.45 = 45%). The Sum columns show the running total — a healthy model has translational sums trending toward 1.0.

### Comparison output

When `--compare` is used, two additional tables are printed after the individual file tables:

**Comparison by Mode Number** — matches modes that share the same mode number between the two files and shows frequency and fraction deltas:

```
  Mode   Freq A   Freq B     Δ Hz      Δ %     ΔTx     ΔTy     ΔTz     ΔRx     ΔRy     ΔRz
---------------------------------------------------------------------------------------------
     1     12.3     12.5      0.2     1.63    0.01    0.00    0.00    0.00    0.00    0.00
     2     18.8     19.1      0.3     1.60    0.00   -0.02    0.00    0.00    0.00    0.00
```

**Comparison by MEFF Match** — for each mode in File A, finds the best-matching mode in File B by cosine similarity of the 6-D MEFFMASS fraction vector. Weak matches (similarity < 0.5) are flagged with `*`:

```
Mode A Match B   Sim   Freq A   Freq B     Δ Hz      Δ %     ΔTx     ΔTy     ΔTz     ΔRx     ΔRy     ΔRz
-----------------------------------------------------------------------------------------------------------
     1       1 0.998     12.3     12.5      0.2     1.63    0.01    0.00    0.00    0.00    0.00    0.00
     2       2 0.995     18.8     19.1      0.3     1.60    0.00   -0.02    0.00    0.00    0.00    0.00
     5       3 0.421*    42.0     25.5    -16.5   -39.29   -0.10    0.15    0.00    0.00    0.00    0.00
```

## Comparison matching strategies

### By Mode Number

Simple: mode 1 in File A is compared to mode 1 in File B. Only modes present in both files are shown. This works well when the mesh/design change is small and modes haven't reordered.

### By MEFF Match (cosine similarity)

For each mode in File A, the algorithm computes cosine similarity of its 6-D MEFFMASS fraction vector against every mode in File B. The absolute value is used to handle eigenvector sign flips between solver runs. The best match is reported along with the similarity score (0–1).

This is a **greedy** match — multiple File A modes can map to the same File B mode. This is intentional: if a mode disappears or splits, you'll see multiple A-modes pointing to the same B-mode (or weak matches), which is a useful diagnostic.

## Excel export

### Single file (`--xlsx` without `--compare`)

Produces a single-sheet workbook titled "Effective Mass Fractions" with dark blue merged direction headers, medium blue sub-headers, right-aligned numbers, light row borders, and frozen panes.

### Comparison (`--xlsx` with `--compare`)

Produces a 4-sheet workbook:

| Sheet | Contents |
|---|---|
| **File A - MEFFMASS** | Full fraction + cumulative sum table for File A |
| **File B - MEFFMASS** | Full fraction + cumulative sum table for File B |
| **Compare - Mode Number** | Mode-number-matched deltas (freq + fraction) |
| **Compare - MEFF Match** | MEFF-similarity-matched deltas with similarity scores; weak-match rows in red |

## GUI usage

The GUI (`python nastran_tool.py`) provides the same functionality through a tab-based interface:

1. **Open primary OP2**: File → Open OP2 (Cmd+O / Ctrl+O)
2. **Open comparison OP2**: File → Open Comparison OP2 (Cmd+Shift+O / Ctrl+Shift+O) — enabled after primary file is loaded
3. **Toggle comparison views**: Radio buttons appear above the table — "By Mode Number" and "By MEFF Match"
4. **Clear comparison**: File → Clear Comparison — reverts to single-file view
5. **Export to Excel**: Click "Export to Excel…" — produces single-sheet or 4-sheet workbook depending on whether a comparison is active

Opening a new primary OP2 automatically clears any existing comparison.

## Nastran model requirements

Your SOL 103 run must include in case control:

```
MEFFMASS(PLOT) = ALL
```

If MEFFMASS data is missing from the OP2, the script prints an error with this instruction.

## Engineering interpretation tips

- **Cumulative Tx, Ty, Tz sums** each trending toward 1.0 as mode count increases indicates good mass participation
- **A few modes dominate** each translational direction — these are your primary structural modes
- **If sums are far below 1.0**, you may need more modes in the analysis
- **Mode swapping**: When the MEFF match assigns a different mode number than the mode-number match, modes have reordered between iterations — common when stiffness changes shift natural frequencies past each other
- **Frequency drift**: Small Δ% values (< 5%) indicate the design change has minimal effect on that mode; large shifts warrant investigation
- **Weak matches** (similarity < 0.5, flagged with `*` in CLI or red in GUI/Excel): The mode shape character has changed significantly — the mode may have merged with another, split, or been replaced by a new mode
- **Multiple A-modes mapping to same B-mode**: Indicates mode coalescence — two distinct modes in the baseline have merged into one in the updated design

## Script

```python
#!/usr/bin/env python3
"""Modal effective mass fraction report for Nastran SOL 103 OP2 files.

Reads MEFFMASS fraction data (EFMFACS) written by Nastran when
MEFFMASS(PLOT) is present in case control and prints a single table
with per-mode fractions and cumulative sums for each direction.

Optionally compares two OP2 files, showing delta tables matched by
mode number and by MEFFMASS cosine similarity.

Requires: pip install pyNastran numpy
Optional: pip install openpyxl  (for --xlsx export)

Usage:
    python modal_effective_mass.py model.op2
    python modal_effective_mass.py model.op2 --xlsx output.xlsx
    python modal_effective_mass.py model.op2 --compare model_v2.op2
    python modal_effective_mass.py model.op2 -c model_v2.op2 -x compare.xlsx
"""
import argparse
import os
import sys

import numpy as np
import scipy.sparse
from pyNastran.op2.op2 import OP2

# Ensure modules/ is importable regardless of working directory
_script_dir = os.path.dirname(os.path.abspath(__file__))
if _script_dir not in sys.path:
    sys.path.insert(0, _script_dir)

from modules.meff import (compare_meff_data, make_meff_styles,
                           write_meff_single_sheet,
                           write_comparison_number_sheet,
                           write_comparison_meff_sheet)


DIRECTIONS = ['Tx', 'Ty', 'Tz', 'Rx', 'Ry', 'Rz']


def read_op2_file(filename: str) -> OP2:
    """Read an OP2 file, exiting on failure."""
    try:
        op2 = OP2()
        op2.read_op2(filename)
    except FileNotFoundError:
        print(f"ERROR: File not found: {filename}")
        sys.exit(1)
    except Exception as exc:
        print(f"ERROR: Could not read OP2 file: {exc}")
        sys.exit(1)
    return op2


def _matrix_to_dense(matrix_obj) -> np.ndarray:
    """Convert a pyNastran Matrix object's data to a dense numpy array."""
    data = matrix_obj.data
    if scipy.sparse.issparse(data):
        return data.toarray()
    return np.asarray(data)


def _read_data(op2):
    """Extract modes, freqs, fractions, and cumulative sums from OP2.

    Returns a dict compatible with compare_meff_data().
    """
    if not op2.eigenvalues:
        print("ERROR: No eigenvalue tables found — is this a SOL 103 run?")
        sys.exit(1)

    eigval_table = next(iter(op2.eigenvalues.values()))
    modes = np.array(eigval_table.mode)
    freqs = np.array(eigval_table.cycles)

    meff_frac = op2.modal_effective_mass_fraction
    if meff_frac is None:
        print("ERROR: No MEFFMASS data found in OP2.")
        print("Add to case control: MEFFMASS(PLOT) = ALL")
        sys.exit(1)

    data = _matrix_to_dense(meff_frac)  # (6, nmodes)
    frac = data.T                        # (nmodes, 6)
    cumsum = np.cumsum(frac, axis=0)

    nmodes = min(frac.shape[0], len(modes))
    return {
        'modes': modes[:nmodes],
        'freqs': freqs[:nmodes],
        'frac': frac[:nmodes],
        'cumsum': cumsum[:nmodes],
    }


def print_table(data):
    """Print the fraction table to stdout."""
    modes, freqs = data['modes'], data['freqs']
    frac, cumsum = data['frac'], data['cumsum']

    hdr = f"{'Mode':>6s} {'Freq':>7s}"
    for d in DIRECTIONS:
        hdr += f" {d + ' Frac':>7s} {d + ' Sum':>7s}"
    print(hdr)
    print("-" * len(hdr))

    for i in range(len(modes)):
        line = f"{modes[i]:6d} {freqs[i]:7.1f}"
        for j in range(6):
            line += f" {frac[i, j]:7.2f} {cumsum[i, j]:7.2f}"
        print(line)


def print_comparison_by_number(comparison):
    """Print the mode-number-matched comparison table."""
    bn = comparison['by_number']

    hdr = f"{'Mode':>6s} {'Freq A':>8s} {'Freq B':>8s} {'Δ Hz':>8s} {'Δ %':>7s}"
    for d in DIRECTIONS:
        hdr += f" {'Δ' + d:>7s}"
    print(hdr)
    print("-" * len(hdr))

    for i in range(len(bn['mode'])):
        line = f"{bn['mode'][i]:6d} {bn['freq_a'][i]:8.1f} {bn['freq_b'][i]:8.1f}"
        line += f" {bn['delta_hz'][i]:8.1f} {bn['delta_pct'][i]:7.2f}"
        for j in range(6):
            line += f" {bn['delta_frac'][i, j]:7.2f}"
        print(line)


def print_comparison_by_meff(comparison):
    """Print the MEFF-matched comparison table.  Weak matches (sim < 0.5)
    are flagged with ``*``."""
    bm = comparison['by_meff']

    hdr = (f"{'Mode A':>6s} {'Match B':>7s} {'Sim':>6s}"
           f" {'Freq A':>8s} {'Freq B':>8s} {'Δ Hz':>8s} {'Δ %':>7s}")
    for d in DIRECTIONS:
        hdr += f" {'Δ' + d:>7s}"
    print(hdr)
    print("-" * len(hdr))

    for i in range(len(bm['mode_a'])):
        sim = bm['similarity'][i]
        weak = '*' if sim < 0.5 else ' '
        line = f"{bm['mode_a'][i]:6d} {bm['match_b'][i]:7d} {sim:5.3f}{weak}"
        line += f" {bm['freq_a'][i]:8.1f} {bm['freq_b'][i]:8.1f}"
        line += f" {bm['delta_hz'][i]:8.1f} {bm['delta_pct'][i]:7.2f}"
        for j in range(6):
            line += f" {bm['delta_frac'][i, j]:7.2f}"
        print(line)


def export_to_excel(data_a, xlsx_path, data_b=None, comparison=None):
    """Export to a formatted Excel workbook.

    When *comparison* is provided, produces a 4-sheet workbook (File A,
    File B, Compare by Mode Number, Compare by MEFF Match).  Otherwise
    produces a single-sheet workbook matching the original format.
    """
    try:
        from openpyxl import Workbook
    except ImportError:
        print("ERROR: openpyxl is required for Excel export. "
              "Install with: pip install openpyxl")
        sys.exit(1)

    wb = Workbook()
    styles = make_meff_styles()
    ws = wb.active

    if comparison is not None and data_b is not None:
        ws.title = "File A - MEFFMASS"
        write_meff_single_sheet(ws, data_a, styles)

        ws_b = wb.create_sheet("File B - MEFFMASS")
        write_meff_single_sheet(ws_b, data_b, styles)

        ws_num = wb.create_sheet("Compare - Mode Number")
        write_comparison_number_sheet(ws_num, comparison, styles)

        ws_meff = wb.create_sheet("Compare - MEFF Match")
        write_comparison_meff_sheet(ws_meff, comparison, styles)
    else:
        ws.title = "Effective Mass Fractions"
        write_meff_single_sheet(ws, data_a, styles)

    try:
        wb.save(xlsx_path)
        print(f"Excel workbook saved to: {xlsx_path}")
    except Exception as exc:
        print(f"ERROR: Failed to write Excel file: {exc}")
        sys.exit(1)


def main() -> None:
    parser = argparse.ArgumentParser(
        description='Modal effective mass fraction report from Nastran OP2.')
    parser.add_argument('op2', help='Path to the OP2 file')
    parser.add_argument('--compare', '-c', type=str, default=None,
                        help='Path to a second OP2 file for comparison')
    parser.add_argument('--xlsx', '-x', type=str, default=None,
                        help='Export to Excel (.xlsx) file')
    args = parser.parse_args()

    op2_a = read_op2_file(args.op2)
    data_a = _read_data(op2_a)

    data_b = None
    comparison = None

    if args.compare:
        print("=== File A ===")
    print_table(data_a)

    if args.compare:
        op2_b = read_op2_file(args.compare)
        data_b = _read_data(op2_b)

        print("\n=== File B ===")
        print_table(data_b)

        comparison = compare_meff_data(data_a, data_b)

        print("\n=== Comparison by Mode Number ===")
        print_comparison_by_number(comparison)

        print("\n=== Comparison by MEFF Match ===")
        print_comparison_by_meff(comparison)

    if args.xlsx:
        export_to_excel(data_a, args.xlsx, data_b, comparison)


if __name__ == '__main__':
    main()
```
