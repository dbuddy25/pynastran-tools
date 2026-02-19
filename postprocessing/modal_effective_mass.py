#!/usr/bin/env python3
"""Modal effective mass fraction report for Nastran SOL 103 OP2 files.

Reads MEFFMASS fraction data (EFMFACS) written by Nastran when
MEFFMASS(PLOT) is present in case control and prints a single table
with per-mode fractions and cumulative sums for each direction.

Requires: pip install pyNastran numpy
Optional: pip install openpyxl  (for --xlsx export)

Usage:
    python modal_effective_mass.py model.op2
    python modal_effective_mass.py model.op2 --xlsx output.xlsx
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

try:
    from modules.meff import (make_meff_styles, write_meff_single_sheet)
except ImportError:
    from postprocessing.modules.meff import (make_meff_styles,
                                              write_meff_single_sheet)


DIRECTIONS = ['Tx', 'Ty', 'Tz', 'Rx', 'Ry', 'Rz']


def read_op2_file(filename: str) -> OP2:
    """Read an OP2 file, exiting on failure."""
    try:
        op2 = OP2(mode='nx')
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

    Returns a dict with keys *modes*, *freqs*, *frac*, *cumsum*.
    """
    if not op2.eigenvalues:
        print("ERROR: No eigenvalue tables found \u2014 is this a SOL 103 run?")
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


def export_to_excel(data, xlsx_path):
    """Export to a formatted single-sheet Excel workbook."""
    try:
        from openpyxl import Workbook
    except ImportError:
        print("ERROR: openpyxl is required for Excel export. "
              "Install with: pip install openpyxl")
        sys.exit(1)

    wb = Workbook()
    styles = make_meff_styles()
    ws = wb.active
    ws.title = "Effective Mass Fractions"
    write_meff_single_sheet(ws, data, styles)

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
    parser.add_argument('--xlsx', '-x', type=str, default=None,
                        help='Export to Excel (.xlsx) file')
    args = parser.parse_args()

    op2 = read_op2_file(args.op2)
    data = _read_data(op2)

    print_table(data)

    if args.xlsx:
        export_to_excel(data, args.xlsx)


if __name__ == '__main__':
    main()
