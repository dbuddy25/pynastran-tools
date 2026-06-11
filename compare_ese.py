#!/usr/bin/env python3
"""Side-by-side compare of ESE from an OP2 (via pyNastran) vs a punch (ASCII).

The punch is ground truth (matches Femap). This shows exactly where and how
pyNastran's OP2 read diverges from it — per mode and per element.

IMPORTANT: the OP2 and punch must be from the SAME run, otherwise modes /
eigenvector bases won't line up. Request both at once with, e.g.:
    ESE(PLOT,PUNCH) = ALL

Usage:
    python compare_ese.py model.op2 model.pch
"""
import os
import sys

import numpy as np

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(_HERE, 'postprocessing', 'modules'))
sys.path.insert(0, os.path.join(_HERE, 'preprocessing'))

from energy_breakdown import parse_ese_punch   # noqa: E402


def read_op2_ese(op2_path):
    """Return (modes, energy_by_eid, percent_by_eid) read from the OP2.

    Uses the same logic as the tool: pick the ESE (field2==1) key, read the
    energy (col 0) and percent (col 1) columns, exclude eid<=0 and eid>=1e8,
    first row wins per eid.
    """
    from pyNastran.op2.op2 import read_op2
    op2 = read_op2(op2_path, build_dataframe=False, debug=False)
    eig = next(iter(op2.eigenvalues.values()))
    modes = np.asarray(eig.mode)
    nmodes = len(modes)

    se = getattr(getattr(op2, 'op2_results', None), 'strain_energy', None) or op2
    energy = {}
    percent = {}
    for attr in dir(se):
        if not attr.endswith('_strain_energy'):
            continue
        d = getattr(se, attr, None)
        if not isinstance(d, dict) or not d:
            continue
        r = None
        for k, v in d.items():
            if isinstance(k, tuple) and len(k) >= 3 and k[2] == 1:
                r = v
                break
        if r is None:
            r = next(iter(d.values()))
        data = getattr(r, 'data', None)
        if data is None or getattr(data, 'ndim', 0) != 3 or data.shape[2] < 2:
            continue
        el = r.element
        two_d = getattr(el, 'ndim', 1) == 2
        nm = min(nmodes, data.shape[0])
        # pyNastran stores a per-mode element order; pair each mode's data row
        # with THAT mode's element order (element[mi]), not element[0].
        for mi in range(nm):
            eids_row = el[mi] if two_d else el
            for j in range(min(len(eids_row), data.shape[1])):
                ei = int(eids_row[j])
                if not (0 < ei < 100000000):
                    continue
                if ei not in energy:
                    energy[ei] = np.zeros(nmodes)
                    percent[ei] = np.zeros(nmodes)
                energy[ei][mi] = data[mi, j, 0]
                percent[ei][mi] = data[mi, j, 1]
    return modes, energy, percent


def main(op2_path, punch_path):
    print(f"Reading OP2 via pyNastran: {op2_path}")
    op2_modes, op2_e, op2_p = read_op2_ese(op2_path)
    print(f"Parsing punch:           {punch_path}")
    pu_modes, pu_freqs, pu_e = parse_ese_punch(punch_path)

    print(f"\nOP2:   {len(op2_modes)} modes, {len(op2_e)} elements")
    print(f"Punch: {len(pu_modes)} modes, {len(pu_e)} elements")

    op2_idx = {int(m): i for i, m in enumerate(op2_modes)}
    pu_idx = {int(m): i for i, m in enumerate(pu_modes)}
    common_modes = sorted(set(op2_idx) & set(pu_idx))
    eids = sorted(set(op2_e) & set(pu_e))
    only_op2 = len(set(op2_e) - set(pu_e))
    only_pu = len(set(pu_e) - set(op2_e))
    print(f"common: {len(common_modes)} modes, {len(eids)} elements "
          f"(OP2-only elems: {only_op2}, punch-only: {only_pu})")
    if not eids or not common_modes:
        print("\nNo overlap — are the OP2 and punch from the same run?")
        return

    # Pre-stack into arrays [nmodes_common, n_eids] for speed.
    oe = np.array([[op2_e[e][op2_idx[m]] for e in eids] for m in common_modes])
    pe = np.array([[pu_e[e][pu_idx[m]] for e in eids] for m in common_modes])
    op = np.array([[op2_p[e][op2_idx[m]] for e in eids] for m in common_modes])

    print(f"\n{'mode':>5} {'OP2 Esum':>11} {'punch Esum':>11} {'E ratio':>8} "
          f"{'OP2 %sum':>9} {'maxrelE':>8} {'#elem>1%':>9}")
    for r, m in enumerate(common_modes):
        oes, pes = oe[r].sum(), pe[r].sum()
        ratio = oes / pes if pes else float('nan')
        denom = np.where(np.abs(pe[r]) < 1e-30, 1.0, np.abs(pe[r]))
        rel = np.abs(oe[r] - pe[r]) / denom
        ndiff = int((rel > 0.01).sum())
        print(f"{m:>5} {oes:>11.4g} {pes:>11.4g} {ratio:>8.3f} "
              f"{op[r].sum():>9.1f} {rel.max():>8.3f} {ndiff:>9}")

    # Detail on the mode with the biggest energy-sum mismatch.
    worst = int(np.argmax(np.abs(oe.sum(axis=1) - pe.sum(axis=1))))
    m = common_modes[worst]
    print(f"\n=== detail: mode {m} (largest OP2-vs-punch energy mismatch) ===")
    diff = np.abs(oe[worst] - pe[worst])
    order = np.argsort(diff)[::-1][:15]
    print(f"{'eid':>10} {'OP2 energy':>14} {'punch energy':>14} "
          f"{'OP2/punch':>10} {'OP2 percent':>12}")
    for k in order:
        ratio = oe[worst][k] / pe[worst][k] if pe[worst][k] else float('nan')
        print(f"{eids[k]:>10} {oe[worst][k]:>14.5g} {pe[worst][k]:>14.5g} "
              f"{ratio:>10.3f} {op[worst][k]:>12.3f}")

    print("\nHow to read this:")
    print("  - 'E ratio' ~1.0 and 'OP2 %sum' ~100 -> OP2 read agrees with punch")
    print("  - 'E ratio' ~2.0 or 'OP2 %sum' ~200 -> OP2 read is doubling that mode")
    print("  - look at the detail rows to see if it's all elements or a subset")


if __name__ == '__main__':
    if len(sys.argv) != 3:
        print("Usage: python compare_ese.py model.op2 model.pch")
        print("  (both from the SAME run, e.g. ESE(PLOT,PUNCH)=ALL)")
        sys.exit(1)
    main(sys.argv[1], sys.argv[2])
