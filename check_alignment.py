#!/usr/bin/env python3
"""Test the 'wrong element order' hypothesis for the OP2 ESE mismatch.

pyNastran stores element IDs PER mode (result.element is [ntimes, nelems]).
Our reader uses element[0] for every mode. If the per-mode order varies, that
mislabels energies on other modes (and can leak the eid=1e8 total row onto a
real element, doubling the mode). This checks:

  1. Does the per-mode element order actually vary (element[0] != element[k])?
  2. Does using element[mode] (instead of element[0]) make the OP2 grand-total
     energy per mode match the punch?

Usage: python check_alignment.py model.op2 model.pch   (both from the same run)
"""
import os
import sys

import numpy as np

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(_HERE, 'postprocessing', 'modules'))
sys.path.insert(0, os.path.join(_HERE, 'preprocessing'))
from energy_breakdown import parse_ese_punch   # noqa: E402


def main(op2_path, punch_path):
    from pyNastran.op2.op2 import read_op2
    op2 = read_op2(op2_path, build_dataframe=False, debug=False)
    eig = next(iter(op2.eigenvalues.values()))
    op2_modes = np.asarray(eig.mode)
    nmodes = len(op2_modes)
    se = getattr(getattr(op2, 'op2_results', None), 'strain_energy', None) or op2

    results = []
    print("=== does element order vary per mode? ===")
    varies = False
    for attr in sorted(dir(se)):
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
        results.append(r)
        el = np.asarray(r.element)
        if el.ndim == 2 and el.shape[0] > 1:
            k = min(5, el.shape[0] - 1)
            same = bool(np.array_equal(el[0], el[k]))
            print(f"  {attr}: element shape {el.shape}  element[0]==element[{k}]? {same}")
            if not same:
                varies = True
        else:
            print(f"  {attr}: element ndim={el.ndim} (single order)")
    print(f"\n  -> per-mode element order VARIES: {varies}")

    def grand_total(use_per_mode):
        tot = np.zeros(nmodes)
        for r in results:
            data = np.asarray(r.data)
            el = np.asarray(r.element)
            if data.ndim != 3 or data.shape[2] < 1:
                continue
            nt = min(nmodes, data.shape[0])
            for mi in range(nt):
                if el.ndim == 2:
                    eids_m = el[mi] if use_per_mode else el[0]
                else:
                    eids_m = el
                mask = (eids_m > 0) & (eids_m < 100000000)
                tot[mi] += data[mi, mask, 0].sum()
        return tot

    pu_modes, _, pu_e = parse_ese_punch(punch_path)
    pu_idx = {int(m): i for i, m in enumerate(pu_modes)}
    op_idx = {int(m): i for i, m in enumerate(op2_modes)}
    common = sorted(set(pu_idx) & set(op_idx))
    pu_eids = list(pu_e)

    g0 = grand_total(False)   # element[0] (current)
    gm = grand_total(True)    # element[mode] (proposed fix)

    print(f"\n{'mode':>5} {'punch':>12} {'elem[0]':>12} {'r0':>6} "
          f"{'elem[mode]':>12} {'rm':>6}")
    for m in common[:12]:
        oi = op_idx[m]
        ps = sum(pu_e[e][pu_idx[m]] for e in pu_eids)
        r0 = g0[oi] / ps if ps else float('nan')
        rm = gm[oi] / ps if ps else float('nan')
        print(f"{m:>5} {ps:>12.4g} {g0[oi]:>12.4g} {r0:>6.3f} "
              f"{gm[oi]:>12.4g} {rm:>6.3f}")

    print("\nVerdict:")
    print("  - elem[mode] ratios all ~1.0  -> it's our element-order bug; FIXABLE")
    print("    (switch the reader to per-mode element ordering)")
    print("  - elem[mode] still ~2.0       -> deeper pyNastran read bug; punch only")


if __name__ == '__main__':
    if len(sys.argv) != 3:
        print("Usage: python check_alignment.py model.op2 model.pch")
        sys.exit(1)
    main(sys.argv[1], sys.argv[2])
