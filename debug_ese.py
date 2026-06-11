#!/usr/bin/env python3
"""Diagnostic for the strain-energy (ESE) tool.

Dumps the structure of the strain-energy results in an OP2 so we can see:
  - how many subcase keys each *_strain_energy table has (the "multiple ESE
    subcases" warning comes from len(keys) > 1),
  - whether each table's PERCENT column already sums to ~100 per mode (if a
    table sums to >100 on its own, the percents are normalized per-partition),
  - the element-id range of each table, and
  - whether any table is actually KINETIC energy (EKE) misfiled as strain
    energy (title / data_code / table_code check).

No BDF needed. Read-only.

Usage:
    python debug_ese.py path/to/model.op2
"""
import sys

import numpy as np
from pyNastran.op2.op2 import read_op2


def _energy_kind(result):
    """Best-effort guess of strain vs kinetic energy for one result object."""
    title = str(getattr(result, 'title', '') or '')
    dc = str(getattr(result, 'data_code', {}) or {})
    tcode = getattr(result, 'table_code', None)
    blob = (title + ' ' + dc).upper()
    if 'KINETIC' in blob:
        return 'KINETIC (EKE)?'
    # table_code 18 == element strain energy in pyNastran's ONRGY reader
    if tcode is not None and tcode not in (18,):
        return f'table_code={tcode} (NOT 18=strain energy)'
    return 'strain energy'


def main(op2_path):
    print(f"Reading {op2_path} ...\n")
    op2 = read_op2(op2_path, build_dataframe=False, debug=False)

    eig = getattr(op2, 'eigenvalues', {}) or {}
    print(f"Eigenvalue subcases: {list(eig.keys())}")
    for k, v in eig.items():
        modes = getattr(v, 'modes', None)
        n = len(modes) if modes is not None else '?'
        print(f"  eig subcase {k}: {n} modes")

    se = getattr(getattr(op2, 'op2_results', None), 'strain_energy', None)
    if se is None:
        print("\nNo op2.op2_results.strain_energy found.")
        return

    print("\n=== strain_energy tables (what the ESE tool reads) ===")
    found = False
    for attr in sorted(dir(se)):
        if not attr.endswith('_strain_energy'):
            continue
        d = getattr(se, attr, None)
        if not isinstance(d, dict) or not d:
            continue
        found = True
        print(f"\n{attr}: {len(d)} key(s) -> {list(d.keys())}")
        for key, r in d.items():
            data = getattr(r, 'data', None)
            if data is None or getattr(data, 'ndim', 0) != 3:
                print(f"  key {key}: (no 3D data)")
                continue
            ntime, nelem, ncols = data.shape
            eids = r.element
            eids = eids[0] if getattr(eids, 'ndim', 1) == 2 else eids
            modes = getattr(r, 'modes', None)
            modes_s = list(np.asarray(modes)[:6]) if modes is not None else None
            try:
                emin, emax = int(np.min(eids)), int(np.max(eids))
            except Exception:
                emin = emax = '?'
            kind = _energy_kind(r)
            print(f"  key {key}: nelem={nelem} ntime={ntime} ncols={ncols} "
                  f"table={getattr(r, 'table_name', '')!r}  kind={kind}")
            print(f"      title={str(getattr(r, 'title', '') or '')!r}")
            print(f"      modes[:6]={modes_s}  eid_range={emin}..{emax}")
            if ncols >= 2:
                psum = data[:6, :, 1].sum(axis=1)
                print(f"      PERCENT_sum/mode[:6]={np.round(psum, 1).tolist()}")
            if 'KINETIC' in kind.upper() or 'NOT 18' in kind:
                print("      *** NOT plain strain energy — likely EKE or other ***")

    if not found:
        print("  (none found)")

    print("\nKey questions answered by the above:")
    print("  - Any table with >1 key  -> source of the 'multiple subcases' note")
    print("  - PERCENT_sum/mode > 100 -> percents normalized per-partition")
    print("  - a KINETIC/NOT-18 table -> EKE is mixed into strain_energy")


if __name__ == '__main__':
    if len(sys.argv) != 2:
        print("Usage: python debug_ese.py path/to/model.op2")
        sys.exit(1)
    main(sys.argv[1])
