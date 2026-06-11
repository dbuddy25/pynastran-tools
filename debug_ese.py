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

    # CLEANED %ESE — EXACTLY what the tool sums: SORT1 only, exclude eid<=0 and
    # eid>=1e8, and DEDUP each eid (first row wins, like the tool). The grand
    # total should equal the GUI flat 'All Elements' view. For any table whose
    # cleaned sum exceeds ~110 % in a mode (a "doubled" table), split it by eid
    # block to reveal whether its energy lives in two separate ID ranges (= two
    # representations of the same structure, the likely double-count source).
    print("\n=== CLEANED %ESE the tool sums (SORT1, no eid<=0/eid>=1e8, deduped) ===")
    grand = None
    seen = set()
    for attr in sorted(dir(se)):
        if not attr.endswith('_strain_energy'):
            continue
        d = getattr(se, attr, None)
        if not isinstance(d, dict) or not d:
            continue
        r = None
        for key, res in d.items():
            if isinstance(key, tuple) and len(key) >= 3 and key[2] == 1:
                r = res
                break
        if r is None:
            r = next(iter(d.values()))
        data = getattr(r, 'data', None)
        if data is None or getattr(data, 'ndim', 0) != 3 or data.shape[2] < 2:
            continue
        eids = r.element
        eids = eids[0] if getattr(eids, 'ndim', 1) == 2 else eids
        nm6 = min(6, data.shape[0])
        keep_eids = []
        keep_cols = []
        dups = 0
        for j, e in enumerate(eids):
            ei = int(e)
            if not (0 < ei < 100000000):
                continue
            if ei in seen:
                dups += 1
                continue
            seen.add(ei)
            keep_eids.append(ei)
            keep_cols.append(j)
        keep_cols = np.array(keep_cols, dtype=int)
        keep_eids = np.array(keep_eids)
        psum = (data[:nm6, :, 1][:, keep_cols].sum(axis=1)
                if len(keep_cols) else np.zeros(nm6))
        print(f"  {attr}: unique={len(keep_eids)} dup_rows={dups}  "
              f"eid {int(keep_eids.min()) if len(keep_eids) else '-'}.."
              f"{int(keep_eids.max()) if len(keep_eids) else '-'}")
        print(f"      cleaned %ESE/mode[:6]={np.round(psum, 1).tolist()}")
        # Split a doubled table by eid median to localize the double-count.
        if len(keep_eids) and np.max(psum) > 110:
            med = int(np.median(keep_eids))
            lo = keep_cols[keep_eids < med]
            hi = keep_cols[keep_eids >= med]
            lo_s = data[:nm6, :, 1][:, lo].sum(axis=1) if len(lo) else np.zeros(nm6)
            hi_s = data[:nm6, :, 1][:, hi].sum(axis=1) if len(hi) else np.zeros(nm6)
            print(f"      SPLIT eid<{med}: {np.round(lo_s, 1).tolist()}")
            print(f"      SPLIT eid>={med}: {np.round(hi_s, 1).tolist()}")
        grand = psum if grand is None else grand + psum
    if grand is not None:
        print(f"\n  >>> GRAND TOTAL deduped %ESE/mode[:6]={np.round(grand, 1).tolist()}")
        print("      (this should equal the GUI flat 'All Elements' view)")

    print("\nKey questions answered by the above:")
    print("  - dup_rows > 0 on a table     -> same-id duplicates (tool dedups them)")
    print("  - a table's SPLIT both ~equal -> energy in TWO id ranges = two reps")
    print("  - GRAND TOTAL still >100       -> double-count is distinct ids, not dups")


if __name__ == '__main__':
    if len(sys.argv) != 2:
        print("Usage: python debug_ese.py path/to/model.op2")
        sys.exit(1)
    main(sys.argv[1])
