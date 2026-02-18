#!/usr/bin/env python3
"""Print a summary of a Nastran BDF model.

Reports node/element counts, element types, properties, materials,
loads, constraints, and coordinate systems.

Requires: pip install pyNastran

Usage:
    python bdf_summary.py model.bdf
    python bdf_summary.py model.bdf --verbose
"""
import argparse
import sys
from collections import Counter
from pyNastran.bdf.bdf import BDF


def summarize(bdf_filename: str, verbose: bool = False) -> None:
    model = BDF()
    model.read_bdf(bdf_filename)

    sep = '=' * 60

    print(sep)
    print(f"BDF SUMMARY: {bdf_filename}")
    print(sep)

    # --- Counts ---
    print(f"\n  SOL          = {model.sol}")
    print(f"  Nodes        = {len(model.nodes):,}")
    print(f"  Elements     = {len(model.elements):,}")
    print(f"  Properties   = {len(model.properties):,}")
    print(f"  Materials    = {len(model.materials):,}")
    print(f"  Coord Systems= {len(model.coords):,}")
    print(f"  Rigid Elems  = {len(model.rigid_elements):,}")
    print(f"  Mass Elems   = {len(model.masses):,}")
    print(f"  Load Sets    = {len(model.loads):,}")
    print(f"  SPC Sets     = {len(model.spcs):,}")
    print(f"  MPC Sets     = {len(model.mpcs):,}")

    # --- Element type breakdown ---
    elem_types = Counter()
    for eid, elem in model.elements.items():
        elem_types[elem.type] += 1

    print(f"\n{'ELEMENT TYPES':^60}")
    print('-' * 60)
    for etype, count in sorted(elem_types.items(), key=lambda x: -x[1]):
        print(f"  {etype:<16s}  {count:>10,}")

    # --- Property type breakdown ---
    prop_types = Counter()
    for pid, prop in model.properties.items():
        prop_types[prop.type] += 1

    if prop_types:
        print(f"\n{'PROPERTY TYPES':^60}")
        print('-' * 60)
        for ptype, count in sorted(prop_types.items(), key=lambda x: -x[1]):
            print(f"  {ptype:<16s}  {count:>10,}")

    # --- Material type breakdown ---
    mat_types = Counter()
    for mid, mat in model.materials.items():
        mat_types[mat.type] += 1

    if mat_types:
        print(f"\n{'MATERIAL TYPES':^60}")
        print('-' * 60)
        for mtype, count in sorted(mat_types.items(), key=lambda x: -x[1]):
            print(f"  {mtype:<16s}  {count:>10,}")

    # --- Rigid element breakdown ---
    if model.rigid_elements:
        rigid_types = Counter()
        for eid, elem in model.rigid_elements.items():
            rigid_types[elem.type] += 1
        print(f"\n{'RIGID ELEMENTS':^60}")
        print('-' * 60)
        for rtype, count in sorted(rigid_types.items(), key=lambda x: -x[1]):
            print(f"  {rtype:<16s}  {count:>10,}")

    # --- Load breakdown ---
    if model.loads:
        print(f"\n{'LOAD SETS':^60}")
        print('-' * 60)
        for sid in sorted(model.loads.keys()):
            load_list = model.loads[sid]
            ltypes = Counter(load.type for load in load_list)
            desc = ', '.join(f"{t}×{c}" for t, c in sorted(ltypes.items()))
            print(f"  SID {sid:<8d}  {desc}")

    # --- Verbose: list properties and materials ---
    if verbose:
        if model.properties:
            print(f"\n{'PROPERTIES (DETAIL)':^60}")
            print('-' * 60)
            for pid in sorted(model.properties.keys()):
                prop = model.properties[pid]
                if prop.type == 'PSHELL':
                    print(f"  PID {pid}: PSHELL mid={prop.mid1} "
                          f"t={prop.t}")
                elif prop.type == 'PCOMP':
                    print(f"  PID {pid}: PCOMP nplies={prop.nplies}")
                elif prop.type == 'PSOLID':
                    print(f"  PID {pid}: PSOLID mid={prop.mid}")
                else:
                    print(f"  PID {pid}: {prop.type}")

        if model.materials:
            print(f"\n{'MATERIALS (DETAIL)':^60}")
            print('-' * 60)
            for mid in sorted(model.materials.keys()):
                mat = model.materials[mid]
                if mat.type == 'MAT1':
                    print(f"  MID {mid}: MAT1 E={mat.e:.3e} "
                          f"nu={mat.nu} rho={mat.rho}")
                elif mat.type == 'MAT8':
                    print(f"  MID {mid}: MAT8 E1={mat.e1:.3e} "
                          f"E2={mat.e2:.3e} rho={mat.rho}")
                else:
                    print(f"  MID {mid}: {mat.type}")

    print(f"\n{sep}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description='Print summary statistics for a Nastran BDF file.')
    parser.add_argument('bdf', help='Path to the BDF file')
    parser.add_argument('-v', '--verbose', action='store_true',
                        help='Show detailed property/material info')
    args = parser.parse_args()
    summarize(args.bdf, verbose=args.verbose)


if __name__ == '__main__':
    main()
