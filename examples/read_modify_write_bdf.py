#!/usr/bin/env python3
"""Read a BDF, modify properties and loads, write back out.

Demonstrates the standard read → cross-reference → modify →
un-cross-reference → write workflow.

Usage:
    python read_modify_write_bdf.py input.bdf output.bdf
"""
import sys
from pyNastran.bdf.bdf import BDF


def main(bdf_in: str, bdf_out: str) -> None:
    # --- Read ---
    model = BDF()
    model.read_bdf(bdf_in)

    # --- Cross-reference (enables _ref access) ---
    model.cross_reference()

    print(f"Nodes:    {len(model.nodes)}")
    print(f"Elements: {len(model.elements)}")
    print(f"Properties: {len(model.properties)}")
    print(f"Materials: {len(model.materials)}")

    # --- Modify shell thicknesses ---
    # Double the thickness of all PSHELL properties
    for pid, prop in model.properties.items():
        if prop.type == 'PSHELL':
            old_t = prop.t
            prop.t *= 2.0
            print(f"  PSHELL {pid}: t {old_t:.4f} -> {prop.t:.4f}")

    # --- Modify material properties ---
    # Increase Young's modulus by 10% for all MAT1
    for mid, mat in model.materials.items():
        if mat.type == 'MAT1':
            mat.e *= 1.10
            print(f"  MAT1 {mid}: E -> {mat.e:.3e}")

    # --- Scale all FORCE loads by 1.5x ---
    for sid, load_list in model.loads.items():
        for load in load_list:
            if load.type == 'FORCE':
                load.mag *= 1.5
                print(f"  FORCE sid={sid} node={load.node}: "
                      f"mag -> {load.mag:.1f}")

    # --- Add a new SPC constraint ---
    # Fix node 1 in all 6 DOFs (if it exists)
    if 1 in model.nodes:
        model.add_spc1(sid=999, components='123456', nodes=[1])
        print("  Added SPC1 on node 1 (all DOFs)")

    # --- Un-cross-reference before writing ---
    model.uncross_reference()

    # --- Write ---
    model.write_bdf(bdf_out)
    print(f"\nWritten to {bdf_out}")


if __name__ == '__main__':
    if len(sys.argv) != 3:
        print("Usage: python read_modify_write_bdf.py <input.bdf> <output.bdf>")
        sys.exit(1)
    main(sys.argv[1], sys.argv[2])
