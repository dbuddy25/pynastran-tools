#!/usr/bin/env python3
"""Extract displacement results from an OP2 file.

Demonstrates displacement extraction for:
- Static analysis (SOL 101): single time step
- Modal analysis (SOL 103): mode shapes

Usage:
    python extract_op2_displacements.py model.op2
"""
import sys
import numpy as np
from pyNastran.op2.op2 import OP2


def extract_static_displacements(op2_obj: OP2) -> None:
    """Extract displacements from a static analysis."""
    print("=" * 60)
    print("STATIC DISPLACEMENTS")
    print("=" * 60)

    for subcase_id, disp in sorted(op2_obj.displacements.items()):
        node_ids = disp.node_gridtype[:, 0]
        data = disp.data  # (ntimes, nnodes, 6): T1 T2 T3 R1 R2 R3

        print(f"\nSubcase {subcase_id}: {len(node_ids)} nodes, "
              f"{data.shape[0]} time steps")

        # For static, typically 1 time step
        for itime in range(data.shape[0]):
            t1 = data[itime, :, 0]
            t2 = data[itime, :, 1]
            t3 = data[itime, :, 2]
            magnitude = np.sqrt(t1**2 + t2**2 + t3**2)

            # Find max displacement
            idx_max = np.argmax(magnitude)
            max_nid = node_ids[idx_max]

            print(f"\n  Time step {itime}:")
            print(f"    Max |disp|   = {magnitude[idx_max]:.6e} "
                  f"at node {max_nid}")
            print(f"    Max T1       = {np.max(np.abs(t1)):.6e}")
            print(f"    Max T2       = {np.max(np.abs(t2)):.6e}")
            print(f"    Max T3       = {np.max(np.abs(t3)):.6e}")

            # Print top 5 displaced nodes
            top5_idx = np.argsort(magnitude)[-5:][::-1]
            print("\n    Top 5 displaced nodes:")
            print(f"    {'Node':>8s}  {'T1':>12s}  {'T2':>12s}  "
                  f"{'T3':>12s}  {'|Disp|':>12s}")
            for idx in top5_idx:
                print(f"    {node_ids[idx]:8d}  {t1[idx]:12.5e}  "
                      f"{t2[idx]:12.5e}  {t3[idx]:12.5e}  "
                      f"{magnitude[idx]:12.5e}")


def extract_modal_displacements(op2_obj: OP2) -> None:
    """Extract mode shapes from a modal analysis."""
    print("\n" + "=" * 60)
    print("MODAL ANALYSIS — EIGENVECTORS")
    print("=" * 60)

    # Eigenvalue table
    for title, eigval_table in op2_obj.eigenvalues.items():
        modes = eigval_table.mode
        freqs = eigval_table.freq

        print(f"\nEigenvalue table: '{title}'")
        print(f"  {'Mode':>4s}  {'Freq (Hz)':>12s}  {'Radians':>12s}")
        for i in range(len(modes)):
            print(f"  {modes[i]:4d}  {freqs[i]:12.4f}  "
                  f"{eigval_table.radians[i]:12.4f}")

    # Mode shapes (eigenvectors)
    for subcase_id, eigvec in sorted(op2_obj.eigenvectors.items()):
        node_ids = eigvec.node_gridtype[:, 0]
        data = eigvec.data  # (nmodes, nnodes, 6)
        nmodes = data.shape[0]

        print(f"\nSubcase {subcase_id}: {nmodes} modes, "
              f"{len(node_ids)} nodes")

        for imode in range(min(nmodes, 5)):  # first 5 modes
            t3 = data[imode, :, 2]  # T3 component
            magnitude = np.sqrt(
                data[imode, :, 0]**2 +
                data[imode, :, 1]**2 +
                data[imode, :, 2]**2
            )
            idx_max = np.argmax(np.abs(magnitude))

            print(f"\n  Mode {imode + 1}:")
            print(f"    Max |eigvec| = {magnitude[idx_max]:.6e} "
                  f"at node {node_ids[idx_max]}")
            print(f"    Max |T3|     = {np.max(np.abs(t3)):.6e}")


def extract_specific_nodes(op2_obj: OP2, node_ids_of_interest: list) -> None:
    """Extract displacements for specific nodes."""
    print("\n" + "=" * 60)
    print("SPECIFIC NODE DISPLACEMENTS")
    print("=" * 60)

    for subcase_id, disp in sorted(op2_obj.displacements.items()):
        node_ids = disp.node_gridtype[:, 0]
        data = disp.data

        print(f"\nSubcase {subcase_id}:")
        for nid in node_ids_of_interest:
            indices = np.where(node_ids == nid)[0]
            if len(indices) == 0:
                print(f"  Node {nid}: NOT FOUND")
                continue
            idx = indices[0]
            for itime in range(data.shape[0]):
                vals = data[itime, idx, :]
                print(f"  Node {nid} [t={itime}]: "
                      f"T1={vals[0]:.5e}  T2={vals[1]:.5e}  "
                      f"T3={vals[2]:.5e}  "
                      f"R1={vals[3]:.5e}  R2={vals[4]:.5e}  "
                      f"R3={vals[5]:.5e}")


def main(op2_filename: str) -> None:
    op2 = OP2()
    op2.read_op2(op2_filename)

    # Static displacements
    if op2.displacements:
        extract_static_displacements(op2)

    # Modal analysis
    if op2.eigenvectors:
        extract_modal_displacements(op2)

    # Export to CSV via DataFrame
    if op2.displacements:
        for subcase_id, disp in op2.displacements.items():
            csv_name = f"displacements_sc{subcase_id}.csv"
            df = disp.data_frame
            df.to_csv(csv_name)
            print(f"\nExported {csv_name}")


if __name__ == '__main__':
    if len(sys.argv) != 2:
        print("Usage: python extract_op2_displacements.py <model.op2>")
        sys.exit(1)
    main(sys.argv[1])
