#!/usr/bin/env python3
"""Extract modal analysis results (SOL 103) from an OP2 file.

Demonstrates:
- Eigenvalue table extraction (frequencies, generalized mass/stiffness)
- Mode shape extraction and normalization
- Modal effective mass computation
- Export to CSV

Usage:
    python modal_results_extraction.py model.op2
"""
import sys
import numpy as np
from pyNastran.op2.op2 import OP2


def extract_eigenvalues(op2_obj: OP2) -> None:
    """Print the eigenvalue summary table."""
    print("=" * 70)
    print("EIGENVALUE SUMMARY")
    print("=" * 70)

    # eigenvalues keys are STRINGS (table titles), not ints
    for title, eigval_table in op2_obj.eigenvalues.items():
        modes = eigval_table.mode
        freqs = eigval_table.freq
        radians = eigval_table.radians
        gen_mass = eigval_table.generalized_mass
        gen_stiff = eigval_table.generalized_stiffness

        print(f"\nTable: '{title}'")
        print(f"  Number of modes: {len(modes)}")

        header = (f"  {'Mode':>4s}  {'Freq (Hz)':>12s}  "
                  f"{'Omega (rad/s)':>14s}  {'Gen Mass':>12s}  "
                  f"{'Gen Stiff':>12s}")
        print(header)
        print("  " + "-" * (len(header) - 2))

        for i in range(len(modes)):
            print(f"  {modes[i]:4d}  {freqs[i]:12.4f}  "
                  f"{radians[i]:14.4f}  {gen_mass[i]:12.4e}  "
                  f"{gen_stiff[i]:12.4e}")


def extract_mode_shapes(op2_obj: OP2, num_modes: int = 5) -> None:
    """Extract and summarize mode shapes."""
    print("\n" + "=" * 70)
    print("MODE SHAPES")
    print("=" * 70)

    for subcase_id, eigvec in sorted(op2_obj.eigenvectors.items()):
        node_ids = eigvec.node_gridtype[:, 0]
        data = eigvec.data  # (nmodes, nnodes, 6)
        nmodes = data.shape[0]

        print(f"\nSubcase {subcase_id}: {nmodes} modes, "
              f"{len(node_ids)} nodes")

        for imode in range(min(nmodes, num_modes)):
            mode_data = data[imode, :, :]  # (nnodes, 6)

            # Translational magnitude
            trans_mag = np.sqrt(
                mode_data[:, 0]**2 +
                mode_data[:, 1]**2 +
                mode_data[:, 2]**2
            )

            # Dominant direction
            max_t1 = np.max(np.abs(mode_data[:, 0]))
            max_t2 = np.max(np.abs(mode_data[:, 1]))
            max_t3 = np.max(np.abs(mode_data[:, 2]))
            max_r1 = np.max(np.abs(mode_data[:, 3]))
            max_r2 = np.max(np.abs(mode_data[:, 4]))
            max_r3 = np.max(np.abs(mode_data[:, 5]))

            directions = ['T1', 'T2', 'T3', 'R1', 'R2', 'R3']
            maxes = [max_t1, max_t2, max_t3, max_r1, max_r2, max_r3]
            dominant = directions[np.argmax(maxes)]

            idx_max = np.argmax(trans_mag)

            print(f"\n  Mode {imode + 1}:")
            print(f"    Max |displacement| = {trans_mag[idx_max]:.6e} "
                  f"at node {node_ids[idx_max]}")
            print(f"    Dominant direction  = {dominant}")
            print(f"    Max components: T1={max_t1:.4e} T2={max_t2:.4e} "
                  f"T3={max_t3:.4e}")
            print(f"                   R1={max_r1:.4e} R2={max_r2:.4e} "
                  f"R3={max_r3:.4e}")


def compute_modal_effective_mass(op2_obj: OP2) -> None:
    """Compute modal effective mass from eigenvectors and mass.

    This is an approximation using the eigenvector components.
    For exact values, use Nastran's MEFFMASS output.
    """
    print("\n" + "=" * 70)
    print("MODAL PARTICIPATION (APPROXIMATE)")
    print("=" * 70)

    for subcase_id, eigvec in sorted(op2_obj.eigenvectors.items()):
        data = eigvec.data  # (nmodes, nnodes, 6)
        nmodes = data.shape[0]

        print(f"\nSubcase {subcase_id}:")
        print(f"  {'Mode':>4s}  {'T1 %':>8s}  {'T2 %':>8s}  "
              f"{'T3 %':>8s}  {'Dominant':>10s}")
        print(f"  {'-'*4}  {'-'*8}  {'-'*8}  {'-'*8}  {'-'*10}")

        for imode in range(nmodes):
            # Sum of squared components as proxy for participation
            t1_sum = np.sum(data[imode, :, 0] ** 2)
            t2_sum = np.sum(data[imode, :, 1] ** 2)
            t3_sum = np.sum(data[imode, :, 2] ** 2)
            total = t1_sum + t2_sum + t3_sum

            if total > 0:
                t1_pct = 100 * t1_sum / total
                t2_pct = 100 * t2_sum / total
                t3_pct = 100 * t3_sum / total
            else:
                t1_pct = t2_pct = t3_pct = 0.0

            maxp = max(t1_pct, t2_pct, t3_pct)
            if maxp == t1_pct:
                dom = 'X (T1)'
            elif maxp == t2_pct:
                dom = 'Y (T2)'
            else:
                dom = 'Z (T3)'

            print(f"  {imode+1:4d}  {t1_pct:7.1f}%  {t2_pct:7.1f}%  "
                  f"{t3_pct:7.1f}%  {dom:>10s}")


def export_mode_shapes_csv(op2_obj: OP2, prefix: str = 'mode') -> None:
    """Export mode shapes to CSV files."""
    for subcase_id, eigvec in op2_obj.eigenvectors.items():
        csv_name = f"{prefix}_sc{subcase_id}.csv"
        df = eigvec.data_frame
        df.to_csv(csv_name)
        print(f"\nExported mode shapes to {csv_name}")


def main(op2_filename: str) -> None:
    op2 = OP2()
    op2.read_op2(op2_filename)

    if op2.eigenvalues:
        extract_eigenvalues(op2)
    else:
        print("No eigenvalue tables found — is this a SOL 103 run?")
        return

    if op2.eigenvectors:
        extract_mode_shapes(op2, num_modes=10)
        compute_modal_effective_mass(op2)
        export_mode_shapes_csv(op2)


if __name__ == '__main__':
    if len(sys.argv) != 2:
        print("Usage: python modal_results_extraction.py <model.op2>")
        sys.exit(1)
    main(sys.argv[1])
