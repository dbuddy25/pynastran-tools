#!/usr/bin/env python3
"""Extract OP2 results to CSV files.

Exports displacements, stresses, and/or forces from a Nastran OP2 file
to CSV format for post-processing in Excel, MATLAB, or Python.

Requires: pip install pyNastran

Usage:
    python extract_results.py model.op2
    python extract_results.py model.op2 --types displacement stress
    python extract_results.py model.op2 --subcase 1 --output-dir results/
"""
import argparse
import os
import sys
from pyNastran.op2.op2 import OP2


RESULT_TYPES = {
    'displacement': 'displacements',
    'velocity': 'velocities',
    'acceleration': 'accelerations',
    'spc_force': 'spc_forces',
    'load_vector': 'load_vectors',
    'cquad4_stress': 'cquad4_stress',
    'ctria3_stress': 'ctria3_stress',
    'chexa_stress': 'chexa_stress',
    'cpenta_stress': 'cpenta_stress',
    'ctetra_stress': 'ctetra_stress',
    'cbar_stress': 'cbar_stress',
    'cbar_force': 'cbar_force',
    'cbeam_stress': 'cbeam_stress',
    'cbeam_force': 'cbeam_force',
    'eigenvector': 'eigenvectors',
}

# Shortcut aliases
ALIASES = {
    'stress': ['cquad4_stress', 'ctria3_stress', 'chexa_stress',
               'cpenta_stress', 'ctetra_stress', 'cbar_stress',
               'cbeam_stress'],
    'force': ['cbar_force', 'cbeam_force'],
    'all': list(RESULT_TYPES.keys()),
}


def export_results(op2_filename: str, types: list, subcase: int,
                   output_dir: str) -> None:
    os.makedirs(output_dir, exist_ok=True)

    op2 = OP2()
    op2.read_op2(op2_filename)

    # Expand aliases
    expanded_types = []
    for t in types:
        if t in ALIASES:
            expanded_types.extend(ALIASES[t])
        else:
            expanded_types.append(t)

    exported = 0
    for type_name in expanded_types:
        attr_name = RESULT_TYPES.get(type_name)
        if attr_name is None:
            print(f"WARNING: Unknown result type '{type_name}', skipping")
            continue

        result_dict = getattr(op2, attr_name, {})
        if not result_dict:
            continue

        for sc_id, result in sorted(result_dict.items()):
            if subcase is not None and sc_id != subcase:
                continue

            csv_name = f"{type_name}_sc{sc_id}.csv"
            csv_path = os.path.join(output_dir, csv_name)

            df = result.data_frame
            df.to_csv(csv_path)
            print(f"  Exported: {csv_path} "
                  f"({df.shape[0]} rows x {df.shape[1]} cols)")
            exported += 1

    # Eigenvalues (special: keyed by string title)
    if 'eigenvector' in expanded_types and op2.eigenvalues:
        for title, eigval_table in op2.eigenvalues.items():
            csv_name = "eigenvalues.csv"
            csv_path = os.path.join(output_dir, csv_name)
            with open(csv_path, 'w') as f:
                f.write("mode,eigenvalue,radians,freq_hz,"
                        "generalized_mass,generalized_stiffness\n")
                for i in range(len(eigval_table.mode)):
                    f.write(f"{eigval_table.mode[i]},"
                            f"{eigval_table.eigenvalue[i]},"
                            f"{eigval_table.radians[i]},"
                            f"{eigval_table.freq[i]},"
                            f"{eigval_table.generalized_mass[i]},"
                            f"{eigval_table.generalized_stiffness[i]}\n")
            print(f"  Exported: {csv_path}")
            exported += 1

    if exported == 0:
        print("No matching results found to export.")
    else:
        print(f"\nExported {exported} file(s) to {output_dir}/")


def main() -> None:
    parser = argparse.ArgumentParser(
        description='Extract OP2 results to CSV files.')
    parser.add_argument('op2', help='Path to the OP2 file')
    parser.add_argument(
        '--types', '-t', nargs='+', default=['all'],
        help=('Result types to extract. Options: '
              + ', '.join(sorted(RESULT_TYPES.keys()))
              + '. Aliases: stress, force, all (default: all)'))
    parser.add_argument('--subcase', '-s', type=int, default=None,
                        help='Extract only this subcase (default: all)')
    parser.add_argument('--output-dir', '-o', default='.',
                        help='Output directory (default: current)')
    args = parser.parse_args()
    export_results(args.op2, args.types, args.subcase, args.output_dir)


if __name__ == '__main__':
    main()
