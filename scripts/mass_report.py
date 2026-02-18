#!/usr/bin/env python3
"""Generate a mass breakdown report by property ID.

Computes total mass, CG, and per-PID mass contributions from a BDF file.

Requires: pip install pyNastran

Usage:
    python mass_report.py model.bdf
    python mass_report.py model.bdf --csv mass_breakdown.csv
"""
import argparse
import sys
from collections import defaultdict
from pyNastran.bdf.bdf import BDF
from pyNastran.bdf.mesh_utils.mass_properties import mass_properties


def mass_report(bdf_filename: str, csv_output: str = None) -> None:
    model = BDF()
    model.read_bdf(bdf_filename)
    model.cross_reference()

    # --- Total mass ---
    total_mass, cg, inertia = mass_properties(model)

    sep = '=' * 65
    print(sep)
    print(f"MASS REPORT: {bdf_filename}")
    print(sep)
    print(f"  Total Mass = {total_mass:.6e}")
    print(f"  CG = ({cg[0]:.4e}, {cg[1]:.4e}, {cg[2]:.4e})")
    print(f"  Ixx={inertia[0]:.4e}  Iyy={inertia[1]:.4e}  "
          f"Izz={inertia[2]:.4e}")
    print(f"  Ixy={inertia[3]:.4e}  Ixz={inertia[4]:.4e}  "
          f"Iyz={inertia[5]:.4e}")

    # --- Per-PID breakdown ---
    pid_data = defaultdict(lambda: {'mass': 0.0, 'count': 0, 'type': ''})

    for eid, elem in model.elements.items():
        try:
            pid = elem.pid
        except AttributeError:
            continue
        try:
            emass = elem.Mass()
        except Exception:
            continue
        pid_data[pid]['mass'] += emass
        pid_data[pid]['count'] += 1
        if not pid_data[pid]['type']:
            prop = model.properties.get(pid)
            pid_data[pid]['type'] = prop.type if prop else '???'

    # Lumped masses
    lumped_mass = 0.0
    lumped_count = 0
    for eid, mass_elem in model.masses.items():
        try:
            lumped_mass += mass_elem.mass
            lumped_count += 1
        except AttributeError:
            pass

    print(f"\n{'PER-PID MASS BREAKDOWN':^65}")
    print('-' * 65)
    header = (f"  {'PID':>8s}  {'Type':>8s}  {'Elements':>8s}  "
              f"{'Mass':>14s}  {'% Total':>8s}")
    print(header)
    print(f"  {'-'*8}  {'-'*8}  {'-'*8}  {'-'*14}  {'-'*8}")

    rows = []
    for pid in sorted(pid_data.keys()):
        d = pid_data[pid]
        pct = 100. * d['mass'] / total_mass if total_mass > 0 else 0.
        row = (pid, d['type'], d['count'], d['mass'], pct)
        rows.append(row)
        print(f"  {pid:8d}  {d['type']:>8s}  {d['count']:8d}  "
              f"{d['mass']:14.6e}  {pct:7.2f}%")

    if lumped_count > 0:
        pct = 100. * lumped_mass / total_mass if total_mass > 0 else 0.
        rows.append(('CONM', 'lumped', lumped_count, lumped_mass, pct))
        print(f"  {'CONM':>8s}  {'lumped':>8s}  {lumped_count:8d}  "
              f"{lumped_mass:14.6e}  {pct:7.2f}%")

    # Sum check
    struct_mass = sum(d['mass'] for d in pid_data.values()) + lumped_mass
    print(f"  {'-'*8}  {'-'*8}  {'-'*8}  {'-'*14}  {'-'*8}")
    print(f"  {'SUM':>8s}  {'':>8s}  {'':>8s}  {struct_mass:14.6e}  "
          f"{'100.00%':>8s}")

    # --- CSV export ---
    if csv_output:
        with open(csv_output, 'w') as f:
            f.write("PID,Type,Elements,Mass,Percent\n")
            for row in rows:
                f.write(f"{row[0]},{row[1]},{row[2]},{row[3]:.6e},"
                        f"{row[4]:.2f}\n")
        print(f"\nCSV exported to {csv_output}")

    print(f"\n{sep}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description='Mass breakdown report by property ID.')
    parser.add_argument('bdf', help='Path to the BDF file')
    parser.add_argument('--csv', default=None,
                        help='Export breakdown to CSV file')
    args = parser.parse_args()
    mass_report(args.bdf, csv_output=args.csv)


if __name__ == '__main__':
    main()
