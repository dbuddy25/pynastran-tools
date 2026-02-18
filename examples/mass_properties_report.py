#!/usr/bin/env python3
"""Generate a mass properties report from a BDF model.

Reports total mass, center of gravity, moments of inertia, and
optionally a breakdown by property ID.

Usage:
    python mass_properties_report.py model.bdf
"""
import sys
import numpy as np
from pyNastran.bdf.bdf import BDF
from pyNastran.bdf.mesh_utils.mass_properties import (
    mass_properties,
    mass_properties_no_xref,
)


def print_separator(char: str = '-', width: int = 60) -> None:
    print(char * width)


def mass_report(bdf_filename: str) -> None:
    model = BDF()
    model.read_bdf(bdf_filename)
    model.cross_reference()

    # --- Total mass properties ---
    mass, cg, inertia = mass_properties(model)

    print_separator('=')
    print("MASS PROPERTIES REPORT")
    print_separator('=')
    print(f"  Model: {bdf_filename}")
    print(f"  Nodes:    {len(model.nodes)}")
    print(f"  Elements: {len(model.elements)}")
    print(f"  Masses:   {len(model.masses)}")

    print_separator()
    print("TOTAL MASS")
    print_separator()
    print(f"  Mass = {mass:.6e}")

    print_separator()
    print("CENTER OF GRAVITY")
    print_separator()
    print(f"  CG_x = {cg[0]:.6e}")
    print(f"  CG_y = {cg[1]:.6e}")
    print(f"  CG_z = {cg[2]:.6e}")

    print_separator()
    print("MOMENTS OF INERTIA (about CG)")
    print_separator()
    # inertia = [Ixx, Iyy, Izz, Ixy, Ixz, Iyz]
    print(f"  Ixx = {inertia[0]:.6e}")
    print(f"  Iyy = {inertia[1]:.6e}")
    print(f"  Izz = {inertia[2]:.6e}")
    print(f"  Ixy = {inertia[3]:.6e}")
    print(f"  Ixz = {inertia[4]:.6e}")
    print(f"  Iyz = {inertia[5]:.6e}")

    # --- Breakdown by property ID ---
    print_separator('=')
    print("MASS BREAKDOWN BY PROPERTY ID")
    print_separator('=')

    pid_masses = {}
    for eid, elem in model.elements.items():
        try:
            pid = elem.pid
        except AttributeError:
            continue
        try:
            elem_mass = elem.Mass()
        except Exception:
            continue
        pid_masses.setdefault(pid, 0.0)
        pid_masses[pid] += elem_mass

    # Add CONM2 / lumped masses
    lumped_mass = 0.0
    for eid, mass_elem in model.masses.items():
        try:
            lumped_mass += mass_elem.mass
        except AttributeError:
            pass

    print(f"\n  {'PID':>8s}  {'Type':>10s}  {'Mass':>14s}  {'% Total':>8s}")
    print(f"  {'-'*8}  {'-'*10}  {'-'*14}  {'-'*8}")

    total_check = 0.0
    for pid in sorted(pid_masses.keys()):
        prop = model.properties.get(pid)
        ptype = prop.type if prop else '???'
        pmass = pid_masses[pid]
        pct = 100.0 * pmass / mass if mass > 0 else 0.0
        total_check += pmass
        print(f"  {pid:8d}  {ptype:>10s}  {pmass:14.6e}  {pct:7.2f}%")

    if lumped_mass > 0.0:
        pct = 100.0 * lumped_mass / mass if mass > 0 else 0.0
        total_check += lumped_mass
        print(f"  {'CONM2':>8s}  {'lumped':>10s}  "
              f"{lumped_mass:14.6e}  {pct:7.2f}%")

    print(f"  {'-'*8}  {'-'*10}  {'-'*14}  {'-'*8}")
    print(f"  {'TOTAL':>8s}  {'':>10s}  {total_check:14.6e}  100.00%")

    # --- Check against lumped total ---
    diff = abs(mass - total_check)
    if diff > 1e-6 * mass:
        print(f"\n  WARNING: Mass mismatch of {diff:.6e} "
              f"({100*diff/mass:.2f}%)")
        print("  This can happen due to NSM, offsets, or coordinate effects.")


if __name__ == '__main__':
    if len(sys.argv) != 2:
        print("Usage: python mass_properties_report.py <model.bdf>")
        sys.exit(1)
    mass_report(sys.argv[1])
