#!/usr/bin/env python3
"""Convert a BDF model between unit systems.

Scales node coordinates, material properties, loads, and masses
according to the specified unit conversion.

Requires: pip install pyNastran

Usage:
    python convert_units.py input.bdf output.bdf --from mm-kg-s --to m-kg-s
    python convert_units.py input.bdf output.bdf --length-scale 0.001
    python convert_units.py input.bdf output.bdf --from mm-tonne-s --to m-kg-s
"""
import argparse
import sys
from pyNastran.bdf.bdf import BDF


# Predefined unit systems: (length_to_m, mass_to_kg, time_to_s)
UNIT_SYSTEMS = {
    'm-kg-s':      (1.0,       1.0,    1.0),
    'mm-kg-s':     (1e-3,      1.0,    1.0),
    'mm-tonne-s':  (1e-3,      1e3,    1.0),
    'mm-g-s':      (1e-3,      1e-3,   1.0),
    'in-lb-s':     (0.0254,    0.4536, 1.0),
    'in-lbf-s':    (0.0254,    0.4536, 1.0),
    'ft-lb-s':     (0.3048,    0.4536, 1.0),
    'in-slinch-s': (0.0254,    175.1268, 1.0),
    'cm-g-s':      (1e-2,      1e-3,   1.0),
}


def get_scale_factors(from_sys: str, to_sys: str):
    """Compute scale factors to convert from one unit system to another."""
    if from_sys not in UNIT_SYSTEMS:
        print(f"ERROR: Unknown unit system '{from_sys}'")
        print(f"Available: {', '.join(sorted(UNIT_SYSTEMS.keys()))}")
        sys.exit(1)
    if to_sys not in UNIT_SYSTEMS:
        print(f"ERROR: Unknown unit system '{to_sys}'")
        print(f"Available: {', '.join(sorted(UNIT_SYSTEMS.keys()))}")
        sys.exit(1)

    from_L, from_M, from_T = UNIT_SYSTEMS[from_sys]
    to_L, to_M, to_T = UNIT_SYSTEMS[to_sys]

    length_scale = from_L / to_L
    mass_scale = from_M / to_M
    time_scale = from_T / to_T

    return length_scale, mass_scale, time_scale


def convert_model(bdf_in: str, bdf_out: str,
                  length_scale: float, mass_scale: float,
                  time_scale: float) -> None:
    """Convert a BDF model with the given scale factors."""
    model = BDF()
    model.read_bdf(bdf_in)

    # Derived scales
    force_scale = mass_scale * length_scale / (time_scale ** 2)
    pressure_scale = force_scale / (length_scale ** 2)
    density_scale = mass_scale / (length_scale ** 3)
    stiffness_scale = pressure_scale  # same as stress (force/area)
    moment_scale = force_scale * length_scale
    inertia_scale = mass_scale * length_scale ** 2
    accel_scale = length_scale / (time_scale ** 2)

    print(f"Scale factors:")
    print(f"  Length:   {length_scale:.6e}")
    print(f"  Mass:     {mass_scale:.6e}")
    print(f"  Time:     {time_scale:.6e}")
    print(f"  Force:    {force_scale:.6e}")
    print(f"  Pressure: {pressure_scale:.6e}")
    print(f"  Density:  {density_scale:.6e}")

    # --- Scale nodes ---
    for nid, node in model.nodes.items():
        node.xyz = [c * length_scale for c in node.xyz]

    # --- Scale materials ---
    for mid, mat in model.materials.items():
        if mat.type == 'MAT1':
            if mat.e is not None:
                mat.e *= stiffness_scale
            if mat.g is not None:
                mat.g *= stiffness_scale
            if mat.rho != 0.:
                mat.rho *= density_scale
        elif mat.type == 'MAT8':
            mat.e1 *= stiffness_scale
            mat.e2 *= stiffness_scale
            mat.g12 *= stiffness_scale
            mat.g1z *= stiffness_scale
            mat.g2z *= stiffness_scale
            if mat.rho != 0.:
                mat.rho *= density_scale

    # --- Scale properties (thickness) ---
    for pid, prop in model.properties.items():
        if prop.type == 'PSHELL':
            if prop.t is not None:
                prop.t *= length_scale
        elif prop.type == 'PCOMP':
            for i in range(len(prop.thicknesses)):
                prop.thicknesses[i] *= length_scale
        elif prop.type in ('PBAR', 'PBEAM'):
            if hasattr(prop, 'A') and prop.A:
                prop.A *= length_scale ** 2
            if hasattr(prop, 'i1') and prop.i1:
                prop.i1 *= length_scale ** 4
            if hasattr(prop, 'i2') and prop.i2:
                prop.i2 *= length_scale ** 4
            if hasattr(prop, 'j') and prop.j:
                prop.j *= length_scale ** 4
        elif prop.type == 'PROD':
            if prop.A:
                prop.A *= length_scale ** 2
            if prop.j:
                prop.j *= length_scale ** 4

    # --- Scale loads ---
    for sid, load_list in model.loads.items():
        for load in load_list:
            if load.type == 'FORCE':
                load.mag *= force_scale
            elif load.type == 'MOMENT':
                load.mag *= moment_scale
            elif load.type == 'PLOAD4':
                load.pressures = [p * pressure_scale
                                  for p in load.pressures]
            elif load.type == 'GRAV':
                load.scale *= accel_scale

    # --- Scale masses ---
    for eid, mass_elem in model.masses.items():
        if hasattr(mass_elem, 'mass'):
            mass_elem.mass *= mass_scale
        if hasattr(mass_elem, 'X') and mass_elem.X is not None:
            mass_elem.X = [x * length_scale for x in mass_elem.X]

    model.write_bdf(bdf_out)
    print(f"\nConverted model written to {bdf_out}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description='Convert a BDF model between unit systems.')
    parser.add_argument('input_bdf', help='Input BDF file')
    parser.add_argument('output_bdf', help='Output BDF file')

    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument(
        '--from', dest='from_sys', default=None,
        help=('Source unit system. Options: '
              + ', '.join(sorted(UNIT_SYSTEMS.keys()))))
    group.add_argument(
        '--length-scale', type=float, default=None,
        help='Manual length scale factor (e.g., 0.001 for mm→m)')

    parser.add_argument(
        '--to', dest='to_sys', default=None,
        help='Target unit system')
    parser.add_argument(
        '--mass-scale', type=float, default=1.0,
        help='Manual mass scale factor (default: 1.0)')
    parser.add_argument(
        '--time-scale', type=float, default=1.0,
        help='Manual time scale factor (default: 1.0)')

    args = parser.parse_args()

    if args.from_sys and not args.to_sys:
        parser.error("--to is required when using --from")

    if args.from_sys:
        L, M, T = get_scale_factors(args.from_sys, args.to_sys)
    else:
        L = args.length_scale
        M = args.mass_scale
        T = args.time_scale

    convert_model(args.input_bdf, args.output_bdf, L, M, T)


if __name__ == '__main__':
    main()
