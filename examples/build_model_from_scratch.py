#!/usr/bin/env python3
"""Build a simple plate model from scratch using pyNastran.

Creates a flat rectangular plate meshed with CQUAD4 elements,
applies a uniform pressure load and fixed boundary conditions,
and writes a SOL 101 BDF ready for analysis.

Usage:
    python build_model_from_scratch.py output.bdf
"""
import sys
from pyNastran.bdf.bdf import BDF


def build_plate(bdf_out: str,
                length: float = 1.0,
                width: float = 0.5,
                nx: int = 10,
                ny: int = 5,
                thickness: float = 0.005,
                E: float = 2.1e11,
                nu: float = 0.3,
                rho: float = 7850.0,
                pressure: float = 1e5) -> None:
    """Build a rectangular plate model.

    Args:
        bdf_out: Output BDF filename.
        length: Plate length in x-direction.
        width: Plate width in y-direction.
        nx: Number of elements in x.
        ny: Number of elements in y.
        thickness: Shell thickness.
        E: Young's modulus.
        nu: Poisson's ratio.
        rho: Density.
        pressure: Uniform pressure load.
    """
    model = BDF()
    model.sol = 101  # linear static

    # --- Material ---
    mid = 1
    model.add_mat1(mid=mid, E=E, G=None, nu=nu, rho=rho)

    # --- Property ---
    pid = 1
    model.add_pshell(pid=pid, mid1=mid, t=thickness)

    # --- Grid points ---
    nid = 1
    dx = length / nx
    dy = width / ny

    node_grid = {}  # (i, j) -> nid mapping
    for j in range(ny + 1):
        for i in range(nx + 1):
            x = i * dx
            y = j * dy
            model.add_grid(nid=nid, xyz=[x, y, 0.0])
            node_grid[(i, j)] = nid
            nid += 1

    # --- CQUAD4 elements ---
    eid = 1
    for j in range(ny):
        for i in range(nx):
            n1 = node_grid[(i, j)]
            n2 = node_grid[(i + 1, j)]
            n3 = node_grid[(i + 1, j + 1)]
            n4 = node_grid[(i, j + 1)]
            model.add_cquad4(eid=eid, pid=pid, nids=[n1, n2, n3, n4])
            eid += 1

    # --- Fixed boundary (x=0 edge) ---
    spc_sid = 1
    fixed_nodes = [node_grid[(0, j)] for j in range(ny + 1)]
    model.add_spc1(sid=spc_sid, components='123456', nodes=fixed_nodes)

    # --- Pressure load on all elements ---
    load_sid = 100
    all_eids = list(range(1, eid))
    model.add_pload4(sid=load_sid, eids=all_eids,
                     pressures=[pressure, pressure, pressure, pressure])

    # --- Gravity load ---
    grav_sid = 200
    model.add_grav(sid=grav_sid, scale=9.81, N=[0., 0., -1.])

    # --- Combined load ---
    combined_sid = 999
    model.add_load(sid=combined_sid, scale=1.0,
                   scale_factors=[1.0, 1.0],
                   load_ids=[load_sid, grav_sid])

    # --- Case control ---
    cc = model.case_control_deck
    subcase = cc.create_new_subcase(1)
    subcase.add('SUBTITLE', 'Pressure + Gravity', options=[], option_type='')
    subcase.add('LOAD', combined_sid, options=[], option_type='')
    subcase.add('SPC', spc_sid, options=[], option_type='')
    subcase.add('DISPLACEMENT', 'ALL', options=['SORT1', 'REAL'],
                option_type='STRESS-type')
    subcase.add('STRESS', 'ALL',
                options=['SORT1', 'REAL', 'VONMISES', 'BILIN'],
                option_type='STRESS-type')
    subcase.add('SPCFORCES', 'ALL', options=['SORT1', 'REAL'],
                option_type='STRESS-type')

    # --- Params ---
    model.add_param('POST', [-1])
    model.add_param('AUTOSPC', ['YES'])
    model.add_param('PRTMAXIM', ['YES'])

    # --- Validate and write ---
    model.validate()
    model.write_bdf(bdf_out)

    total_nodes = (nx + 1) * (ny + 1)
    total_elems = nx * ny
    print(f"Built plate model: {total_nodes} nodes, {total_elems} CQUAD4s")
    print(f"  Size: {length} x {width}, t={thickness}")
    print(f"  Pressure: {pressure:.1e}, Gravity: 9.81 m/s²")
    print(f"  Written to {bdf_out}")


if __name__ == '__main__':
    if len(sys.argv) != 2:
        print("Usage: python build_model_from_scratch.py <output.bdf>")
        sys.exit(1)
    build_plate(sys.argv[1])
