#!/usr/bin/env python3
"""Define composite layups with PCOMP and MAT8 cards.

Demonstrates:
- MAT8 orthotropic material definition
- PCOMP symmetric and asymmetric layup definition
- Querying ply properties
- Building a composite panel model

Usage:
    python composite_layup_definition.py output.bdf
"""
import sys
from pyNastran.bdf.bdf import BDF


def build_composite_panel(bdf_out: str) -> None:
    model = BDF()
    model.sol = 101

    # ---------------------------------------------------------------
    # MAT8: Carbon/Epoxy Unidirectional Tape
    # ---------------------------------------------------------------
    mid_tape = 1
    model.add_mat8(
        mid=mid_tape,
        E1=140.0e9,       # longitudinal modulus (Pa)
        E2=10.0e9,        # transverse modulus
        nu12=0.3,         # major Poisson's ratio
        G12=5.0e9,        # in-plane shear modulus
        G1z=5.0e9,        # transverse shear modulus (1-z plane)
        G2z=3.0e9,        # transverse shear modulus (2-z plane)
        rho=1600.0,       # density (kg/m³)
        a1=0.0e-6,        # CTE fiber direction (1/°C)
        a2=30.0e-6,       # CTE transverse direction
        tref=20.0,        # reference temperature
        Xt=2100.0e6,      # tensile strength, fiber dir
        Xc=1200.0e6,      # compressive strength, fiber dir
        Yt=50.0e6,        # tensile strength, transverse
        Yc=250.0e6,       # compressive strength, transverse
        S=70.0e6,         # in-plane shear strength
    )
    print(f"MAT8 {mid_tape}: CFRP UD Tape (E1={140e9:.1e})")

    # ---------------------------------------------------------------
    # MAT8: Glass/Epoxy Woven Fabric
    # ---------------------------------------------------------------
    mid_fabric = 2
    model.add_mat8(
        mid=mid_fabric,
        E1=25.0e9,
        E2=25.0e9,        # woven → quasi-isotropic in-plane
        nu12=0.12,
        G12=4.0e9,
        G1z=3.0e9,
        G2z=3.0e9,
        rho=1900.0,
        a1=12.0e-6,
        a2=12.0e-6,
        tref=20.0,
    )
    print(f"MAT8 {mid_fabric}: GFRP Woven Fabric (E1={25e9:.1e})")

    # ---------------------------------------------------------------
    # PCOMP: Quasi-Isotropic Layup [0/45/-45/90]_s (symmetric)
    # ---------------------------------------------------------------
    pid_qi = 10
    ply_t = 0.125e-3  # ply thickness (m)
    model.add_pcomp(
        pid=pid_qi,
        mids=[mid_tape] * 4,
        thicknesses=[ply_t] * 4,
        thetas=[0., 45., -45., 90.],
        souts=['YES'] * 4,
        lam='SYM',                     # symmetric → total 8 plies
        sb=0.,                         # allowable interlaminar shear
        ft='TSAI',                     # failure theory
        tref=120.,                     # cure temperature
        ge=0.0,                        # structural damping
    )
    print(f"PCOMP {pid_qi}: [0/45/-45/90]_s — "
          f"{4*2} plies, total t={ply_t*8*1e3:.2f} mm")

    # ---------------------------------------------------------------
    # PCOMP: Hybrid Layup (asymmetric, tape + fabric)
    # ---------------------------------------------------------------
    pid_hybrid = 20
    mids = [mid_fabric, mid_tape, mid_tape, mid_tape, mid_tape, mid_fabric]
    thicks = [0.25e-3, 0.125e-3, 0.125e-3, 0.125e-3, 0.125e-3, 0.25e-3]
    angles = [0., 0., 45., -45., 90., 0.]

    model.add_pcomp(
        pid=pid_hybrid,
        mids=mids,
        thicknesses=thicks,
        thetas=angles,
        souts=['YES'] * 6,
        lam=None,                      # asymmetric (no SYM)
    )
    total_t = sum(thicks)
    print(f"PCOMP {pid_hybrid}: Hybrid fabric/tape — "
          f"{len(mids)} plies, total t={total_t*1e3:.2f} mm")

    # ---------------------------------------------------------------
    # Query ply properties
    # ---------------------------------------------------------------
    pcomp = model.properties[pid_qi]
    print(f"\nPCOMP {pid_qi} details:")
    print(f"  nplies (defined) = {pcomp.nplies}")
    print(f"  is_symmetrical   = {pcomp.is_symmetrical}")

    for i in range(pcomp.nplies):
        print(f"  Ply {i+1}: MID={pcomp.material_ids[i]}, "
              f"t={pcomp.thicknesses[i]*1e3:.3f} mm, "
              f"θ={pcomp.thetas[i]:.1f}°")

    # ---------------------------------------------------------------
    # Build a simple panel mesh
    # ---------------------------------------------------------------
    nx, ny = 8, 4
    dx, dy = 0.05, 0.05  # 400mm x 200mm panel

    nid = 1
    node_grid = {}
    for j in range(ny + 1):
        for i in range(nx + 1):
            model.add_grid(nid=nid, xyz=[i * dx, j * dy, 0.0])
            node_grid[(i, j)] = nid
            nid += 1

    # Upper half: quasi-isotropic, lower half: hybrid
    eid = 1
    for j in range(ny):
        for i in range(nx):
            n1 = node_grid[(i, j)]
            n2 = node_grid[(i+1, j)]
            n3 = node_grid[(i+1, j+1)]
            n4 = node_grid[(i, j+1)]
            pid = pid_qi if j >= ny // 2 else pid_hybrid
            model.add_cquad4(eid=eid, pid=pid, nids=[n1, n2, n3, n4])
            eid += 1

    # Boundary conditions and load
    spc_sid = 1
    fixed_nodes = [node_grid[(0, j)] for j in range(ny + 1)]
    model.add_spc1(sid=spc_sid, components='123456', nodes=fixed_nodes)

    load_sid = 100
    tip_nodes = [node_grid[(nx, j)] for j in range(ny + 1)]
    for n in tip_nodes:
        model.add_force(sid=load_sid, node=n, mag=100.0,
                        xyz=[1., 0., 0.])

    # Case control
    cc = model.case_control_deck
    sc = cc.create_new_subcase(1)
    sc.add('SUBTITLE', 'Composite Panel — Tensile', options=[], option_type='')
    sc.add('LOAD', load_sid, options=[], option_type='')
    sc.add('SPC', spc_sid, options=[], option_type='')
    sc.add('DISPLACEMENT', 'ALL', options=['SORT1', 'REAL'],
           option_type='STRESS-type')
    sc.add('STRESS', 'ALL',
           options=['SORT1', 'REAL', 'VONMISES', 'BILIN'],
           option_type='STRESS-type')

    model.add_param('POST', [-1])

    model.validate()
    model.write_bdf(bdf_out)
    print(f"\nWritten composite panel model to {bdf_out}")


if __name__ == '__main__':
    if len(sys.argv) != 2:
        print("Usage: python composite_layup_definition.py <output.bdf>")
        sys.exit(1)
    build_composite_panel(sys.argv[1])
