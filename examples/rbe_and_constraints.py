#!/usr/bin/env python3
"""Demonstrate RBE2, RBE3, SPC, MPC, and CONM2 usage.

Creates a simple model showing:
- RBE2 rigid connections (e.g., bolt pattern)
- RBE3 load distribution (e.g., distributed load application)
- CONM2 concentrated masses
- SPC/SPC1 boundary conditions
- MPC multi-point constraints

Usage:
    python rbe_and_constraints.py output.bdf
"""
import sys
import math
from pyNastran.bdf.bdf import BDF


def build_model(bdf_out: str) -> None:
    model = BDF()
    model.sol = 101

    # --- Material and property ---
    model.add_mat1(mid=1, E=2.1e11, G=None, nu=0.3, rho=7850.)
    model.add_pshell(pid=1, mid1=1, t=0.003)

    # ---------------------------------------------------------------
    # Create a small plate mesh (10x10 elements)
    # ---------------------------------------------------------------
    nx, ny = 10, 10
    lx, ly = 1.0, 1.0
    dx, dy = lx / nx, ly / ny

    nid = 1
    node_grid = {}
    for j in range(ny + 1):
        for i in range(nx + 1):
            model.add_grid(nid=nid, xyz=[i * dx, j * dy, 0.0])
            node_grid[(i, j)] = nid
            nid += 1

    eid = 1
    for j in range(ny):
        for i in range(nx):
            n1 = node_grid[(i, j)]
            n2 = node_grid[(i+1, j)]
            n3 = node_grid[(i+1, j+1)]
            n4 = node_grid[(i, j+1)]
            model.add_cquad4(eid=eid, pid=1, nids=[n1, n2, n3, n4])
            eid += 1

    next_nid = nid  # track next available node ID
    next_eid = eid  # track next available element ID

    # ---------------------------------------------------------------
    # RBE2: Rigid connection — simulate a bolt at center of plate
    # ---------------------------------------------------------------
    # Independent node (bolt center) — new node above the plate
    bolt_nid = next_nid
    cx, cy = lx / 2, ly / 2
    model.add_grid(nid=bolt_nid, xyz=[cx, cy, 0.05])
    next_nid += 1

    # Find nearby plate nodes (within radius)
    bolt_radius = 0.15
    bolt_dep_nodes = []
    for (i, j), n in node_grid.items():
        node = model.nodes[n]
        dist = math.sqrt((node.xyz[0] - cx)**2 + (node.xyz[1] - cy)**2)
        if dist <= bolt_radius:
            bolt_dep_nodes.append(n)

    model.add_rbe2(
        eid=next_eid,
        gn=bolt_nid,              # independent node
        cm='123456',              # all DOFs rigid
        Gmi=bolt_dep_nodes,       # dependent nodes
    )
    print(f"RBE2 EID={next_eid}: bolt at node {bolt_nid}, "
          f"{len(bolt_dep_nodes)} dependent nodes")
    next_eid += 1

    # ---------------------------------------------------------------
    # RBE3: Load distribution — distribute force from a point to an
    # edge (does NOT add stiffness)
    # ---------------------------------------------------------------
    # Reference (dependent) node — load application point
    load_ref_nid = next_nid
    model.add_grid(nid=load_ref_nid, xyz=[lx, ly / 2, 0.0])
    next_nid += 1

    # Independent nodes — right edge of plate
    right_edge_nodes = [node_grid[(nx, j)] for j in range(ny + 1)]

    model.add_rbe3(
        eid=next_eid,
        refgrid=load_ref_nid,     # dependent (reference) node
        refc='123456',            # DOFs for dependent node
        weights=[1.0],            # single weight set
        comps=['123'],            # weight applies to T1,T2,T3
        Gijs=[right_edge_nodes],  # independent nodes
    )
    print(f"RBE3 EID={next_eid}: load ref node {load_ref_nid}, "
          f"{len(right_edge_nodes)} independent nodes on right edge")
    next_eid += 1

    # ---------------------------------------------------------------
    # CONM2: Concentrated masses — equipment on the plate
    # ---------------------------------------------------------------
    # Equipment mass at bolt location
    model.add_conm2(
        eid=next_eid,
        nid=bolt_nid,
        mass=5.0,                 # 5 kg
        cid=0,
        X=[0., 0., 0.1],         # offset 100mm above
        I=[0.01, 0., 0.01, 0., 0., 0.01],  # I11, I21, I22, I31, I32, I33
    )
    print(f"CONM2 EID={next_eid}: 5 kg at node {bolt_nid} "
          f"with 100mm Z offset")
    next_eid += 1

    # Additional point masses along top edge
    top_edge_nodes = [node_grid[(i, ny)] for i in range(0, nx + 1, 2)]
    for n in top_edge_nodes:
        model.add_conm2(eid=next_eid, nid=n, mass=0.5)
        next_eid += 1
    print(f"CONM2: {len(top_edge_nodes)} x 0.5 kg masses on top edge")

    # ---------------------------------------------------------------
    # SPC1: Fixed boundary — left edge
    # ---------------------------------------------------------------
    spc_sid = 1
    left_edge = [node_grid[(0, j)] for j in range(ny + 1)]
    model.add_spc1(
        sid=spc_sid,
        components='123456',      # fix all 6 DOFs
        nodes=left_edge,
    )
    print(f"SPC1 SID={spc_sid}: {len(left_edge)} nodes fixed (left edge)")

    # SPC with enforced displacement — push bottom-right corner 1mm in Z
    spc_enf_sid = 2
    corner_nid = node_grid[(nx, 0)]
    model.add_spc(
        sid=spc_enf_sid,
        nodes=[corner_nid],
        components=['3'],          # T3 (Z) direction
        enforced_values=[0.001],   # 1 mm displacement
    )
    print(f"SPC SID={spc_enf_sid}: enforced Z=1mm at node {corner_nid}")

    # ---------------------------------------------------------------
    # MPC: Multi-point constraint — tie two nodes together in T3
    # ---------------------------------------------------------------
    mpc_sid = 1
    node_a = node_grid[(5, 3)]
    node_b = node_grid[(5, 7)]
    # Constraint: 1.0 * T3_a - 1.0 * T3_b = 0  →  T3_a = T3_b
    model.add_mpc(
        sid=mpc_sid,
        nodes=[node_a, node_b],
        components=['3', '3'],
        coefficients=[1.0, -1.0],
    )
    print(f"MPC SID={mpc_sid}: T3 of node {node_a} = T3 of node {node_b}")

    # ---------------------------------------------------------------
    # Loads
    # ---------------------------------------------------------------
    load_sid = 100
    # Force applied at the RBE3 reference node (distributed to edge)
    model.add_force(sid=load_sid, node=load_ref_nid, mag=10000.,
                    xyz=[0., 0., -1.])

    # Gravity
    grav_sid = 200
    model.add_grav(sid=grav_sid, scale=9.81, N=[0., 0., -1.])

    # Combined load
    combined_sid = 999
    model.add_load(sid=combined_sid, scale=1.0,
                   scale_factors=[1.0, 1.0],
                   load_ids=[load_sid, grav_sid])

    # ---------------------------------------------------------------
    # Case control
    # ---------------------------------------------------------------
    cc = model.case_control_deck

    sc1 = cc.create_new_subcase(1)
    sc1.add('SUBTITLE', 'Force + Gravity', options=[], option_type='')
    sc1.add('LOAD', combined_sid, options=[], option_type='')
    sc1.add('SPC', spc_sid, options=[], option_type='')
    sc1.add('MPC', mpc_sid, options=[], option_type='')
    sc1.add('DISPLACEMENT', 'ALL', options=['SORT1', 'REAL'],
            option_type='STRESS-type')
    sc1.add('STRESS', 'ALL',
            options=['SORT1', 'REAL', 'VONMISES', 'BILIN'],
            option_type='STRESS-type')
    sc1.add('SPCFORCES', 'ALL', options=['SORT1', 'REAL'],
            option_type='STRESS-type')
    sc1.add('MPCFORCES', 'ALL', options=['SORT1', 'REAL'],
            option_type='STRESS-type')

    model.add_param('POST', [-1])
    model.add_param('AUTOSPC', ['YES'])

    model.validate()
    model.write_bdf(bdf_out)
    print(f"\nWritten to {bdf_out}")


if __name__ == '__main__':
    if len(sys.argv) != 2:
        print("Usage: python rbe_and_constraints.py <output.bdf>")
        sys.exit(1)
    build_model(sys.argv[1])
