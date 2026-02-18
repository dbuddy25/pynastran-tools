# BDF Card Reference

Complete reference for BDF card types available through pyNastran.
Cards are accessed via typed dictionaries on the `BDF` model object.

---

## Nodes

### GRID
Defines a structural grid point.

```
Access:  model.nodes[nid]
Fields:  nid, cp, xyz, cd, ps, seid
Methods: get_position() — global XYZ
         get_position_wrt(model, cid) — XYZ in coord system cid
```

```python
model.add_grid(nid=1, xyz=[0., 0., 0.], cp=0, cd=0, ps='', seid=0)
```

### SPOINT
Scalar point for extra DOFs.

```
Access:  model.spoints
```

```python
model.add_spoint([1001, 1002, 1003])
```

---

## Shell Elements

### CQUAD4
4-node isoparametric shell element.

```
Access:  model.elements[eid]
Fields:  eid, pid, nids (4), theta_mcid, zoffset, tflag, T1-T4
Methods: Area(), Normal(), Centroid(), Mass()
```

```python
model.add_cquad4(eid=1, pid=1, nids=[1, 2, 3, 4])
```

### CQUAD8
8-node quadrilateral shell (parabolic edges).

```
Access:  model.elements[eid]
Fields:  eid, pid, nids (8 — corners + midside, midside can be 0/None)
```

```python
model.add_cquad8(eid=1, pid=1, nids=[1, 2, 3, 4, 5, 6, 7, 8])
```

### CTRIA3
3-node triangular shell element.

```
Access:  model.elements[eid]
Fields:  eid, pid, nids (3), theta_mcid, zoffset
Methods: Area(), Normal(), Centroid(), Mass()
```

```python
model.add_ctria3(eid=2, pid=1, nids=[1, 2, 3])
```

### CTRIA6
6-node triangular shell (parabolic edges).

```
Access:  model.elements[eid]
Fields:  eid, pid, nids (6)
```

```python
model.add_ctria6(eid=2, pid=1, nids=[1, 2, 3, 4, 5, 6])
```

---

## Solid Elements

### CHEXA
8-node or 20-node hexahedral solid.

```
Access:  model.elements[eid]
Fields:  eid, pid, nids (8 or 20)
Methods: Volume(), Centroid(), Mass()
```

```python
model.add_chexa(eid=100, pid=5, nids=[1,2,3,4,5,6,7,8])
```

### CPENTA
6-node or 15-node pentahedral (wedge) solid.

```
Access:  model.elements[eid]
Fields:  eid, pid, nids (6 or 15)
```

```python
model.add_cpenta(eid=101, pid=5, nids=[1,2,3,4,5,6])
```

### CTETRA
4-node or 10-node tetrahedral solid.

```
Access:  model.elements[eid]
Fields:  eid, pid, nids (4 or 10)
```

```python
model.add_ctetra(eid=102, pid=5, nids=[1,2,3,4])
```

---

## Bar / Beam Elements

### CBAR
Simple bar element (axial, bending, torsion).

```
Access:  model.elements[eid]
Fields:  eid, pid, nids (2), x/g0 (orientation), offt, pa, pb, wa, wb
```

```python
model.add_cbar(eid=200, pid=10, nids=[1, 2], x=[0., 0., 1.])
```

### CBEAM
Beam element with variable cross-section.

```
Access:  model.elements[eid]
Fields:  eid, pid, nids (2), x/g0, offt, wa, wb, sa, sb
```

```python
model.add_cbeam(eid=201, pid=10, nids=[1, 2], x=[0., 0., 1.])
```

### CROD
Rod element (axial + torsion only).

```
Access:  model.elements[eid]
Fields:  eid, pid, nids (2)
```

```python
model.add_crod(eid=202, pid=11, nids=[1, 2])
```

### CONROD
Rod element with inline properties (no separate PROD).

```
Access:  model.elements[eid]
Fields:  eid, nids (2), mid, A, j, c, nsm
```

```python
model.add_conrod(eid=203, nids=[1, 2], mid=1, A=0.001)
```

---

## Mass Elements

### CONM2
Concentrated mass at a grid point (6 DOF).

```
Access:  model.masses[eid]
Fields:  eid, nid, cid, mass, X (offset), I (inertia: I11,I21,I22,I31,I32,I33)
```

```python
model.add_conm2(eid=300, nid=1, mass=10.0, cid=0,
                X=[0., 0., 0.], I=[0., 0., 0., 0., 0., 0.])
```

### CONM1
Concentrated mass defined via a 6x6 mass matrix.

```
Access:  model.masses[eid]
Fields:  eid, nid, cid, mass_matrix (lower triangular)
```

### CMASS2
Scalar mass (single DOF).

```
Access:  model.masses[eid]
Fields:  eid, mass, G1, C1, G2, C2
```

---

## Spring / Damper / Bush Elements

### CELAS2
Scalar spring with inline stiffness.

```
Access:  model.elements[eid]
Fields:  eid, k, G1, C1, G2, C2
```

```python
model.add_celas2(eid=400, k=1.0e6, nids=[1, 2], c1=1, c2=1)
```

### CBUSH
Generalized bushing element (6 DOF spring/damper).

```
Access:  model.elements[eid]
Fields:  eid, pid, nids (1 or 2), x/g0, cid
```

```python
model.add_cbush(eid=401, pid=20, nids=[1, 2], x=[1., 0., 0.])
```

### CDAMP2
Scalar damper with inline damping.

```
Access:  model.elements[eid]
Fields:  eid, b, nids, c1, c2
```

---

## Rigid Elements

### RBE2
Rigid body element — independent node drives dependent nodes.

```
Access:  model.rigid_elements[eid]
Fields:  eid, gn (independent), cm (dof string), Gmi (dependent nodes list)
```

```python
model.add_rbe2(eid=500, gn=1, cm='123456', Gmi=[2, 3, 4, 5])
```

### RBE3
Interpolation element — dependent node is weighted average of
independent nodes. Does NOT add stiffness.

```
Access:  model.rigid_elements[eid]
Fields:  eid, refgrid, refc, weights, comps, Gijs
```

```python
model.add_rbe3(eid=501, refgrid=1, refc='123456',
               weights=[1.0], comps=['123'],
               Gijs=[[2, 3, 4, 5]])
```

### RBAR
Rigid bar between two nodes.

```
Access:  model.rigid_elements[eid]
Fields:  eid, nids (2), CNA, CNB, CMA, CMB
```

---

## Shell Properties

### PSHELL
Isotropic/anisotropic shell property.

```
Access:  model.properties[pid]
Fields:  pid, mid1, t, mid2, bk (12I/t^3), mid3, tst, nsm, z1, z2
```

```python
model.add_pshell(pid=1, mid1=1, t=0.005)
```

### PCOMP
Composite layup (symmetric/asymmetric).

```
Access:  model.properties[pid]
Fields:  pid, lam, z0, nsm, sb, ft, tref, ge
         Per ply: thicknesses[], mids[], thetas[], souts[]
Methods: nplies, total_thickness, is_symmetrical
```

```python
model.add_pcomp(pid=2, mids=[1, 1, 1, 1],
                thicknesses=[0.125e-3]*4,
                thetas=[0., 45., -45., 90.])
```

### PCOMPG
Composite layup with global ply IDs.

```
Access:  model.properties[pid]
Fields:  pid, plus PCOMP fields with global_ply_ids[]
```

---

## Solid Properties

### PSOLID
Isotropic solid property.

```
Access:  model.properties[pid]
Fields:  pid, mid, cordm, integ, stress, isop, fctn
```

```python
model.add_psolid(pid=5, mid=1)
```

---

## Bar / Beam Properties

### PBAR
Simple bar property.

```
Access:  model.properties[pid]
Fields:  pid, mid, A, i1, i2, j, nsm, c1-d2, e1-f2, k1, k2
```

```python
model.add_pbar(pid=10, mid=1, A=0.001, i1=1e-8, i2=1e-8, j=2e-8)
```

### PBARL
Bar property from library cross-section (ROD, BAR, BOX, I, T, etc.).

```
Access:  model.properties[pid]
Fields:  pid, mid, Type, dim, nsm, group
```

```python
model.add_pbarl(pid=11, mid=1, Type='ROD', dim=[0.01])  # radius
```

### PBEAM
Beam property with full cross-section data.

```
Access:  model.properties[pid]
Fields:  pid, mid, A, i1, i2, j, (stations A-K), nsm
```

### PBEAML
Beam property from library cross-section.

```
Access:  model.properties[pid]
Fields:  pid, mid, Type, dim, nsm, group
```

### PROD
Rod property.

```
Access:  model.properties[pid]
Fields:  pid, mid, A, j, c, nsm
```

```python
model.add_prod(pid=12, mid=1, A=0.001, j=1e-8)
```

### PBUSH
Bush property (6 DOF stiffness/damping).

```
Access:  model.properties[pid]
Fields:  pid, Ki (stiffness), Bi (damping), GEi
```

```python
model.add_pbush(pid=20, Ki=[1e6, 1e6, 1e6, 1e4, 1e4, 1e4])
```

---

## Isotropic Materials

### MAT1
Isotropic material.

```
Access:  model.materials[mid]
Fields:  mid, E, G, nu, rho, a, tref, ge, St, Sc, Ss
```

```python
model.add_mat1(mid=1, E=2.1e11, G=None, nu=0.3, rho=7850.,
               a=1.2e-5, tref=20.)
```

### MAT2
Anisotropic shell material (2D).

```
Access:  model.materials[mid]
Fields:  mid, G11-G33, rho, a1-a3, tref, ge
```

---

## Orthotropic / Composite Materials

### MAT8
Orthotropic material for shells/composites.

```
Access:  model.materials[mid]
Fields:  mid, E1, E2, nu12, G12, G1z, G2z, rho, a1, a2, tref, ge,
         Xt, Xc, Yt, Yc, S, strn
```

```python
model.add_mat8(mid=2, E1=140e9, E2=10e9, nu12=0.3,
               G12=5e9, G1z=5e9, G2z=3e9, rho=1600.)
```

### MAT9
Anisotropic solid material (3D).

```
Access:  model.materials[mid]
```

---

## Static Loads

### FORCE
Concentrated force at a grid point.

```
Access:  model.loads[sid]  →  list of load cards
Fields:  sid, node, mag, xyz (direction), cid
```

```python
model.add_force(sid=100, node=1, mag=1000., xyz=[0., 0., -1.], cid=0)
```

### MOMENT
Concentrated moment at a grid point.

```
Access:  model.loads[sid]  →  list
Fields:  sid, node, mag, xyz, cid
```

```python
model.add_moment(sid=100, node=1, mag=500., xyz=[1., 0., 0.])
```

### PLOAD4
Pressure load on shell/solid element faces.

```
Access:  model.loads[sid]  →  list
Fields:  sid, eids, pressures (4), g1, g34, cid, nvector, surf_or_line, line_load_dir
```

```python
model.add_pload4(sid=100, eids=[1, 2, 3], pressures=[1e5, 1e5, 1e5, 1e5])
```

### GRAV
Gravity load.

```
Access:  model.loads[sid]  →  list
Fields:  sid, scale, N (direction), cid, mb
```

```python
model.add_grav(sid=100, scale=9.81, N=[0., 0., -1.])
```

### LOAD
Load combination card — scales and combines other load sets.

```
Access:  model.load_combinations[sid]  →  list
Fields:  sid, scale, scale_factors, load_ids
```

```python
model.add_load(sid=999, scale=1.0, scale_factors=[1.0, 0.5],
               load_ids=[100, 200])
```

### TEMP
Nodal temperature.

```
Access:  model.loads[sid]  →  list
Fields:  sid, temperatures (dict of nid → temp)
```

```python
model.add_temp(sid=100, nodes=[1, 2, 3],
               temperatures=[100., 150., 200.])
```

### TEMPD
Default temperature for all unlisted nodes.

```
Access:  model.tempds[sid]
Fields:  sid, temperature
```

```python
model.add_tempd(sid=100, temperature=20.)
```

---

## Dynamic Loads

### RLOAD1 / RLOAD2
Frequency-dependent dynamic loads.

```
Access:  model.loads[sid]  →  list
Fields:  sid, excite_id, delay, dphase, tc, td, Type
```

### TLOAD1 / TLOAD2
Time-dependent dynamic loads.

```
Access:  model.loads[sid]  →  list
Fields:  sid, excite_id, delay, Type, tid (TABLED1 reference)
```

### DAREA
Dynamic load application point/DOF.

```
Access:  model.dareas[sid]
Fields:  sid, nodes, components, scales
```

### TABLED1
Tabular function (x, y pairs) for dynamic analysis.

```
Access:  model.tables[tid]
Fields:  tid, x, y, xaxis, yaxis
```

```python
model.add_tabled1(tid=1, x=[0., 1., 2.], y=[0., 100., 0.])
```

### EIGRL
Eigenvalue extraction (Lanczos).

```
Access:  model.methods[sid]
Fields:  sid, v1, v2, nd, msglvl, maxset, shfscl
```

```python
model.add_eigrl(sid=10, v1=0., v2=1000., nd=20)
```

---

## Constraints

### SPC
Single-point constraint (prescribe DOF value).

```
Access:  model.spcs[sid]  →  list
Fields:  sid, nodes, components (DOF string), enforced_values
```

```python
model.add_spc(sid=1, nodes=[1], components=['123456'],
              enforced_values=[0.])
```

### SPC1
Single-point constraint (zero enforced, multiple nodes, one DOF set).

```
Access:  model.spcs[sid]  →  list
Fields:  sid, components (DOF string), nodes
```

```python
model.add_spc1(sid=1, components='123456', nodes=[1, 2, 3, 4])
```

### MPC
Multi-point constraint (linear equation between DOFs).

```
Access:  model.mpcs[sid]  →  list
Fields:  sid, nodes, components, coefficients
```

```python
model.add_mpc(sid=1, nodes=[10, 20], components=['1', '1'],
              coefficients=[1.0, -1.0])
```

### SUPORT / SUPORT1
Reference point for inertia relief.

```
Access:  model.suport / model.suport1
```

---

## Coordinate Systems

### CORD2R
Rectangular coordinate system defined by 3 points.

```
Access:  model.coords[cid]
Fields:  cid, origin, zaxis, xzplane, rid (reference coord)
Methods: transform_node_to_global(xyz)
         transform_node_to_local(xyz)
```

```python
model.add_cord2r(cid=1, origin=[0.,0.,0.], zaxis=[0.,0.,1.],
                 xzplane=[1.,0.,0.])
```

### CORD2C
Cylindrical coordinate system.

```
Access:  model.coords[cid]
```

### CORD2S
Spherical coordinate system.

```
Access:  model.coords[cid]
```

---

## Sets

### SET1
A set of grid or element IDs.

```
Access:  model.sets[sid]
Fields:  sid, ids (list of ints)
```

```python
model.add_set1(sid=1, ids=[1, 2, 3, 4, 5])
```

### SET3
Similar to SET1, used in design optimization.

```
Access:  model.sets[sid]
```
