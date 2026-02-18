# Model Building API Reference

Full `add_*` method signatures for programmatically building a BDF model.

## Creating an Empty Model

```python
from pyNastran.bdf.bdf import BDF

model = BDF()
model.sol = 101          # solution type
```

---

## Nodes

### add_grid
```python
model.add_grid(nid, xyz, cp=0, cd=0, ps='', seid=0, comment='')
```
- **nid**: int — grid point ID
- **xyz**: list[float] — [x, y, z] in coordinate system cp
- **cp**: int — coordinate system for input
- **cd**: int — coordinate system for output
- **ps**: str — permanent SPC DOFs (e.g., '456' for shells)
- **seid**: int — superelement ID

### add_spoint
```python
model.add_spoint(ids, comment='')
```
- **ids**: list[int] — scalar point IDs

---

## Shell Elements

### add_cquad4
```python
model.add_cquad4(eid, pid, nids, theta_mcid=0.0, zoffset=0.,
                 tflag=0, T1=None, T2=None, T3=None, T4=None, comment='')
```
- **nids**: list[int] — 4 node IDs (counterclockwise)
- **theta_mcid**: float or int — material angle or coord system ID
- **zoffset**: float — offset from shell surface

### add_cquad8
```python
model.add_cquad8(eid, pid, nids, theta_mcid=0.0, zoffset=0.,
                 tflag=0, T1=None, T2=None, T3=None, T4=None, comment='')
```
- **nids**: list[int] — 8 node IDs (4 corners + 4 midside; midside=0 if absent)

### add_ctria3
```python
model.add_ctria3(eid, pid, nids, zoffset=0., theta_mcid=0.0,
                 tflag=0, T1=None, T2=None, T3=None, comment='')
```
- **nids**: list[int] — 3 node IDs

### add_ctria6
```python
model.add_ctria6(eid, pid, nids, theta_mcid=0.0, zoffset=0.,
                 tflag=0, T1=None, T2=None, T3=None, comment='')
```
- **nids**: list[int] — 6 node IDs (3 corners + 3 midside)

---

## Solid Elements

### add_chexa
```python
model.add_chexa(eid, pid, nids, comment='')
```
- **nids**: list[int] — 8 or 20 node IDs

### add_cpenta
```python
model.add_cpenta(eid, pid, nids, comment='')
```
- **nids**: list[int] — 6 or 15 node IDs

### add_ctetra
```python
model.add_ctetra(eid, pid, nids, comment='')
```
- **nids**: list[int] — 4 or 10 node IDs

---

## Bar / Beam Elements

### add_cbar
```python
model.add_cbar(eid, pid, nids, x, g0=None, offt='GGG',
               pa=0, pb=0, wa=None, wb=None, comment='')
```
- **nids**: list[int] — [node_a, node_b]
- **x**: list[float] — orientation vector [x1, x2, x3]
- **g0**: int — orientation node (alternative to x)
- **offt**: str — offset type flag

### add_cbeam
```python
model.add_cbeam(eid, pid, nids, x, g0=None, offt='GGG',
                bit=None, pa=0, pb=0, wa=None, wb=None,
                sa=0, sb=0, comment='')
```

### add_crod
```python
model.add_crod(eid, pid, nids, comment='')
```

### add_conrod
```python
model.add_conrod(eid, nids, mid, A=0., j=0., c=0., nsm=0., comment='')
```

---

## Mass Elements

### add_conm2
```python
model.add_conm2(eid, nid, mass, cid=0, X=None, I=None, comment='')
```
- **X**: list[float] — offset [x1, x2, x3] from grid point
- **I**: list[float] — [I11, I21, I22, I31, I32, I33] inertia

### add_cmass2
```python
model.add_cmass2(eid, mass, nids, c1=0, c2=0, comment='')
```

---

## Spring / Damper / Bush Elements

### add_celas2
```python
model.add_celas2(eid, k, nids, c1=0, c2=0, ge=0., s=0., comment='')
```

### add_cbush
```python
model.add_cbush(eid, pid, nids, x=None, g0=None, cid=None, comment='')
```

### add_cdamp2
```python
model.add_cdamp2(eid, b, nids, c1=0, c2=0, comment='')
```

---

## Rigid Elements

### add_rbe2
```python
model.add_rbe2(eid, gn, cm, Gmi, alpha=0., comment='')
```
- **gn**: int — independent grid point
- **cm**: str — DOF components (e.g., '123456')
- **Gmi**: list[int] — dependent grid points

### add_rbe3
```python
model.add_rbe3(eid, refgrid, refc, weights, comps, Gijs,
               Gmi=None, Cmi=None, alpha=0., comment='')
```
- **refgrid**: int — dependent (reference) grid point
- **refc**: str — dependent DOF components
- **weights**: list[float] — weighting factors for each set
- **comps**: list[str] — DOF component strings for each set
- **Gijs**: list[list[int]] — independent grid points for each set

---

## Shell Properties

### add_pshell
```python
model.add_pshell(pid, mid1=None, t=None, mid2=None, twelveIt3=1.0,
                 mid3=None, tst=0.833333, nsm=0.,
                 z1=None, z2=None, mid4=None, comment='')
```
- **mid1**: int — membrane material
- **t**: float — thickness
- **mid2**: int — bending material (default same as mid1)
- **twelveIt3**: float — bending stiffness ratio (12I/t^3)

### add_pcomp
```python
model.add_pcomp(pid, mids, thicknesses, thetas=None, souts=None,
                nsm=0., sb=0., ft=None, tref=0., ge=0.,
                lam=None, z0=None, comment='')
```
- **mids**: list[int] — material IDs per ply
- **thicknesses**: list[float] — ply thicknesses
- **thetas**: list[float] — ply angles in degrees (default all 0)
- **souts**: list[str] — stress output per ply ('YES'/'NO')
- **lam**: str — 'SYM' for symmetric layup, None for full definition

---

## Solid Properties

### add_psolid
```python
model.add_psolid(pid, mid, cordm=0, integ=None, stress=None,
                 isop=None, fctn='SMECH', comment='')
```

---

## Bar / Beam Properties

### add_pbar
```python
model.add_pbar(pid, mid, A=0., i1=0., i2=0., i12=0., j=0., nsm=0.,
               c1=0., c2=0., d1=0., d2=0., e1=0., e2=0., f1=0., f2=0.,
               k1=1.e8, k2=1.e8, comment='')
```

### add_pbarl
```python
model.add_pbarl(pid, mid, Type, dim, group='MSCBML0', nsm=0., comment='')
```
- **Type**: str — cross-section type: 'ROD', 'BAR', 'BOX', 'I', 'T', 'TUBE', etc.
- **dim**: list[float] — dimensions (varies by Type)

### add_pbeam
```python
model.add_pbeam(pid, mid, xxb, so, area, i1, i2, i12, j, nsm=None,
                c1=None, c2=None, d1=None, d2=None,
                e1=None, e2=None, f1=None, f2=None,
                k1=1., k2=1., s1=0., s2=0., nsia=0., nsib=None,
                cwa=0., cwb=None, m1a=0., m2a=0., m1b=None, m2b=None,
                n1a=0., n2a=0., n1b=None, n2b=None, comment='')
```

### add_pbeaml
```python
model.add_pbeaml(pid, mid, Type, xxb, dims, so=None, nsm=None,
                 group='MSCBML0', comment='')
```

### add_prod
```python
model.add_prod(pid, mid, A, j=0., c=0., nsm=0., comment='')
```

### add_pbush
```python
model.add_pbush(pid, Ki=None, Bi=None, GEi=None, rcv=None, comment='')
```
- **Ki**: list[float] — stiffness [K1..K6]
- **Bi**: list[float] — damping [B1..B6]

---

## Isotropic Materials

### add_mat1
```python
model.add_mat1(mid, E=None, G=None, nu=None, rho=0.,
               a=0., tref=0., ge=0., St=0., Sc=0., Ss=0.,
               mcsid=0, comment='')
```
- Provide 2 of E, G, nu — the third is computed.

---

## Orthotropic / Composite Materials

### add_mat8
```python
model.add_mat8(mid, E1=0., E2=0., nu12=0.,
               G12=0., G1z=1e8, G2z=1e8, rho=0.,
               a1=0., a2=0., tref=0., ge=0.,
               Xt=0., Xc=None, Yt=0., Yc=None, S=0.,
               strn=0., comment='')
```

---

## Static Loads

### add_force
```python
model.add_force(sid, node, mag, xyz, cid=0, comment='')
```

### add_moment
```python
model.add_moment(sid, node, mag, xyz, cid=0, comment='')
```

### add_pload4
```python
model.add_pload4(sid, eids, pressures, g1=None, g34=None,
                 cid=0, nvector=None, surf_or_line='SURF',
                 line_load_dir='NORM', comment='')
```
- **pressures**: list[float] — 4 pressure values (uniform if all equal)
- **eids**: list[int] — element IDs

### add_grav
```python
model.add_grav(sid, scale, N, cid=0, mb=0, comment='')
```
- **scale**: float — acceleration magnitude
- **N**: list[float] — direction vector [N1, N2, N3]

### add_load
```python
model.add_load(sid, scale, scale_factors, load_ids, comment='')
```
- **scale**: float — overall scale factor
- **scale_factors**: list[float] — per-load-set scale factors
- **load_ids**: list[int] — referenced load set IDs

### add_temp
```python
model.add_temp(sid, nodes, temperatures, comment='')
```

### add_tempd
```python
model.add_tempd(sid, temperature, comment='')
```

---

## Dynamic Loads

### add_eigrl
```python
model.add_eigrl(sid, v1=None, v2=None, nd=None, msglvl=0,
                maxset=None, shfscl=None, norm=None,
                options=None, values=None, comment='')
```
- **v1, v2**: float — frequency range (Hz)
- **nd**: int — number of desired roots

### add_tabled1
```python
model.add_tabled1(tid, x, y, xaxis='LINEAR', yaxis='LINEAR', comment='')
```

### add_darea
```python
model.add_darea(sid, nid, component, scale, comment='')
```

---

## Constraints

### add_spc
```python
model.add_spc(sid, nodes, components, enforced_values, comment='')
```

### add_spc1
```python
model.add_spc1(sid, components, nodes, comment='')
```
- **components**: str — DOF string (e.g., '123456')
- **nodes**: list[int] — constrained nodes

### add_mpc
```python
model.add_mpc(sid, nodes, components, coefficients, comment='')
```

---

## Coordinate Systems

### add_cord2r
```python
model.add_cord2r(cid, origin, zaxis, xzplane, rid=0, comment='')
```
- **origin**: list[float] — [x, y, z] of origin
- **zaxis**: list[float] — point on z-axis
- **xzplane**: list[float] — point in x-z plane

### add_cord2c
```python
model.add_cord2c(cid, origin, zaxis, xzplane, rid=0, comment='')
```

### add_cord2s
```python
model.add_cord2s(cid, origin, zaxis, xzplane, rid=0, comment='')
```

---

## Sets

### add_set1
```python
model.add_set1(sid, ids, is_skin=False, comment='')
```

---

## Case Control Deck Setup

```python
# Set solution type
model.sol = 101  # static

# Create case control
cc = model.case_control_deck

# Create a subcase
subcase = cc.create_new_subcase(1)

# Add entries to subcase
subcase.add('SUBTITLE', 'My Load Case', options=[], option_type='')
subcase.add('LOAD', 100, options=[], option_type='')
subcase.add('SPC', 1, options=[], option_type='')
subcase.add('DISPLACEMENT', 'ALL', options=['SORT1', 'REAL'],
            option_type='STRESS-type')
subcase.add('STRESS', 'ALL', options=['SORT1', 'REAL', 'VONMISES', 'BILIN'],
            option_type='STRESS-type')
subcase.add('SPCFORCES', 'ALL', options=['SORT1', 'REAL'],
            option_type='STRESS-type')

# Global entries (subcase 0)
cc.subcases[0].add_integer_type('ECHO', 'NONE')
```

### Params

```python
model.add_param('POST', [-1])     # OP2 output
model.add_param('PRTMAXIM', ['YES'])  # print maximums
model.add_param('AUTOSPC', ['YES'])
```

---

## Validation Workflow

```python
model = BDF()
# ... add cards ...

# Validate before writing
model.validate()

# Cross-reference to check connectivity
model.cross_reference()

# Un-cross-reference before writing
model.uncross_reference()
model.write_bdf('output.bdf')
```
