# OP2 Results Reference

Complete reference for pyNastran OP2 result objects, data shapes, and
column ordering.

## Reading an OP2 File

```python
from pyNastran.op2.op2 import OP2

op2 = OP2()
op2.read_op2('model.op2')

# To exclude specific result types (faster loading):
op2.set_results(['displacements', 'cquad4_stress'])
op2.read_op2('model.op2')
```

## Result Object Anatomy

All result objects share a common structure:

```python
result = op2.displacements[subcase_id]

# Core arrays
result.data                  # numpy array: (ntimes, n, ncomponents)
result.data_frame            # pandas DataFrame view

# ID mapping
result.node_gridtype         # (n, 2) for nodal results — [node_id, grid_type]
result.element_node          # (n, 2) for element results — [element_id, node_id]
result.element               # (n,) for element-only results — [element_id]

# Time / mode / frequency axis
result._times                # 1D array of time/mode/freq values
result.modes                 # mode numbers (modal analysis)
```

### Grid Types
- 1 = GRID (structural)
- 2 = SPOINT (scalar point)
- 7 = EPOINT (extra point)

---

## Nodal Results

### Displacements — `op2.displacements[subcase]`

```
Shape: (ntimes, nnodes, 6)
Columns: T1, T2, T3, R1, R2, R3
ID map: node_gridtype[:, 0] → node IDs
```

```python
disp = op2.displacements[1]
node_ids = disp.node_gridtype[:, 0]

# All T3 (vertical) displacements at first time step
tz = disp.data[0, :, 2]

# Find specific node
idx = np.where(node_ids == 100)[0][0]
tz_100 = disp.data[0, idx, 2]

# DataFrame access
df = disp.data_frame
```

### Velocities — `op2.velocities[subcase]`

```
Shape: (ntimes, nnodes, 6)
Columns: T1, T2, T3, R1, R2, R3
```

### Accelerations — `op2.accelerations[subcase]`

```
Shape: (ntimes, nnodes, 6)
Columns: T1, T2, T3, R1, R2, R3
```

### SPC Forces — `op2.spc_forces[subcase]`

```
Shape: (ntimes, nnodes, 6)
Columns: T1, T2, T3, R1, R2, R3
```

### Applied Loads — `op2.load_vectors[subcase]`

```
Shape: (ntimes, nnodes, 6)
Columns: T1, T2, T3, R1, R2, R3
```

---

## Plate Stress / Strain

### CQUAD4 Stress — `op2.cquad4_stress[subcase]`

```
Shape: (ntimes, n, 8)
Columns: fiber_distance, oxx, oyy, txy, angle, omax, omin, von_mises
ID map: element_node[:, 0] → element IDs
         element_node[:, 1] → node IDs (0 = centroid)
```

**BILINEAR output (default):** For each CQUAD4, there are 5 rows:
1 centroid (node_id=0) + 4 corner nodes. Two layers (Z1 top, Z2 bottom)
are interleaved, giving 10 rows total per element.

**CENTROID output:** 2 rows per element (Z1 top, Z2 bottom).

```python
stress = op2.cquad4_stress[1]

# Centroid-only mask
centroid_mask = stress.element_node[:, 1] == 0

# Max von Mises across all elements, centroid, all times
vm = stress.data[:, centroid_mask, 7]  # column 7 = von_mises
max_vm = vm.max()

# DataFrame
df = stress.data_frame
```

### CTRIA3 Stress — `op2.ctria3_stress[subcase]`

```
Shape: (ntimes, n, 8)
Columns: fiber_distance, oxx, oyy, txy, angle, omax, omin, von_mises
ID map: element_node or element
```

CTRIA3 always outputs centroid results: 2 rows per element (Z1, Z2).

### CQUAD4 Strain — `op2.cquad4_strain[subcase]`

Same shape and columns as stress but with strain values.

### Composite Plate Stress — `op2.cquad4_composite_stress[subcase]`

Per-ply stress for PCOMP elements.

```
Shape: (ntimes, n, 9)
Columns: o11, o22, t12, t1z, t2z, angle, omax, omin, von_mises
```

Each ply produces a separate row.

---

## Solid Stress / Strain

### CHEXA Stress — `op2.chexa_stress[subcase]`

```
Shape: (ntimes, n, 10)
Columns: oxx, oyy, ozz, txy, tyz, txz, omax, omid, omin, von_mises
ID map: element_node[:, 0] → element IDs
         element_node[:, 1] → node IDs (0 = centroid)
```

For each CHEXA, there is 1 centroid row + 8 corner node rows (9 total).

```python
stress = op2.chexa_stress[1]
centroid_mask = stress.element_node[:, 1] == 0
vm = stress.data[:, centroid_mask, 9]  # von_mises
```

### CPENTA Stress — `op2.cpenta_stress[subcase]`

```
Shape: (ntimes, n, 10)
Columns: oxx, oyy, ozz, txy, tyz, txz, omax, omid, omin, von_mises
```

1 centroid + 6 corner node rows per element.

### CTETRA Stress — `op2.ctetra_stress[subcase]`

```
Shape: (ntimes, n, 10)
Columns: oxx, oyy, ozz, txy, tyz, txz, omax, omid, omin, von_mises
```

1 centroid + 4 corner node rows per element.

---

## Bar / Beam Forces and Stress

### CBAR Force — `op2.cbar_force[subcase]`

```
Shape: (ntimes, n, 8)
Columns: bending_moment_a1, bending_moment_a2, bending_moment_b1,
         bending_moment_b2, shear1, shear2, axial, torque
ID map: element → element IDs
```

### CBAR Stress — `op2.cbar_stress[subcase]`

```
Shape: (ntimes, n, 15)
Columns: s1a, s2a, s3a, s4a, axial, smaxa, smina, margin_tension_a,
         s1b, s2b, s3b, s4b, smaxb, sminb, margin_tension_b
```

### CBEAM Force — `op2.cbeam_force[subcase]`

```
Shape: (ntimes, n, 11)
Per station (A and B).
```

---

## Grid Point Forces — `op2.grid_point_forces[subcase]`

Force balance at each node (applied loads, SPC forces, MPC forces,
element forces).

```
Shape: (ntimes, n, 6)
Columns: T1, T2, T3, R1, R2, R3
```

Requires `GPFORCE = ALL` in case control.

---

## Eigenvalue Tables — `op2.eigenvalues`

**Keys are strings** (the eigenvalue table title), not integers.

```python
for title, eigval_table in op2.eigenvalues.items():
    modes = eigval_table.mode            # array of mode numbers
    extraction_order = eigval_table.extraction_order
    eigenvalues = eigval_table.eigenvalue # raw eigenvalues (rad/s)^2
    radians = eigval_table.radians       # circular freq (rad/s)
    freqs = eigval_table.cycles          # frequency (Hz)
    gen_mass = eigval_table.generalized_mass
    gen_stiffness = eigval_table.generalized_stiffness
```

Mode shapes are in `op2.eigenvectors[subcase]`:

```python
eigvec = op2.eigenvectors[1]
# data shape: (nmodes, nnodes, 6)
mode1_tz = eigvec.data[0, :, 2]  # mode 1, all nodes, T3
```

---

## Grid Point Weight — `op2.grid_point_weight`

Mass summary from Nastran's GPWG output.

```python
gpw = next(iter(op2.grid_point_weight.values()))
mass = gpw.mass          # (6,6) mass matrix
cg = gpw.cg              # (3,3) CG location per axis
IS = gpw.IS              # (3,3) inertia at CG (principal axes)
IQ = gpw.IQ              # (3,) principal inertias
Q = gpw.Q                # (3,3) principal axis directions
```

The scalar total mass is `gpw.mass[0, 0]`.

---

## MEFFMASS Matrices

When `MEFFMASS(PLOT)` or `MEFFMASS(PLOT,FRACSUM)` is in case control,
Nastran writes exact modal effective mass data as named matrices into the
OP2. pyNastran exposes these as convenience properties:

| Property | Matrix Key | Shape | Content |
|---|---|---|---|
| `op2.modal_effective_mass_fraction` | `EFMFACS` | (6, nmodes) | Mass fractions (0–1) |
| `op2.modal_effective_mass` | `MEFMASS` | (6, nmodes) | Effective mass |
| `op2.modal_participation_factors` | `MPFACS` | (6, nmodes) | Participation factors |
| `op2.modal_effective_weight` | `MEFWTS` | (6, nmodes) | Effective weight |
| `op2.total_effective_mass_matrix` | `EFMFSMS` | (6, 1) | Total effective mass (cumulative sum) |
| `op2.effective_mass_matrix` | `EFMASSS` | (6, 6) | Effective mass matrix |
| `op2.rigid_body_mass_matrix` | `RBMASS` | (6, 6) | Rigid body mass matrix |

Rows correspond to T1, T2, T3, R1, R2, R3. Each property returns a
`Matrix` object whose `.data` may be a `numpy.ndarray` or
`scipy.sparse` matrix.

```python
import scipy.sparse

# Read mass fractions (6 x nmodes)
meff_frac = op2.modal_effective_mass_fraction
if meff_frac is not None:
    data = meff_frac.data
    if scipy.sparse.issparse(data):
        data = data.toarray()
    # Transpose to (nmodes, 6) for per-mode rows
    per_mode = data.T
    print(per_mode.shape)  # (nmodes, 6)

# Read rigid body mass matrix (6x6)
rbmass = op2.rigid_body_mass_matrix
if rbmass is not None:
    rb = rbmass.data
    if scipy.sparse.issparse(rb):
        rb = rb.toarray()
    total_mass = rb[0, 0]
```

---

## Coordinate Transforms

Result data is usually in the **output coordinate system** (CD field
of the GRID card). To get results in the global coordinate system when
CD ≠ 0:

```python
disp = op2.displacements[1]
# If PARAM,POST,-1 is set, data is already in global
# Otherwise check disp.data_transformation_msg
```

For manual transformation:

```python
# After cross-referencing the BDF:
node = model.nodes[nid]
T = node.cd_ref.beta()  # 3x3 rotation from local to global
disp_global = T @ disp_local
```

---

## DataFrame Access

All result objects support `data_frame` for pandas access:

```python
df = op2.displacements[1].data_frame
# MultiIndex columns, rows indexed by node ID
# Easy filtering, grouping, export to CSV

df.to_csv('displacements.csv')
```

For stress results:

```python
df = op2.cquad4_stress[1].data_frame
# Columns: fiber_distance, oxx, oyy, txy, angle, omax, omin, von_mises
# Index: (ElementID, NodeID) — NodeID=0 for centroid
```
