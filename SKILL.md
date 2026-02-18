---
name: pyNastran FEA Library
description: >
  Deep reference for pyNastran — the Python library for reading, writing,
  and manipulating Nastran BDF (bulk data) and OP2 (binary results) files.
  Use when writing pyNastran code, building FEA models programmatically,
  extracting structural analysis results, computing mass properties, or
  working with Nastran input/output files. Covers BDF cards, OP2 result
  objects, mesh utilities, and common workflows.
  Trigger phrases: pyNastran, Nastran, BDF, OP2, bulk data, finite element,
  FEA model, structural analysis results, mass properties, CQUAD4, CTRIA3,
  eigenvalue, von Mises stress, plate stress, displacement extraction.
version: "1.0"
metadata:
  library: pyNastran
  library_version: ">=1.3"
  python: ">=3.9"
  install: "pip install pyNastran"
---

# pyNastran — Python FEA Toolkit

pyNastran reads, writes, and manipulates MSC/NX/Optistruct Nastran BDF
(bulk data file) models and OP2 (output2 binary) result files entirely
in Python. It enables programmatic model building, automated
post-processing, mass-property extraction, and mesh manipulation without
a Nastran license.

## Scope

**In scope:** SOL 101 (static), SOL 103 (modal), SOL 105 (buckling),
SOL 106 (nonlinear static), SOL 108 (direct frequency response),
SOL 109 (direct transient), SOL 110 (modal complex eigenvalue),
SOL 111 (modal frequency response), SOL 112 (modal transient),
SOL 144 (static aero), SOL 145 (flutter), SOL 146 (dynamic aero),
SOL 200 (optimization). Temperature *loads* (TEMP, TEMPD cards) are
in-scope.

**Out of scope:** Heat transfer analyses (SOL 153 / SOL 159).

## Quick Start

### Reading a BDF

```python
from pyNastran.bdf.bdf import BDF

model = BDF()
model.read_bdf('model.bdf')       # parse
model.cross_reference()            # link cards (enables .pid_ref, .nid_ref, etc.)
```

### Writing a BDF

```python
model.uncross_reference()          # MUST un-cross-ref before writing
model.write_bdf('model_out.bdf')
```

### Reading an OP2

```python
from pyNastran.op2.op2 import OP2

op2 = OP2()
op2.read_op2('model.op2')
```

## Model Data Access

After `read_bdf`, the model stores cards in typed dictionaries keyed by ID:

| Attribute               | Type                | Description                       |
|-------------------------|---------------------|-----------------------------------|
| `model.nodes`           | `dict[int, GRID]`   | Grid points                       |
| `model.elements`        | `dict[int, elem]`   | All elements (shells, solids, …)  |
| `model.properties`      | `dict[int, prop]`   | Property cards                    |
| `model.materials`       | `dict[int, mat]`    | Material cards                    |
| `model.coords`          | `dict[int, coord]`  | Coordinate systems                |
| `model.loads`           | `dict[int, list]`   | Load sets (value is a **list**)   |
| `model.spcs`            | `dict[int, list]`   | SPC constraint sets               |
| `model.mpcs`            | `dict[int, list]`   | MPC constraint sets               |
| `model.rigid_elements`  | `dict[int, elem]`   | RBE2, RBE3, RBAR, etc.           |
| `model.masses`          | `dict[int, mass]`   | CONM1, CONM2, CMASS              |
| `model.sets`            | `dict[int, SET1]`   | SET1/SET3 definitions             |

### Accessing a card

```python
node = model.nodes[100]          # GRID 100
elem = model.elements[5000]      # CQUAD4 / CTRIA3 / CHEXA / …
prop = model.properties[10]      # PSHELL / PSOLID / …
mat  = model.materials[1]        # MAT1 / MAT8 / …
```

## Programmatic Model Building

Create an empty model and add cards with `add_*` methods:

```python
model = BDF()
model.add_grid(nid=1, xyz=[0., 0., 0.])
model.add_grid(nid=2, xyz=[1., 0., 0.])
model.add_grid(nid=3, xyz=[1., 1., 0.])
model.add_grid(nid=4, xyz=[0., 1., 0.])

model.add_mat1(mid=1, E=2.1e11, G=None, nu=0.3, rho=7850.)
model.add_pshell(pid=1, mid1=1, t=0.005)
model.add_cquad4(eid=1, pid=1, nids=[1, 2, 3, 4])
```

See `references/model-building-api.md` for full `add_*` method signatures.

## OP2 Results

Result objects live in typed dictionaries on the OP2 object. Each result
is stored per subcase.

| Attribute                    | Result Type              |
|------------------------------|--------------------------|
| `op2.displacements`         | Nodal displacements       |
| `op2.velocities`            | Nodal velocities          |
| `op2.accelerations`         | Nodal accelerations       |
| `op2.spc_forces`            | SPC reaction forces       |
| `op2.load_vectors`          | Applied load vectors      |
| `op2.cquad4_stress`         | CQUAD4 plate stress       |
| `op2.ctria3_stress`         | CTRIA3 plate stress       |
| `op2.chexa_stress`          | CHEXA solid stress        |
| `op2.cpenta_stress`         | CPENTA solid stress       |
| `op2.ctetra_stress`         | CTETRA solid stress       |
| `op2.cbar_stress`           | CBAR stress               |
| `op2.cbar_force`            | CBAR element forces       |
| `op2.cbeam_stress`          | CBEAM stress              |
| `op2.grid_point_forces`     | Grid point force balance  |
| `op2.eigenvalues`           | Eigenvalue tables         |
| `op2.grid_point_weight`     | Grid point weight gen.    |
| `op2.modal_effective_mass_fraction` | MEFFMASS fractions (6 x nmodes) |
| `op2.modal_effective_mass`  | MEFFMASS effective mass (6 x nmodes) |
| `op2.modal_participation_factors` | MEFFMASS participation factors (6 x nmodes) |
| `op2.modal_effective_weight` | MEFFMASS effective weight (6 x nmodes) |
| `op2.total_effective_mass_matrix` | MEFFMASS total eff. mass (6 x 1) |
| `op2.effective_mass_matrix` | MEFFMASS eff. mass matrix (6 x 6) |
| `op2.rigid_body_mass_matrix` | MEFFMASS rigid body mass (6 x 6) |

### Extracting displacement data

```python
disp = op2.displacements[1]        # subcase 1
print(disp.node_gridtype)          # (n, 2) — node IDs & grid type
print(disp.data.shape)             # (ntimes, nnodes, 6) — t1,t2,t3,r1,r2,r3
print(disp.data[0, :, 2])          # all T3 displacements, time step 0

# DataFrame access
df = disp.data_frame
```

### Extracting plate stress

```python
stress = op2.cquad4_stress[1]      # subcase 1
print(stress.data.shape)           # (ntimes, n, 8)
# Columns: fiber_distance, oxx, oyy, txy, angle, omax, omin, von_mises
print(stress.element_node)         # (n, 2) — element ID, node ID
```

### Accessing MEFFMASS matrices

Requires `MEFFMASS(PLOT)` in case control. Each property returns a
`Matrix` object (`.data` may be dense or sparse):

```python
import scipy.sparse

meff_frac = op2.modal_effective_mass_fraction  # (6, nmodes)
if meff_frac is not None:
    data = meff_frac.data
    if scipy.sparse.issparse(data):
        data = data.toarray()
    per_mode = data.T  # (nmodes, 6) — rows=modes, cols=T1-R3
```

See `references/op2-results-reference.md` for full details on every
result type, data shapes, and column ordering.

## Mass Properties

```python
from pyNastran.bdf.mesh_utils.mass_properties import mass_properties

mass, cg, inertia = mass_properties(model)
# mass: float, cg: (3,) array, inertia: (6,) — Ixx, Iyy, Izz, Ixy, Ixz, Iyz
```

Breakdown by property ID:

```python
from pyNastran.bdf.mesh_utils.mass_properties import mass_properties_breakdown
result = mass_properties_breakdown(model)
```

## Reference Files

For detailed API docs, card references, and solution type details:

- `references/bdf-card-reference.md` — every BDF card type, fields, access patterns
- `references/op2-results-reference.md` — result data shapes, column names, DataFrame access
- `references/model-building-api.md` — all `add_*` method signatures
- `references/mesh-utilities.md` — mass properties, free edges, merge, renumber
- `references/case-control-reference.md` — subcase management, output requests
- `references/coordinate-systems.md` — CORD types, transforms, GRID cp/cd
- `references/common-pitfalls.md` — 15 gotchas and their fixes
- `references/nastran-solution-types.md` — SOL-specific required cards and outputs

## Example Scripts

Complete runnable examples in `examples/`:

- `read_modify_write_bdf.py` — read, modify properties/loads, write
- `build_model_from_scratch.py` — create a plate model programmatically
- `extract_op2_displacements.py` — displacement extraction (static + modal)
- `extract_op2_stresses.py` — plate/solid stress, max von Mises
- `mass_properties_report.py` — mass / CG / inertia breakdown
- `modal_results_extraction.py` — eigenvalues and mode shapes
- `composite_layup_definition.py` — PCOMP / MAT8 laminate setup
- `rbe_and_constraints.py` — RBE2, RBE3, SPC, MPC, CONM2

## CLI Tools

Standalone scripts in `scripts/` (require `pip install pyNastran`):

- `bdf_summary.py` — print model statistics
- `extract_results.py` — OP2 → CSV result extraction
- `mass_report.py` — mass breakdown by property ID
- `find_free_edges.py` — detect unconnected shell edges
- `renumber_model.py` — renumber nodes/elements with offset
- `convert_units.py` — convert between unit systems
