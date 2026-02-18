# Mesh Utilities Reference

pyNastran provides mesh manipulation utilities in
`pyNastran.bdf.mesh_utils`.

---

## Mass Properties

### mass_properties

Compute total mass, CG, and inertia from a BDF model (lumped approach).

```python
from pyNastran.bdf.mesh_utils.mass_properties import mass_properties

mass, cg, inertia = mass_properties(model)
# mass: float — total mass
# cg: ndarray (3,) — center of gravity [x, y, z]
# inertia: ndarray (6,) — [Ixx, Iyy, Izz, Ixy, Ixz, Iyz]
```

Optionally filter by element IDs:

```python
mass, cg, inertia = mass_properties(model, element_ids=[1, 2, 3])
```

Or by reference point:

```python
mass, cg, inertia = mass_properties(model, reference_point=[0., 0., 0.])
```

### mass_properties_breakdown

Mass breakdown by property ID (and optionally by material ID).

```python
from pyNastran.bdf.mesh_utils.mass_properties import mass_properties_breakdown

result = mass_properties_breakdown(model)
# Returns dict with per-PID mass breakdown
```

### mass_properties_no_xref

Compute mass properties without cross-referencing (simpler, faster).

```python
from pyNastran.bdf.mesh_utils.mass_properties import mass_properties_no_xref

mass, cg, inertia = mass_properties_no_xref(model)
```

---

## Free Edges

Find unconnected edges on shell elements (indicates mesh discontinuities).

```python
from pyNastran.bdf.mesh_utils.free_edges import free_edges

edges = free_edges(model)
# Returns list of (nid1, nid2) tuples representing free edges
```

Usage pattern:

```python
model = BDF()
model.read_bdf('model.bdf')
model.cross_reference()

edges = free_edges(model)
if edges:
    print(f"Found {len(edges)} free edges!")
    for n1, n2 in edges:
        print(f"  Edge: {n1} - {n2}")
else:
    print("No free edges (mesh is watertight)")
```

---

## Skin Solid Elements

Extract the skin (outer faces) of solid elements as shell elements.

```python
from pyNastran.bdf.mesh_utils.skin_solid_elements import skin_solid_elements

skin_model = skin_solid_elements(model)
# Returns a new BDF with CQUAD4/CTRIA3 elements on the solid surfaces
```

---

## Mesh Equivalencing (Node Merging)

Merge coincident nodes within a tolerance.

```python
from pyNastran.bdf.mesh_utils.collapse_bad_quads import collapse_bad_quads
from pyNastran.bdf.mesh_utils.delete_bad_elements import delete_bad_elements

# Typically done through:
from pyNastran.bdf.mesh_utils.bdf_equivalence import bdf_equivalence_nodes

bdf_equivalence_nodes(bdf_filename, bdf_filename_out, tol=0.001,
                      renumber_nodes=False, neq_max=4, xref=True)
```

- **tol**: float — coincidence tolerance
- **neq_max**: int — max equivalencing iterations
- **renumber_nodes**: bool — renumber after merging

---

## Renumber

Renumber node and element IDs with offsets.

```python
from pyNastran.bdf.mesh_utils.bdf_renumber import bdf_renumber

bdf_renumber(bdf_filename, bdf_filename_out,
             size=8, is_double=False,
             starting_id_dict=None)
```

Using `starting_id_dict`:

```python
starting_id_dict = {
    'nid': 1000,       # start node IDs at 1000
    'eid': 2000,       # start element IDs at 2000
    'pid': 100,        # start property IDs at 100
    'mid': 100,        # start material IDs at 100
    'cid': 10,         # start coord IDs at 10
    'spc_id': 1,
    'mpc_id': 1,
    'load_id': 1,
}
bdf_renumber(bdf_filename, bdf_filename_out,
             starting_id_dict=starting_id_dict)
```

---

## Merge Models

Merge two BDF files into one.

```python
from pyNastran.bdf.mesh_utils.bdf_merge import bdf_merge

bdf_merge(bdf_filenames, bdf_filename_out,
          renumber=True, encoding=None, size=8, is_double=False)
```

- **bdf_filenames**: list[str] — input BDF files to merge
- **renumber**: bool — auto-renumber to avoid ID conflicts

---

## Mirror

Mirror a model about a plane.

```python
from pyNastran.bdf.mesh_utils.mirror_mesh import bdf_mirror

bdf_mirror(bdf_filename, bdf_filename_out, plane='xz',
           size=8, is_double=False)
```

- **plane**: str — mirror plane: 'xy', 'yz', or 'xz'

---

## Convert Units

Convert model between unit systems.

```python
from pyNastran.bdf.mesh_utils.convert import convert

# Define unit scale factors
units_from = {'length': 'mm', 'mass': 'kg', 'time': 's'}
units_to = {'length': 'm', 'mass': 'kg', 'time': 's'}

# Or specify scale factors directly:
xyz_scale = 0.001      # mm → m
mass_scale = 1.0       # kg → kg
time_scale = 1.0       # s → s

convert(model, xyz_scale=xyz_scale)
```

---

## Remove Unused Cards

Remove unused properties, materials, coordinate systems, etc.

```python
from pyNastran.bdf.mesh_utils.remove_unused import remove_unused

model = BDF()
model.read_bdf('model.bdf')
remove_unused(model)
model.write_bdf('cleaned.bdf')
```

---

## Get Element Faces

Get faces of elements for visualization or contact surface definition.

```python
# For a specific element:
elem = model.elements[eid]
faces = elem.faces  # dict of face_id → node_ids
```

---

## Summary of Import Paths

| Function                  | Module                                              |
|---------------------------|-----------------------------------------------------|
| `mass_properties`         | `pyNastran.bdf.mesh_utils.mass_properties`          |
| `mass_properties_breakdown` | `pyNastran.bdf.mesh_utils.mass_properties`        |
| `free_edges`              | `pyNastran.bdf.mesh_utils.free_edges`               |
| `skin_solid_elements`     | `pyNastran.bdf.mesh_utils.skin_solid_elements`      |
| `bdf_equivalence_nodes`   | `pyNastran.bdf.mesh_utils.bdf_equivalence`          |
| `bdf_renumber`            | `pyNastran.bdf.mesh_utils.bdf_renumber`             |
| `bdf_merge`               | `pyNastran.bdf.mesh_utils.bdf_merge`                |
| `bdf_mirror`              | `pyNastran.bdf.mesh_utils.mirror_mesh`              |
| `convert`                 | `pyNastran.bdf.mesh_utils.convert`                  |
| `remove_unused`           | `pyNastran.bdf.mesh_utils.remove_unused`            |
