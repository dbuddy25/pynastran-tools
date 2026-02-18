# Common pyNastran Pitfalls

## 1. Accessing `_ref` attributes without cross-referencing

**Problem:** `node.cp_ref` raises `AttributeError` because `_ref`
attributes are only populated after cross-referencing.

```python
# WRONG
model.read_bdf('model.bdf')
cp = model.nodes[1].cp_ref  # AttributeError

# RIGHT
model.read_bdf('model.bdf')
model.cross_reference()
cp = model.nodes[1].cp_ref  # now works
```

## 2. Writing a cross-referenced model

**Problem:** `write_bdf` produces corrupt output with object reprs
instead of IDs when the model is still cross-referenced.

```python
# WRONG
model.cross_reference()
model.write_bdf('out.bdf')  # garbage output

# RIGHT
model.cross_reference()
# ... do work ...
model.uncross_reference()
model.write_bdf('out.bdf')
```

## 3. Stress results are per element type, not one big dict

**Problem:** Looking for all stresses in a single `op2.stress` dict.
Stresses are split by element type.

```python
# WRONG
stress = op2.stress[1]              # no such attribute

# RIGHT
quad_stress = op2.cquad4_stress[1]  # CQUAD4 stresses for subcase 1
tria_stress = op2.ctria3_stress[1]  # CTRIA3 stresses for subcase 1
hexa_stress = op2.chexa_stress[1]   # CHEXA stresses for subcase 1
```

## 4. 0-based data array vs 1-based IDs

**Problem:** Using element/node IDs as indices into `data` arrays.
The `data` array is 0-based; use `node_gridtype` or `element_node`
to map IDs to indices.

```python
disp = op2.displacements[1]
node_ids = disp.node_gridtype[:, 0]
idx = np.where(node_ids == 100)[0][0]
tz = disp.data[0, idx, 2]  # T3 for node 100, first time step
```

## 5. BILINEAR vs CENTROID stress output

**Problem:** Assuming plate stress always has one row per element.
With BILINEAR output (default for CQUAD4), there is one centroid row
plus one row per corner node (5 rows per element).

```python
stress = op2.cquad4_stress[1]
# stress.element_node has shape (n, 2)
# Column 0: element ID, Column 1: node ID (0 = centroid)

# To get centroid-only results:
centroid_mask = stress.element_node[:, 1] == 0
centroid_stress = stress.data[:, centroid_mask, :]
```

## 6. Composite layer indexing

**Problem:** PCOMP plies are 1-indexed in Nastran but 0-indexed in
pyNastran's internal list.

```python
pcomp = model.properties[10]  # PCOMP
t_ply0 = pcomp.thicknesses[0]   # first ply (Nastran ply 1)
theta0 = pcomp.thetas[0]        # first ply angle
nplies = pcomp.nplies
```

## 7. Eigenvalue dict keys

**Problem:** `op2.eigenvalues` keys are **strings** (the eigenvalue
table title), not integers.

```python
# WRONG
eigvals = op2.eigenvalues[1]

# RIGHT
for title, eigval_table in op2.eigenvalues.items():
    print(title)                    # e.g. 'EIGENVALUE'
    modes = eigval_table.mode       # array of mode numbers
    freqs = eigval_table.freq       # array of frequencies (Hz)
    radians = eigval_table.radians  # array of circular frequencies
```

## 8. PCOMP ply numbering and MID references

**Problem:** Assuming all plies share one material. Each ply in PCOMP
can have a different MID.

```python
pcomp = model.properties[10]
for i in range(pcomp.nplies):
    mid = pcomp.material_ids[i]
    thick = pcomp.thicknesses[i]
    theta = pcomp.thetas[i]
```

## 9. RBE3 vs RBE2 semantics

**Problem:** Confusing RBE2 (rigid) and RBE3 (interpolation).

- **RBE2**: Independent node drives dependent nodes rigidly.
  `model.add_rbe2(eid, gn=independent_nid, cm='123456', Gmi=[dep1, dep2, ...])`
- **RBE3**: Dependent (reference) node motion is weighted average of independent nodes.
  `model.add_rbe3(eid, refgrid=dep_nid, refc='123456', weights=[1.0], comps=['123'], Gijs=[[ind1, ind2, ...]])`

RBE2 adds stiffness; RBE3 does not. Use RBE3 for load distribution.

## 10. `model.loads[sid]` returns a list, not a single card

**Problem:** Treating `model.loads[sid]` as a single load card.
A load set ID can contain multiple load cards.

```python
for load_card in model.loads[100]:
    print(type(load_card))  # FORCE, PLOAD4, MOMENT, etc.
```

## 11. LOAD combination card

**Problem:** Forgetting that the LOAD card is a combination card that
references other load sets, not a direct load. It lives in `model.load_combinations`.

```python
# The LOAD card:   LOAD, SID, S, S1, L1, S2, L2, ...
# It scales and combines other load sets.
# Accessed via:
for combo in model.load_combinations[sid]:
    print(combo.scale_factors)
    print(combo.load_ids)
```

## 12. TEMPD default temperature

**Problem:** Forgetting that TEMPD sets the default temperature for
nodes not explicitly listed in a TEMP card. Without TEMPD, unlisted
nodes have undefined temperature.

```python
model.add_tempd(sid=100, temperature=20.0)  # default temp
model.add_temp(sid=100, nodes=[1, 2], temperatures=[150., 200.])
```

## 13. Mass property differences (lumped vs coupled)

**Problem:** `mass_properties()` uses a simplified lumped-mass
approach. Results may differ slightly from Nastran's GPWG (which uses
the full mass matrix). For accurate comparison, use `op2.grid_point_weight`.

```python
# Approximate (pyNastran lumped)
from pyNastran.bdf.mesh_utils.mass_properties import mass_properties
mass, cg, I = mass_properties(model)

# Exact (from Nastran run)
gpw = op2.grid_point_weight[0]  # GridPointWeight object
print(gpw.mass)     # total mass
print(gpw.cg)       # center of gravity
print(gpw.IS)       # inertia at CG
```

## 14. Card format: 8-char vs 16-char fields

**Problem:** Writing cards that exceed 8-character field width without
switching to large-field format. Nastran truncates values silently.

```python
# pyNastran handles this automatically when you use write_bdf()
# But when using write_bdf with is_double=True, you get 16-char fields:
model.write_bdf('out.bdf', is_double=True)

# Individual card:
node = model.nodes[1]
print(node.write_card(size=16))  # large field format
```

## 15. Duplicate ID handling

**Problem:** Adding a card with an ID that already exists silently
overwrites the old card (for most card types).

```python
model.add_grid(nid=1, xyz=[0., 0., 0.])
model.add_grid(nid=1, xyz=[1., 1., 1.])  # overwrites the first!
print(model.nodes[1].xyz)                 # [1., 1., 1.]

# Check before adding:
if nid not in model.nodes:
    model.add_grid(nid=nid, xyz=xyz)
```

## Summary Checklist

1. Always `cross_reference()` before accessing `_ref` attributes
2. Always `uncross_reference()` before `write_bdf()`
3. Use element-type-specific stress dicts (`cquad4_stress`, etc.)
4. Map IDs to array indices via `node_gridtype` / `element_node`
5. Account for BILINEAR multi-row stress output
6. Remember pyNastran uses 0-based ply indexing
7. `op2.eigenvalues` keys are strings
8. Each PCOMP ply can have its own MID
9. RBE2 = rigid, RBE3 = interpolation (no stiffness)
10. `model.loads[sid]` is always a list
11. LOAD card is a combination — check `load_combinations`
12. Use TEMPD for default node temperatures
13. `mass_properties()` is approximate; GPWG is exact
14. Use `size=16` or `is_double=True` for high-precision output
15. Duplicate IDs silently overwrite — check before adding
