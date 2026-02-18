# Coordinate Systems Reference

Nastran coordinate system types and their usage in pyNastran.

---

## Overview

Coordinate systems define local frames for:
- **GRID cp**: input coordinate system (defines where the node is)
- **GRID cd**: output/displacement coordinate system
- **Element MCID**: material coordinate direction
- **Load cid**: load application direction

All coordinate systems are stored in `model.coords` and accessed by
their CID (coordinate ID). CID 0 is always the basic (global)
rectangular system.

---

## CORD2R — Rectangular

Defined by origin + two points (z-axis point and xz-plane point).

```python
model.add_cord2r(cid=1,
                 origin=[10., 0., 0.],     # point A (origin)
                 zaxis=[10., 0., 1.],      # point B (on z-axis)
                 xzplane=[11., 0., 0.],    # point C (in x-z plane)
                 rid=0)                     # reference coord (0=basic)
```

**Axes:**
- Z = B - A (normalized)
- Y = Z × (C - A) (normalized)
- X = Y × Z

---

## CORD2C — Cylindrical

Same definition points as CORD2R, but coordinates are (R, θ, Z).

```python
model.add_cord2c(cid=2,
                 origin=[0., 0., 0.],
                 zaxis=[0., 0., 1.],
                 xzplane=[1., 0., 0.],
                 rid=0)
```

**Coordinates:** (R, θ°, Z) where θ is measured from the local x-axis.

---

## CORD2S — Spherical

Same definition points as CORD2R, but coordinates are (R, θ, φ).

```python
model.add_cord2s(cid=3,
                 origin=[0., 0., 0.],
                 zaxis=[0., 0., 1.],
                 xzplane=[1., 0., 0.],
                 rid=0)
```

**Coordinates:** (R, θ°, φ°) where θ is co-latitude from z-axis.

---

## CORD1R / CORD1C / CORD1S

Defined by 3 grid points (instead of 3 coordinate triples).
Less common; mainly used when the coord system must move with the mesh.

```python
# CORD1R: defined by grid points G1, G2, G3
# G1 = origin, G2 = z-axis point, G3 = xz-plane point
# Access: model.coords[cid]
```

---

## Using Coordinate Systems with GRIDs

### Input Coordinate System (cp)

The `cp` field defines the coordinate system in which the GRID xyz
values are specified. During cross-referencing, pyNastran converts
these to the basic (global) system.

```python
# Node at (R=10, θ=45°, Z=5) in cylindrical coord 2
model.add_grid(nid=100, xyz=[10., 45., 5.], cp=2)
```

### Output/Displacement Coordinate System (cd)

The `cd` field defines the coordinate system for displacement output.
When cd ≠ 0, displacements/forces are output in the local system.

```python
# Node with output in local cylindrical system
model.add_grid(nid=100, xyz=[10., 45., 5.], cp=2, cd=2)
```

---

## Coordinate Transforms

### Getting Global Position

After cross-referencing:

```python
model.cross_reference()
node = model.nodes[100]

# Position in global (basic) coordinate system
xyz_global = node.get_position()

# Position in a specific coordinate system
xyz_local = node.get_position_wrt(model, cid=1)
```

### Transformation Matrix

```python
coord = model.coords[1]

# Rotation matrix from local to global (3x3)
beta = coord.beta()

# Transform a vector from local to global
v_global = beta @ v_local

# Transform from global to local
v_local = beta.T @ v_global

# Transform a point (includes origin offset)
xyz_global = coord.transform_node_to_global(xyz_local)
xyz_local = coord.transform_node_to_local(xyz_global)
```

---

## Material Coordinate System (MCID)

Shell elements (CQUAD4, CTRIA3, etc.) have a `theta_mcid` field.

- If float: material angle in degrees relative to element x-axis
- If int: reference coordinate system ID for material direction

```python
# Material angle = 45° from element x-axis
model.add_cquad4(eid=1, pid=1, nids=[1,2,3,4], theta_mcid=45.0)

# Material direction from coord system 1
model.add_cquad4(eid=2, pid=1, nids=[5,6,7,8], theta_mcid=1)
```

---

## Common Patterns

### Cylindrical Loads

Apply a force in the radial direction using a cylindrical coord:

```python
# Create cylindrical system at center
model.add_cord2c(cid=10,
                 origin=[0., 0., 0.],
                 zaxis=[0., 0., 1.],
                 xzplane=[1., 0., 0.])

# Force in radial direction (component 1 in cylindrical = R)
model.add_force(sid=100, node=50, mag=1000.,
                xyz=[1., 0., 0.], cid=10)
```

### Symmetric Boundary Conditions

For symmetry about the XZ plane (Y=0), constrain Y-translation and
X/Z-rotation:

```python
model.add_spc1(sid=1, components='246', nodes=symmetry_nodes)
```

### Multiple Coordinate Systems

When merging models with different orientations:

```python
# Model A is in global
# Model B needs rotation: z-axis = global x-axis
model.add_cord2r(cid=100,
                 origin=[500., 0., 0.],
                 zaxis=[501., 0., 0.],     # z along global x
                 xzplane=[500., 0., 1.])   # x along global z

# All Model B nodes use cp=100
for nid in model_b_nodes:
    model.add_grid(nid=nid, xyz=local_coords[nid], cp=100)
```

---

## Querying Coordinate Systems

```python
# List all coordinate systems
for cid, coord in sorted(model.coords.items()):
    print(f"CID {cid}: {coord.type} — origin={coord.origin}")

# Check coordinate type
coord = model.coords[1]
print(coord.type)  # 'CORD2R', 'CORD2C', 'CORD2S', etc.
```
