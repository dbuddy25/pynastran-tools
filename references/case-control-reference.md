# Case Control Deck Reference

Programmatic access to the Nastran case control deck through pyNastran.

---

## Overview

The case control deck defines subcases, output requests, and analysis
configuration. In pyNastran it is represented by a `CaseControlDeck`
object accessible via `model.case_control_deck`.

---

## Creating the Case Control Deck

When building from scratch, pyNastran creates the case control deck
automatically. When reading a BDF, it is parsed from the file.

```python
from pyNastran.bdf.bdf import BDF

model = BDF()
model.sol = 101  # static analysis

# Access (auto-created if needed)
cc = model.case_control_deck
```

---

## Subcases

### Subcase 0 (Global)

Subcase 0 is the global subcase — entries here apply to all subcases
unless overridden.

```python
# Access global subcase
global_sub = cc.subcases[0]
global_sub.add_integer_type('ECHO', 'NONE')
```

### Creating Subcases

```python
subcase1 = cc.create_new_subcase(1)
subcase2 = cc.create_new_subcase(2)
```

### Adding Entries

```python
# Integer entries (LOAD, SPC, MPC, METHOD, etc.)
subcase1.add('LOAD', 100, options=[], option_type='')
subcase1.add('SPC', 1, options=[], option_type='')

# String entries (SUBTITLE, LABEL, etc.)
subcase1.add('SUBTITLE', 'Static Load Case 1', options=[], option_type='')
subcase1.add('LABEL', 'Gravity + Pressure', options=[], option_type='')

# Output requests (DISPLACEMENT, STRESS, STRAIN, FORCE, SPCFORCES, etc.)
subcase1.add('DISPLACEMENT', 'ALL', options=['SORT1', 'REAL'],
             option_type='STRESS-type')
subcase1.add('STRESS', 'ALL', options=['SORT1', 'REAL', 'VONMISES', 'BILIN'],
             option_type='STRESS-type')
subcase1.add('SPCFORCES', 'ALL', options=['SORT1', 'REAL'],
             option_type='STRESS-type')
subcase1.add('STRAIN', 'ALL', options=['SORT1', 'REAL', 'VONMISES'],
             option_type='STRESS-type')
subcase1.add('FORCE', 'ALL', options=['SORT1', 'REAL'],
             option_type='STRESS-type')
```

### Output Request Options

| Option     | Meaning                              |
|------------|--------------------------------------|
| SORT1      | Sort by node/element (default)       |
| SORT2      | Sort by frequency/time               |
| REAL       | Real output (static)                 |
| IMAG       | Imaginary part (frequency response)  |
| PHASE      | Phase angle (frequency response)     |
| VONMISES   | von Mises stress output              |
| MAXS       | Maximum shear stress                 |
| BILIN      | Bilinear stress at nodes (CQUAD4)    |
| CENTER     | Centroid-only stress                 |

### Referencing Sets

Instead of ALL, output can be limited to specific sets:

```python
model.add_set1(sid=100, ids=[1, 2, 3, 4, 5])
subcase1.add('DISPLACEMENT', 100, options=['SORT1', 'REAL'],
             option_type='STRESS-type')
```

---

## Common Case Control Entries

### Load and Constraint References

| Entry   | Description                    | Example                        |
|---------|--------------------------------|--------------------------------|
| LOAD    | Load set ID                    | `LOAD = 100`                   |
| SPC     | SPC constraint set ID          | `SPC = 1`                      |
| MPC     | MPC constraint set ID          | `MPC = 1`                      |
| METHOD  | Eigenvalue method ID (EIGRL)   | `METHOD = 10`                  |
| TSTEP   | Time step set ID               | `TSTEP = 200`                  |
| FREQ    | Frequency set ID               | `FREQ = 300`                   |
| SDAMP   | Structural damping ID          | `SDAMP = 50`                   |
| DLOAD   | Dynamic load set ID            | `DLOAD = 400`                  |
| TEMP    | Temperature set ID (for loads) | `TEMP(LOAD) = 100`             |
| DEFORM  | Enforced deformation set       | `DEFORM = 500`                 |

```python
# Integer entry pattern:
subcase.add('LOAD', 100, options=[], option_type='')
subcase.add('SPC', 1, options=[], option_type='')
subcase.add('METHOD', 10, options=[], option_type='')
```

### Output Requests

| Entry        | Description                |
|--------------|----------------------------|
| DISPLACEMENT | Nodal displacements        |
| VELOCITY     | Nodal velocities           |
| ACCELERATION | Nodal accelerations        |
| STRESS       | Element stresses           |
| STRAIN       | Element strains            |
| FORCE        | Element forces             |
| SPCFORCES    | SPC reaction forces        |
| MPCFORCES    | MPC reaction forces        |
| GPFORCE      | Grid point force balance   |
| ESE          | Element strain energy      |
| OLOAD        | Applied loads              |

### Analysis Configuration

| Entry     | Description                    |
|-----------|--------------------------------|
| SUBTITLE  | Subcase subtitle               |
| LABEL     | Subcase label                  |
| ANALYSIS  | Analysis type override         |
| SUBSEQ    | Subsequent subcase             |

---

## Modal Analysis (SOL 103)

```python
model.sol = 103

# Add eigenvalue method
model.add_eigrl(sid=10, v1=0., v2=1000., nd=20)

subcase = cc.create_new_subcase(1)
subcase.add('METHOD', 10, options=[], option_type='')
subcase.add('SPC', 1, options=[], option_type='')
subcase.add('DISPLACEMENT', 'ALL', options=['SORT1', 'REAL'],
            option_type='STRESS-type')
```

---

## Buckling Analysis (SOL 105)

```python
model.sol = 105

model.add_eigrl(sid=20, nd=5)  # buckling modes

# Subcase 1: static preload
subcase1 = cc.create_new_subcase(1)
subcase1.add('LOAD', 100, options=[], option_type='')
subcase1.add('SPC', 1, options=[], option_type='')
subcase1.add('STRESS', 'ALL', options=['SORT1', 'REAL', 'VONMISES'],
             option_type='STRESS-type')

# Subcase 2: buckling
subcase2 = cc.create_new_subcase(2)
subcase2.add('METHOD', 20, options=[], option_type='')
subcase2.add('SPC', 1, options=[], option_type='')
subcase2.add('DISPLACEMENT', 'ALL', options=['SORT1', 'REAL'],
             option_type='STRESS-type')
```

---

## Frequency Response (SOL 111)

```python
model.sol = 111

model.add_eigrl(sid=10, v1=0., v2=500., nd=50)

subcase = cc.create_new_subcase(1)
subcase.add('METHOD', 10, options=[], option_type='')
subcase.add('SPC', 1, options=[], option_type='')
subcase.add('DLOAD', 400, options=[], option_type='')
subcase.add('FREQUENCY', 300, options=[], option_type='')
subcase.add('DISPLACEMENT', 'ALL', options=['SORT1', 'REAL', 'IMAG'],
            option_type='STRESS-type')
subcase.add('STRESS', 'ALL', options=['SORT1', 'REAL', 'IMAG', 'VONMISES'],
            option_type='STRESS-type')
```

---

## Reading Existing Case Control

```python
model = BDF()
model.read_bdf('model.bdf')

cc = model.case_control_deck

# List subcases
for subcase_id, subcase in sorted(cc.subcases.items()):
    print(f"Subcase {subcase_id}")
    print(subcase)

# Get a specific entry
if subcase.has_parameter('LOAD'):
    load_id = subcase.get_parameter('LOAD')
    print(f"LOAD = {load_id}")
```

---

## Deleting Subcases

```python
cc.delete_subcase(2)  # remove subcase 2
```

---

## PARAMs

PARAMs are bulk data entries, not case control, but are closely related
to analysis configuration:

```python
model.add_param('POST', [-1])         # request OP2 output
model.add_param('AUTOSPC', ['YES'])   # auto-SPC singular DOFs
model.add_param('PRTMAXIM', ['YES'])  # print max values
model.add_param('GRDPNT', [0])        # grid point weight output
model.add_param('WTMASS', [1.0])      # mass-to-weight conversion
model.add_param('COUPMASS', [1])      # coupled mass matrix
```
