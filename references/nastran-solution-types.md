# Nastran Solution Types Reference

Solution types supported by pyNastran with required cards and expected
outputs. Heat transfer solutions (SOL 153/159) are out of scope.

---

## SOL 101 — Linear Static

**Purpose:** Static stress/displacement under applied loads.

**Required bulk data:**
- GRID, elements, properties, materials
- Load cards (FORCE, PLOAD4, GRAV, etc.)
- SPC/SPC1 constraints
- PARAM,POST,-1 (for OP2 output)

**Case control:**
```
SUBCASE 1
  LOAD = 100
  SPC = 1
  DISPLACEMENT(SORT1,REAL) = ALL
  STRESS(SORT1,REAL,VONMISES,BILIN) = ALL
  SPCFORCES(SORT1,REAL) = ALL
```

**OP2 outputs:** displacements, stresses (per element type), spc_forces,
load_vectors, grid_point_forces (if requested).

```python
model.sol = 101
```

---

## SOL 103 — Normal Modes (Real Eigenvalue)

**Purpose:** Natural frequencies and mode shapes.

**Required bulk data:**
- GRID, elements, properties, materials (with rho for mass)
- SPC/SPC1 constraints
- EIGRL or EIGR (eigenvalue extraction method)

**Case control:**
```
SUBCASE 1
  METHOD = 10
  SPC = 1
  DISPLACEMENT(SORT1,REAL) = ALL
```

**OP2 outputs:** eigenvalues (frequency table), eigenvectors (mode
shapes), optionally stresses.

```python
model.sol = 103
model.add_eigrl(sid=10, v1=0., v2=1000., nd=20)
```

---

## SOL 105 — Buckling

**Purpose:** Linear buckling eigenvalues and mode shapes.

**Required bulk data:**
- Same as SOL 101 (structural model + loads + constraints)
- EIGRL for buckling mode extraction

**Case control:**
```
SUBCASE 1
  LOAD = 100
  SPC = 1
  STRESS(SORT1,REAL,VONMISES) = ALL
SUBCASE 2
  METHOD = 20
  SPC = 1
  DISPLACEMENT(SORT1,REAL) = ALL
```

Subcase 1 = static preload. Subcase 2 = buckling extraction.
The buckling eigenvalue λ is the load multiplier (buckling occurs at
λ × applied load).

```python
model.sol = 105
model.add_eigrl(sid=20, nd=5)
```

---

## SOL 106 — Nonlinear Static

**Purpose:** Geometrically or materially nonlinear static analysis.

**Required bulk data:**
- Same as SOL 101
- NLPARM (nonlinear parameters: increments, convergence)
- Optionally MATS1 (plasticity), TABLES1 (stress-strain curves)

**Case control:**
```
SUBCASE 1
  LOAD = 100
  SPC = 1
  NLPARM = 50
  DISPLACEMENT(SORT1,REAL) = ALL
  STRESS(SORT1,REAL,VONMISES) = ALL
```

**OP2 outputs:** displacements (at each load increment), stresses,
nonlinear stress/strain.

```python
model.sol = 106
model.add_nlparm(nlparm_id=50, ninc=10, dt=0., kmethod='AUTO',
                 kstep=5, max_iter=25, conv='PW')
```

---

## SOL 108 — Direct Frequency Response

**Purpose:** Steady-state response to harmonic excitation (direct method).

**Required bulk data:**
- Structural model
- DLOAD referencing RLOAD1/RLOAD2 + DAREA
- FREQ/FREQ1/FREQ2 (frequency list)
- SPC constraints

**Case control:**
```
SUBCASE 1
  DLOAD = 400
  SPC = 1
  FREQUENCY = 300
  DISPLACEMENT(SORT1,REAL,IMAG) = ALL
  STRESS(SORT1,REAL,IMAG,VONMISES) = ALL
```

**OP2 outputs:** complex displacements, complex stresses.

```python
model.sol = 108
```

---

## SOL 109 — Direct Transient Response

**Purpose:** Time-domain response using direct integration.

**Required bulk data:**
- Structural model
- DLOAD referencing TLOAD1/TLOAD2 + DAREA
- TSTEP (time step definition)
- SPC constraints

**Case control:**
```
SUBCASE 1
  DLOAD = 400
  SPC = 1
  TSTEP = 200
  DISPLACEMENT(SORT1,REAL) = ALL
```

```python
model.sol = 109
model.add_tstep(sid=200, N=[100], DT=[0.001], NO=[1])
```

---

## SOL 110 — Modal Complex Eigenvalue

**Purpose:** Complex eigenvalues for stability analysis (e.g., brake
squeal, flutter).

**Required bulk data:**
- Structural model with damping (PARAM,G or element GE)
- EIGRL for real modes
- EIGC for complex eigenvalue extraction

**Case control:**
```
SUBCASE 1
  METHOD = 10
  CMETHOD = 30
  SPC = 1
```

```python
model.sol = 110
```

---

## SOL 111 — Modal Frequency Response

**Purpose:** Frequency response using modal superposition (faster than
SOL 108 for many DOFs).

**Required bulk data:**
- Structural model
- EIGRL (real modes as basis)
- DLOAD referencing RLOAD1/RLOAD2 + DAREA
- FREQ/FREQ1/FREQ2
- SPC constraints

**Case control:**
```
SUBCASE 1
  METHOD = 10
  DLOAD = 400
  SPC = 1
  FREQUENCY = 300
  DISPLACEMENT(SORT1,PHASE) = ALL
  STRESS(SORT1,PHASE,VONMISES) = ALL
```

```python
model.sol = 111
model.add_eigrl(sid=10, v1=0., v2=2000., nd=100)
model.add_freq1(sid=300, f1=10., df=1., ndf=490)
```

---

## SOL 112 — Modal Transient Response

**Purpose:** Transient response using modal superposition.

**Required bulk data:**
- Structural model
- EIGRL (real modes)
- DLOAD referencing TLOAD1/TLOAD2 + DAREA
- TSTEP

**Case control:**
```
SUBCASE 1
  METHOD = 10
  DLOAD = 400
  SPC = 1
  TSTEP = 200
  DISPLACEMENT(SORT1,REAL) = ALL
```

```python
model.sol = 112
```

---

## SOL 144 — Static Aeroelasticity

**Purpose:** Static aeroelastic trim analysis.

**Required bulk data:**
- Structural model + aerodynamic panels (CAERO1, PAERO1, SPLINE1)
- TRIM card (trim variables)
- AEROS (reference geometry)

```python
model.sol = 144
```

---

## SOL 145 — Flutter (Aeroelastic)

**Purpose:** Flutter analysis (V-g, V-f plots).

**Required bulk data:**
- Structural model + aero panels
- EIGRL (structural modes)
- FLFACT (Mach, density ratio, velocity)
- FLUTTER card
- AERO (reference geometry)

```python
model.sol = 145
```

---

## SOL 146 — Dynamic Aeroelasticity

**Purpose:** Gust response and buffet analysis.

**Required bulk data:**
- Structural model + aero panels
- EIGRL + GUST card
- AERO

```python
model.sol = 146
```

---

## SOL 200 — Design Optimization

**Purpose:** Structural optimization (sizing, shape, topology).

**Required bulk data:**
- Base structural model (like SOL 101/103)
- DESVAR (design variables)
- DVPREL1/DVPREL2 (property relationships)
- DRESP1/DRESP2 (design responses: stress, displacement, mass, freq)
- DCONSTR (design constraints)
- DOPTPRM (optimization parameters)

**Case control:**
```
SUBCASE 1
  ANALYSIS = STATICS
  LOAD = 100
  SPC = 1
  DESSUB = 10
  DISPLACEMENT = ALL
  STRESS = ALL
```

```python
model.sol = 200
model.add_desvar(desvar_id=1, label='THICK', xinit=0.005,
                 xlb=0.001, xub=0.05)
model.add_dvprel1(dvprel_id=1, prop_type='PSHELL', pid=1,
                  pname_fid='T', desvar_ids=[1], coeffs=[1.0])
model.add_dresp1(dresp_id=1, label='MASS', response_type='WEIGHT')
model.add_dresp1(dresp_id=2, label='STRESS', response_type='STRESS',
                 property_type='PSHELL', region=None, atta=9,
                 attb=None, atti=[1])
```

---

## Summary Table

| SOL | Name                    | Key Cards                    | Primary Outputs         |
|-----|-------------------------|------------------------------|-------------------------|
| 101 | Linear Static           | LOAD, SPC                    | disp, stress, forces    |
| 103 | Normal Modes            | METHOD(EIGRL), SPC           | eigenvalues, eigvecs    |
| 105 | Buckling                | LOAD + METHOD(EIGRL)         | buckling factors, modes |
| 106 | Nonlinear Static        | LOAD, NLPARM, SPC            | disp, stress (NL)       |
| 108 | Direct Freq Response    | DLOAD, FREQ, SPC             | complex disp/stress     |
| 109 | Direct Transient        | DLOAD, TSTEP, SPC            | time-history disp       |
| 110 | Modal Complex Eigenval  | METHOD, CMETHOD              | complex eigenvalues     |
| 111 | Modal Freq Response     | METHOD, DLOAD, FREQ, SPC     | complex disp/stress     |
| 112 | Modal Transient         | METHOD, DLOAD, TSTEP, SPC    | time-history disp       |
| 144 | Static Aeroelasticity   | TRIM, CAERO, AEROS           | trim results            |
| 145 | Flutter                 | FLUTTER, FLFACT, AERO        | V-g, V-f plots          |
| 146 | Dynamic Aeroelasticity  | GUST, AERO, METHOD           | gust response           |
| 200 | Optimization            | DESVAR, DVPREL, DRESP, DCONSTR | optimized design      |
