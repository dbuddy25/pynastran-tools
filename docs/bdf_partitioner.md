# BDF Partitioner

Standalone tool that splits a monolithic Nastran BDF into component-level include files. Parts are detected automatically by flood-filling the mesh — boundaries are identified at **RBE2-CBUSH-RBE2** interfaces and **glue contact** surfaces (BCTABLE, BSURF, BCPROP/BCPROPS).

## Features

- **Automatic part detection** via flood-fill across shared nodes
- **Joint identification** — CBUSH chains and glue contact pairs between parts
- **Interactive parts table** — rename parts, review element/node/PID counts
- **Merge parts** — select 2+ parts to combine, absorbing internal joints
- **3D Preview** — pyvista visualization colored by part (optional dependency)
- **Organized output** — master.bdf, shared.bdf, per-part files, per-joint files

## Output Structure

```
output_dir/
  master.bdf              # Exec/case control + INCLUDEs + PARAMs
  shared.bdf              # Materials, properties, coordinate systems
  part_001.bdf            # GRIDs, elements, mass elements, SPCs, loads
  part_002.bdf
  part_001-to-part_002.bdf  # Boundary CBUSHes + RBE2 pairs + PBUSH
```

## Setup on Windows

### Prerequisites

- **Python 3.9+** — download from [python.org](https://www.python.org/downloads/). During install, check **"Add python.exe to PATH"**.
- **Git** (optional, for cloning) — download from [git-scm.com](https://git-scm.com/download/win)

### Step-by-step installation

1. **Clone or download the repository**

   ```cmd
   git clone https://github.com/your-org/pynastran.git
   cd pynastran
   ```

   Or download and extract the ZIP from GitHub.

2. **Create a virtual environment** (recommended)

   ```cmd
   python -m venv .venv
   .venv\Scripts\activate
   ```

3. **Install dependencies**

   ```cmd
   pip install -r requirements.txt
   ```

   This installs pyNastran, customtkinter, tksheet, numpy, scipy, and pyvista.

   > **Note:** pyvista is optional (used only for 3D Preview). If installation fails on your system, the tool will still work — the 3D Preview button will simply be disabled. To skip it: `pip install pyNastran customtkinter tksheet numpy scipy`

4. **Verify the install**

   ```cmd
   python -c "from pyNastran.bdf.bdf import BDF; print('pyNastran OK')"
   python -c "import customtkinter; print('customtkinter OK')"
   python -c "import pyvista; print('pyvista OK')"
   ```

### Running the tool

```cmd
cd preprocessing
python partition_gui.py
```

Or from the repo root:

```cmd
python preprocessing\partition_gui.py
```

### Troubleshooting on Windows

| Issue | Fix |
|-------|-----|
| `python` not recognized | Use `py` instead of `python`, or re-install Python with "Add to PATH" checked |
| `pip install pyvista` fails | Try `pip install --only-binary :all: pyvista`. If still failing, skip it — 3D Preview is optional |
| DPI scaling looks wrong | Right-click `python.exe` > Properties > Compatibility > "Override high DPI scaling" > System |
| tkinter not found | Re-install Python and ensure "tcl/tk and IDLE" is checked in the installer |
| `ModuleNotFoundError: preprocessing.bdf_utils` | Run from the `preprocessing/` directory, or add the repo root to `PYTHONPATH` |

## Usage

1. **Open BDF** — click Browse, select your main .bdf/.dat file
2. **Partition** — click the Partition button; the tool loads the model and runs flood-fill
3. **Review** — the parts table shows each detected component with element/node counts
4. **Rename** — double-click the Name column to rename parts (used in output filenames)
5. **Merge** (optional) — select 2+ rows, click Merge Selected to combine parts
6. **3D Preview** (optional) — click to open a pyvista window with parts colored by ID
7. **Set output dir** — defaults to `<bdf_name>_partitioned/` next to the input file
8. **Write** — click Write Include Files to generate the output

## Card Assignment

| Card Type | Destination |
|-----------|-------------|
| GRID | Part file (by flood-fill node ownership) |
| Structural elements (CQUAD4, CHEXA, etc.) | Part file |
| Boundary CBUSH + paired RBE2s | Joint file |
| Interior RBE2/RBE3/RBAR | Part file |
| CONM2, CMASS | Part file (by attached node) |
| PBUSH for boundary CBUSHes | Joint file |
| PSHELL, PCOMP, PSOLID, etc. | shared.bdf |
| MAT1, MAT8, etc. | shared.bdf |
| Coordinate systems | shared.bdf |
| SPC/SPC1 | Part file if all nodes in one part; else shared.bdf |
| FORCE/MOMENT/PLOAD4 | Part file if all nodes/elems in one part; else shared.bdf |
| BCTPARA/BCTPARM | shared.bdf |
| BCPROP/BCPROPS | Joint file (passthrough, mapped by PID) |
| PARAM, EIGRL | master.bdf |

## Architecture

- **`partition_bdf.py`** — pure algorithm + pyvista viz, no GUI dependencies
- **`partition_gui.py`** — standalone customtkinter GUI, imports from `partition_bdf.py`

The core `partition_model(model)` function can be used programmatically:

```python
from preprocessing.partition_bdf import partition_model, write_partition
from preprocessing.bdf_utils import make_model

model = make_model(['BCPROP', 'BCPROPS', 'BCPARA', 'BOUTPUT', 'BGPARM'])
model.read_bdf('my_model.bdf')
model.cross_reference()

result = partition_model(model)
print(f"Found {len(result.parts)} parts, {len(result.joints)} joints")

write_partition(model, result, 'output_dir/', 'my_model.bdf')
```
