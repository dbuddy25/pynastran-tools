# Nastran Include File Renumbering Tool

## Problem

Nastran models often split data across multiple INCLUDE files (mesh, properties, loads, contact, etc.), each with its own ID numbering range. When reorganizing or merging models, engineers need to renumber entities per-file into new ranges while preserving all cross-references — boundary conditions, loads, contact definitions, etc. pyNastran's built-in `bdf_renumber()` consolidates everything into a single file, doesn't support per-include ranges, and doesn't handle contact cards. This tool fills that gap.

## What It Does

A standalone Python GUI (CustomTkinter + tksheet) that:

1. Reads a .dat/.bdf file and discovers all INCLUDE files (including nested, with or without quotes)
2. Catalogs entity types, counts, and current ID ranges per file
3. Presents an editable spreadsheet table (tksheet) in **Simple** or **Advanced** mode
4. Validates ranges (sufficient capacity, no overlaps, positive IDs)
5. Renumbers all cards — including contact, loads, BCs, case control — and writes per-file output to a new directory
6. Includes a **fallback writer** that guarantees no cards are silently dropped, even for card types not in CARD_ORDER

## Requirements

- Python 3.9+
- `pip install pyNastran` (>= 1.3)
- `pip install customtkinter tksheet`
- tkinter (built-in with Python)

## Usage

```
python nastran_tools.py                  # unified app (Pre-Processing > Renumber)
python preprocessing/renumber_includes.py  # standalone
```

### Workflow

1. **Browse** -> select main .dat/.bdf
2. **Scan** -> tool parses includes, populates entity table
3. Choose **Simple** or **Advanced** mode:
   - **Simple**: one start/end range per file; entity types auto-allocated equal sub-ranges
   - **Advanced**: one start/end range per entity type per file (full control)
4. Fill in **New Start** / **New End** in the sheet cells
5. **Validate** -> checks ranges, reports errors in log
6. Select **Output Dir**
7. **Apply Renumbering** -> renumbers all files, writes to output dir, runs post-validation
8. Optionally: **Save Config** exports ranges to JSON, **Load Config** restores them

### Simple Mode Auto-Allocation

When using Simple mode, the file's range is divided into equal-sized blocks in canonical order:
nid, eid, pid, mid, cid, spc_id, mpc_id, load_id, contact_id, set_id, method_id, table_id.
Only entity types present in the file get a block. The last block gets any remainder.

Example: range 100000-199999, 4 entity types -> nid: 100000-124999, eid: 125000-149999, pid: 150000-174999, mid: 175000-199999.

## Supported Card Types (109 total across 12 entity types)

| Entity Type | Cards |
|---|---|
| Node ID | GRID, SPOINT |
| Element ID | CQUAD4, CTRIA3, CHEXA, CPENTA, CTETRA, CBAR, CBEAM, CROD, CONROD, CBUSH, CELAS1/2/3/4, CDAMP1/2/3/4, CGAP, CQUAD8, CTRIA6, CQUADR, CTRIAR, CSHEAR, PLOTEL, CWELD, CFAST, CVISC, CHBDYG, CHBDYE, RBE2, RBE3, RBAR, CONM1, CONM2, CMASS1/2/3/4 |
| Property ID | PSHELL, PCOMP, PCOMPG, PCOMPLS, PSOLID, PLSOLID, PBAR, PBARL, PBEAM, PBEAML, PROD, PBUSH, PBUSHT, PELAS, PDAMP, PGAP, PSHEAR, PWELD, PFAST, PVISC |
| Material ID | MAT1, MAT2, MAT8, MAT9, MAT10 |
| Coord ID | CORD2R, CORD2C, CORD2S, CORD1R, CORD1C, CORD1S |
| SPC ID | SPC, SPC1, SPCADD |
| MPC ID | MPC, MPCADD |
| Load ID | FORCE, MOMENT, PLOAD4, GRAV, LOAD, TEMP, TEMPD, RFORCE, RLOAD1/2, TLOAD1/2, DAREA, DLOAD, PLOAD, PLOAD2 |
| Contact ID | BSURF, BSURFS, BCTSET, BCTADD, BCONP, BCBODY, BCTPARA, BCTPARM, BLSEG, BFRIC |
| Set ID | SET1, SET3 |
| Method ID | EIGRL, EIGR |
| Table ID | TABLED1, TABLEM1 |

**Fallback safety net**: Any card type that pyNastran parses but isn't listed in CARD_ORDER will still be written via the fallback writer, with a diagnostic warning in the log.

## Cross-References Preserved

All inter-card references are updated: element->node, element->property, property->material, load->node, load->element, PLOAD4->CID, CBUSH->CID, shell->theta_mcid (including CQUADR/CTRIAR), CBAR/CBEAM->g0, BSURF->element, BCTSET->contact, SPCADD->SPC, LOAD combo->load IDs, case control LOAD=/SPC=/MPC=/METHOD=, etc.

## Architecture (7 sections)

| Section | Class | Purpose |
|---|---|---|
| 1 | `IncludeFileParser` (from `bdf_utils`) | Parses raw BDF text, discovers INCLUDEs recursively (quoted & unquoted), catalogs entity IDs per file, handles large-field format |
| 2 | `MappingBuilder` | Builds old->new ID maps from user ranges |
| 3 | `Validator` | Pre-validation (capacity, overlaps, CID 0) and post-validation (count/connectivity) |
| 4 | `CardRenumberer` | Applies ID maps to all card types in pyNastran model, rebuilds dicts |
| 5 | `CaseControlRenumberer` | Updates LOAD=, SPC=, MPC=, METHOD=, DLOAD=, TEMPERATURE= in case control |
| 6 | `OutputWriter` | Writes renumbered cards per-file with fallback writer and diagnostic logging |
| 7 | `RenumberIncludesTool` | CustomTkinter + tksheet GUI: Simple/Advanced mode, file browser, spreadsheet editor, validate, apply, save/load config |

## Source

See `preprocessing/renumber_includes.py` (imports `IncludeFileParser`, `CARD_ENTITY_MAP`, `ENTITY_TYPES`, `ENTITY_LABELS`, and `make_model` from `bdf_utils.py`).
