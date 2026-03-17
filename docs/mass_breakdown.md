# Mass Breakdown Tool

Computes element mass distribution from a Nastran BDF file, grouped by property ID or include file. Supports part superelements, DMIG mass matrices (via M2GG case control), and optional OP2 GPWG validation.

## Requirements

```
pip install pyNastran numpy
```

**openpyxl** is only needed for Excel export:

```
pip install openpyxl
```

Python 3.9 or later.

## Usage

### GUI (recommended)

Launch from the unified GUI under Post-Processing:

```
python nastran_tools.py
```

Or run standalone:

```
python postprocessing/modules/mass_breakdown.py
```

### Workflow

1. **Open BDF** — select a BDF file to extract element masses
2. **Select grouping** — "Property ID" or "Include File" from the dropdown
3. **Review** — the table shows mass per group with percentages
4. **Manage Groups** — combine multiple IDs into named groups (optional)
5. **Open OP2** — load an OP2 for GPWG validation (optional)
6. **Export to Excel** — save as a formatted .xlsx workbook

## Grouping Modes

### Property ID (default)

Groups elements by their property ID. Each row shows "PID N" with the summed mass. If a property card has a comment, the comment text is used as the display name.

Special groups:
- **CONROD (no PID)** — CONROD elements have no property card
- **Mass Elements** — CONM2, CMASS1-4, CONM1 lumped mass cards

### Include File

Groups elements by their source BDF include file. Useful for models organized by structural component (e.g., `wing_upper.bdf`, `fuselage.bdf`). Files appear in the order they're encountered in the BDF.

> **Note:** Include file grouping currently maps only residual structure elements. Superelement elements won't appear in this mode.

## Custom Groups (Manage Groups)

Click **Manage Groups** to open the grouping dialog:

- **Create a group** — enter a name, select member IDs from the available list, click Create
- **Delete a group** — select an existing group and click Delete
- **Reorder groups** — use the up/down buttons to control display order
- **Show ungrouped as individual** — when unchecked, unmerged IDs collapse into an "Other" row
- **CSV import/export** — save and reload group definitions

Groups persist for the session. All ID types (PIDs, SE PIDs, DMIG matrices) can be merged together.

### Example

To combine wing properties into a single row:

1. Open Manage Groups
2. Enter name: "Wing Structure"
3. Select PID 10, PID 11, PID 12 from the available list
4. Click Create
5. Click Apply

The table now shows "Wing Structure" with the combined mass of those three PIDs.

## Superelement Support

### Part Superelements (BEGIN SUPER)

When the BDF contains partitioned superelements via `BEGIN SUPER=N`, each SE's elements are prefixed with the SE ID:

| Group | Mass | % |
|---|---|---|
| PID 1 | 50.0 | 25.0 |
| SE10:PID 5 | 30.0 | 15.0 |
| SE10:PID 6 | 20.0 | 10.0 |
| SE20:PID 5 | 40.0 | 20.0 |

You can merge SE and residual PIDs together via Manage Groups.

### DMIG Mass Matrices (M2GG)

External superelements are often represented as DMIG matrices with no physical elements. If the case control deck contains M2GG entries:

```
SUBCASE 1
  M2GG = 1.03*MPART1, 1.06*MPART2, 1.06*MPART3, 1.06*MPART4, 1.06*MPART5
```

The tool:
1. Parses each `scale * matrix_name` term from M2GG
2. Retrieves the DMIG matrix from the BDF
3. Sums diagonal entries at translational DOFs (components 1, 2, 3)
4. Divides by 3 (each node's mass appears on all 3 translational DOFs)
5. Applies the scale factor

Each matrix appears as its own group:

| Group | Mass | % |
|---|---|---|
| PID 1 | 120.0 | 30.0 |
| M2GG: MPART1 (x1.03) | 45.2 | 11.3 |
| M2GG: MPART2 (x1.06) | 38.7 | 9.7 |
| M2GG: MPART3 (x1.06) | 22.1 | 5.5 |
| M2GG: MPART4 (x1.06) | 31.5 | 7.9 |
| M2GG: MPART5 (x1.06) | 19.8 | 5.0 |
| Mass Elements | 122.7 | 30.6 |
| TOTAL | 400.0 | 100.0 |

DMIG groups appear in both Property ID and Include File grouping modes, and can be merged into custom groups alongside regular PIDs.

The status bar shows `[N M2GG]` when DMIG matrices are detected.

## GPWG Validation (Grid Point Weight Generator)

The GPWG check compares the tool's BDF-computed mass total against Nastran's own mass calculation from the OP2 results. This catches:
- Missing mass (elements the tool can't compute mass for)
- Discrepancies between BDF input and assembled model
- Confirmation that DMIG mass extraction is correct

### Setup

Add to your Nastran **bulk data** section:

```
PARAM,GRDPNT,0
```

This tells Nastran to compute and output the Grid Point Weight Generator table at the basic origin. The GPWG contains total mass, center of gravity, and inertia for the fully assembled model — including all superelements and DMIG contributions.

### Using GPWG in the tool

1. Run your Nastran analysis (any SOL) with `PARAM,GRDPNT,0`
2. In the Mass Breakdown tool, click **Open OP2**
3. Select the OP2 results file
4. A validation row appears below the total:

| Group | Mass | % |
|---|---|---|
| ... | ... | ... |
| TOTAL | 400.0 | 100.0 |
| GPWG Total (Δ: +0.123) | 400.123 | |

- **Small Δ** (< 1% of total) — good agreement, rounding differences
- **Large Δ** — mass is unaccounted for; check for elements the tool can't extract mass from, or DMIG matrices not referenced in M2GG
- **Positive Δ** — Nastran sees more mass than the BDF extraction found
- **Negative Δ** — BDF extraction found more than Nastran (unusual, may indicate cross-referencing issues)

For superelement models, the tool sums GPWG entries across all SEs to get a single total.

## Excel Export

Click **Export to Excel** to save a formatted .xlsx workbook with:

- Optional title row (enter in the Title field before exporting)
- BDF filename row
- Column headers: Group, Mass, % of Total
- Data rows with number formatting (3 decimal places for mass, 1 for %)
- Bold total row with dark blue background
- GPWG validation row in grey italic (if OP2 loaded)
- Frozen panes and light blue cell borders

## Source

See `postprocessing/modules/mass_breakdown.py` (GUI module). The unified GUI is `nastran_tools.py`.
