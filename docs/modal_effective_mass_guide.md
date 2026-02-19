# Modal Effective Mass Fraction Report

A Python CLI tool that reads a Nastran SOL 103 (normal modes) OP2 file and prints a modal effective mass fraction table. Reads **exact MEFFMASS data** written by Nastran (via `MEFFMASS(PLOT)` case control) — no approximation. Works with **NX Nastran**, **MSC Nastran**, and **Optistruct** OP2 files.

## Requirements

```
pip install pyNastran numpy
```

**openpyxl** is only needed for `--xlsx` Excel export:

```
pip install openpyxl
```

Python 3.9 or later.

## Usage

### Print to console

```
python postprocessing/modal_effective_mass.py model.op2
```

### Export to Excel

```
python postprocessing/modal_effective_mass.py model.op2 --xlsx output.xlsx
```

## Output

A single table with one row per mode. For each of six directions (Tx-Rz) two columns are shown: the per-mode mass fraction and its cumulative sum.

```
  Mode    Freq Tx Frac  Tx Sum Ty Frac  Ty Sum Tz Frac  Tz Sum Rx Frac  Rx Sum Ry Frac  Ry Sum Rz Frac  Rz Sum
----------------------------------------------------------------------------------------------------------------
     1    12.3    0.45    0.45    0.00    0.00    0.00    0.00    0.00    0.00    0.00    0.00    0.00    0.00
     2    18.8    0.00    0.45    0.39    0.39    0.00    0.00    0.00    0.00    0.00    0.00    0.00    0.00
     3    25.1    0.00    0.45    0.00    0.39    0.51    0.51    0.00    0.00    0.00    0.00    0.00    0.00
```

Fraction values are decimals (0.45 = 45%). The Sum columns show the running total — a healthy model has translational sums trending toward 1.0.

## Excel export

Produces a single-sheet workbook titled "Effective Mass Fractions" with dark blue merged direction headers, medium blue sub-headers, right-aligned numbers, light row borders, and frozen panes. Fraction cells at or above the threshold are bolded.

## GUI usage

The MEFF Viewer is available in the unified GUI (`python nastran_tools.py`, under Post-Processing) or as a standalone module. The toolbar provides:

1. **Open OP2...** — load an OP2 file
2. **Threshold** — highlight fraction cells at or above this value
3. **Export to Excel...** — produces a single-sheet workbook

## Nastran model requirements

Your SOL 103 run must include in case control:

```
MEFFMASS(PLOT) = ALL
```

If MEFFMASS data is missing from the OP2, the script prints an error with this instruction.

## Engineering interpretation tips

- **Cumulative Tx, Ty, Tz sums** each trending toward 1.0 as mode count increases indicates good mass participation
- **A few modes dominate** each translational direction — these are your primary structural modes
- **If sums are far below 1.0**, you may need more modes in the analysis

## Source

See `postprocessing/modal_effective_mass.py` (CLI) and `postprocessing/modules/meff.py` (GUI module + shared logic). The unified GUI is `nastran_tools.py`.
