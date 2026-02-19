# Modal Effective Mass Fraction Report

A Python CLI tool that reads a Nastran SOL 103 (normal modes) OP2 file and prints a modal effective mass fraction table. Optionally compares two OP2 files to track how modes shift between design iterations. Reads **exact MEFFMASS data** written by Nastran (via `MEFFMASS(PLOT)` case control) — no approximation. Works with **NX Nastran**, **MSC Nastran**, and **Optistruct** OP2 files.

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

### Compare two OP2 files

```
python postprocessing/modal_effective_mass.py baseline.op2 --compare updated.op2
```

This prints four tables: File A fractions, File B fractions, comparison by mode number, and comparison by MEFF similarity match.

### Compare with Excel export

```
python postprocessing/modal_effective_mass.py baseline.op2 -c updated.op2 -x compare.xlsx
```

Produces a 4-sheet Excel workbook (see Excel export section below).

## Output

### Single file

A single table with one row per mode. For each of six directions (Tx-Rz) two columns are shown: the per-mode mass fraction and its cumulative sum.

```
  Mode    Freq Tx Frac  Tx Sum Ty Frac  Ty Sum Tz Frac  Tz Sum Rx Frac  Rx Sum Ry Frac  Ry Sum Rz Frac  Rz Sum
----------------------------------------------------------------------------------------------------------------
     1    12.3    0.45    0.45    0.00    0.00    0.00    0.00    0.00    0.00    0.00    0.00    0.00    0.00
     2    18.8    0.00    0.45    0.39    0.39    0.00    0.00    0.00    0.00    0.00    0.00    0.00    0.00
     3    25.1    0.00    0.45    0.00    0.39    0.51    0.51    0.00    0.00    0.00    0.00    0.00    0.00
```

Fraction values are decimals (0.45 = 45%). The Sum columns show the running total — a healthy model has translational sums trending toward 1.0.

### Comparison output

When `--compare` is used, two additional tables are printed after the individual file tables:

**Comparison by Mode Number** — matches modes that share the same mode number between the two files and shows frequency and fraction deltas:

```
  Mode   Freq A   Freq B     D Hz      D %     DTx     DTy     DTz     DRx     DRy     DRz
---------------------------------------------------------------------------------------------
     1     12.3     12.5      0.2     1.63    0.01    0.00    0.00    0.00    0.00    0.00
     2     18.8     19.1      0.3     1.60    0.00   -0.02    0.00    0.00    0.00    0.00
```

**Comparison by MEFF Match** — for each mode in File A, finds the best-matching mode in File B by cosine similarity of the 6-D MEFFMASS fraction vector. Weak matches (similarity < 0.5) are flagged with `*`:

```
Mode A Match B   Sim   Freq A   Freq B     D Hz      D %     DTx     DTy     DTz     DRx     DRy     DRz
-----------------------------------------------------------------------------------------------------------
     1       1 0.998     12.3     12.5      0.2     1.63    0.01    0.00    0.00    0.00    0.00    0.00
     2       2 0.995     18.8     19.1      0.3     1.60    0.00   -0.02    0.00    0.00    0.00    0.00
     5       3 0.421*    42.0     25.5    -16.5   -39.29   -0.10    0.15    0.00    0.00    0.00    0.00
```

## Comparison matching strategies

### By Mode Number

Simple: mode 1 in File A is compared to mode 1 in File B. Only modes present in both files are shown. This works well when the mesh/design change is small and modes haven't reordered.

### By MEFF Match (cosine similarity)

For each mode in File A, the algorithm computes cosine similarity of its 6-D MEFFMASS fraction vector against every mode in File B. The absolute value is used to handle eigenvector sign flips between solver runs. The best match is reported along with the similarity score (0-1).

This is a **greedy** match — multiple File A modes can map to the same File B mode. This is intentional: if a mode disappears or splits, you'll see multiple A-modes pointing to the same B-mode (or weak matches), which is a useful diagnostic.

## Excel export

### Single file (`--xlsx` without `--compare`)

Produces a single-sheet workbook titled "Effective Mass Fractions" with dark blue merged direction headers, medium blue sub-headers, right-aligned numbers, light row borders, and frozen panes.

### Comparison (`--xlsx` with `--compare`)

Produces a 4-sheet workbook:

| Sheet | Contents |
|---|---|
| **File A - MEFFMASS** | Full fraction + cumulative sum table for File A |
| **File B - MEFFMASS** | Full fraction + cumulative sum table for File B |
| **Compare - Mode Number** | Mode-number-matched deltas (freq + fraction) |
| **Compare - MEFF Match** | MEFF-similarity-matched deltas with similarity scores; weak-match rows in red |

## GUI usage

The GUI (`python postprocessing/nastran_tool.py`) provides the same functionality through a tab-based interface:

1. **Open primary OP2**: File -> Open OP2 (Cmd+O / Ctrl+O)
2. **Open comparison OP2**: File -> Open Comparison OP2 (Cmd+Shift+O / Ctrl+Shift+O) — enabled after primary file is loaded
3. **Toggle comparison views**: Radio buttons appear above the table — "By Mode Number" and "By MEFF Match"
4. **Clear comparison**: File -> Clear Comparison — reverts to single-file view
5. **Export to Excel**: Click "Export to Excel..." — produces single-sheet or 4-sheet workbook depending on whether a comparison is active

Opening a new primary OP2 automatically clears any existing comparison.

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
- **Mode swapping**: When the MEFF match assigns a different mode number than the mode-number match, modes have reordered between iterations — common when stiffness changes shift natural frequencies past each other
- **Frequency drift**: Small D% values (< 5%) indicate the design change has minimal effect on that mode; large shifts warrant investigation
- **Weak matches** (similarity < 0.5, flagged with `*` in CLI or red in GUI/Excel): The mode shape character has changed significantly — the mode may have merged with another, split, or been replaced by a new mode
- **Multiple A-modes mapping to same B-mode**: Indicates mode coalescence — two distinct modes in the baseline have merged into one in the updated design

## Source

See `postprocessing/modal_effective_mass.py` and `postprocessing/modules/meff.py`.
