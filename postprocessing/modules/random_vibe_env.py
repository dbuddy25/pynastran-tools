"""Random Vibration Environment Generator.

Generates weight-adjusted random vibration test environments per specification.

Supported specs:
  SMC-S-016 (2014), Appendix B
    - Baseline:  Section 6.3.5.3, Figure 6.3.5-1
    - Reduction: Section B.2.1, Equation B.9 (units > 50 lb / 23 kg)
  GEVS — GSFC-STD-7000B, Table 2.4-3 (components, ELV)
    - Baseline:  Table 2.4-3 (acceptance column; qual is exactly 2x)
    - Reduction: Table 2.4-3 notes (components > 50 lb / 22.7 kg)

Each spec is a _SPECS entry supplying its own breakpoints, test levels,
reduce() and describe() callables, so adding another is a registry entry
plus two functions -- no GUI changes.
"""

import math
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog

import customtkinter as ctk
import numpy as np
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure
import matplotlib.ticker as ticker

from .asd_common import _THEMES, grms_loglog, interp_loglog


_LB_PER_KG = 2.20462
_KG_PER_LB = 0.453592

_COLOR_SPEC    = "#1f77b4"   # blue  — spec level (pre-reduction)
_COLOR_REDUCED = "#d62728"   # red   — weight-reduced level


# ── helpers ──────────────────────────────────────────────────────────────────

def _db_oct_exponent(db_per_oct):
    """Convert dB/octave slope to log-log power-law exponent.

    Derivation: 1 oct = factor-of-2 in freq.  For PSD, dB = 10*log10(ratio).
    So S dB/oct → PSD(f) = PSD0*(f/f0)^n where n = S/(10*log10(2)).
    """
    return db_per_oct / (10.0 * math.log10(2.0))


def _expand_profile(bp_freqs, bp_asd, n_pts=600):
    """Return a dense log-spaced profile from breakpoints for smooth log-log plotting."""
    freqs = np.geomspace(bp_freqs[0], bp_freqs[-1], n_pts)
    asd   = interp_loglog(np.asarray(bp_freqs), np.asarray(bp_asd), freqs)
    return freqs, asd


def _show_popup(parent, title, text):
    try:
        from structures_tools import show_guide
        show_guide(parent, title, text,
                   font=ctk.CTkFont(family="Courier", size=12),
                   width=600, height=520)
    except ImportError:
        pass


# ── SMC-S-016 reduction ───────────────────────────────────────────────────────

def _reduce_smc_b9(spec, weight_lb):
    """Apply SMC-S-016 Eq. B.9 broadband reduction.

    Reference: SMC-S-016 (2014), Section B.2.1, Equation B.9.

    Returns (bp_freqs, bp_asd, details).
    details['reduced'] is False when weight is at or below the threshold.
    """
    bp        = spec["baseline"]           # list of (freq, asd) breakpoints
    if len(bp) != 4:
        raise ValueError(
            f"_reduce_smc_b9 requires exactly 4 baseline breakpoints, got {len(bp)}"
        )
    wa        = spec["weight_adjust"]
    threshold = wa["threshold_lb"]         # 50 lb
    ref_flat  = wa["ref_flat"]             # 0.04 g²/Hz
    ref_w     = wa["ref_weight_lb"]        # 50 lb
    max_db    = wa["max_reduction_db"]     # 6 dB

    bp_freqs = np.array([p[0] for p in bp])
    bp_asd   = np.array([p[1] for p in bp])

    if weight_lb <= threshold:
        return bp_freqs.copy(), bp_asd.copy(), {
            "reduced": False, "weight_lb": weight_lb,
            "weight_kg": weight_lb * _KG_PER_LB,
        }

    # Eq. B.9
    new_flat = ref_flat * (ref_w / weight_lb)

    # Clamp to max reduction (6 dB → min flat = 0.01 g²/Hz)
    min_flat = ref_flat * 10.0 ** (-max_db / 10.0)
    capped   = new_flat < min_flat
    new_flat = max(new_flat, min_flat)
    effective_weight_lb = ref_flat * ref_w / new_flat   # actual W used

    reduction_db = 10.0 * math.log10(ref_flat / new_flat)

    # New low-freq breakpoint: +3 dB/oct ramp from anchor meets new flat
    # Anchor is always bp[0] = (20 Hz, 0.0053 g²/Hz) — fixed per spec
    n_up       = _db_oct_exponent(wa["ramp_up_db_oct"])   # ≈ +0.997
    anchor_f   = bp[0][0]
    anchor_asd = bp[0][1]
    f_break    = anchor_f * (new_flat / anchor_asd) ** (1.0 / n_up)

    # High-freq endpoint: -6 dB/oct from flat_end_freq with new flat level
    n_down      = _db_oct_exponent(wa["ramp_down_db_oct"])  # ≈ -1.993
    flat_end_f  = bp[2][0]   # 800 Hz
    end_f       = bp[3][0]   # 2000 Hz
    new_end_asd = new_flat * (end_f / flat_end_f) ** n_down

    new_bp_freqs = np.array([anchor_f, f_break,   flat_end_f, end_f])
    new_bp_asd   = np.array([anchor_asd, new_flat, new_flat, new_end_asd])

    return new_bp_freqs, new_bp_asd, {
        "reduced":             True,
        "weight_lb":           weight_lb,
        "weight_kg":           weight_lb * _KG_PER_LB,
        "new_flat":            new_flat,
        "f_break":             f_break,
        "new_end_asd":         new_end_asd,
        "reduction_db":        reduction_db,
        "capped":              capped,
        "effective_weight_lb": effective_weight_lb,
        "min_flat":            min_flat,
        "anchor_f":            anchor_f,
        "anchor_asd":          anchor_asd,
        "flat_end_f":          flat_end_f,
        "end_f":               end_f,
    }


# GEVS end-point ASD floor (g²/Hz) held at 20 and 2000 Hz for heavy components,
# per the Table 2.4-3 note. GEVS quotes the bare value 0.01 without naming a test
# level; it is taken here as a QUALIFICATION/protoflight-level requirement, since
# that is the spectrum the table is written around. The registry baseline is the
# acceptance column, so the floor is stored divided down by the qual offset and
# lands back on exactly 0.01 once _recompute() applies the +3 dB level scale.
# If your program reads that note as an acceptance-level floor, drop the divisor.
_GEVS_END_ASD_QUAL = 0.01
# GEVS labels the qual/protoflight column "+3 dB" over acceptance, but the
# published numbers are an exact factor of 2 (0.08→0.16, 0.013→0.026), i.e.
# 3.0103 dB. Using a literal 3.0 here reproduces the table only to ~0.24%,
# which shows up as 14.12 Grms against the published 14.1. Carry the exact
# ratio; the UI still renders it as "+3 dB".
_GEVS_QUAL_DB = 10.0 * math.log10(2.0)


def _reduce_gevs(spec, weight_lb):
    """Apply the GEVS Table 2.4-3 component weight reduction.

    Reference: GSFC-STD-7000B, Table 2.4-3 (Generalized Random Vibration Test
    Levels, Components (ELV)).

    Four regimes, per the notes beneath the table:
        W <= 50 lb        no reduction; baseline as published
        50 < W <= 130 lb  plateau reduced; +/-6 dB/oct slopes MAINTAINED, so the
                          20/2000 Hz endpoints fall with the plateau
        130 < W <= 400 lb plateau reduced; slopes ADJUSTED so the endpoints hold
                          at the 0.01 g²/Hz floor (130 lb is precisely where the
                          6 dB/oct slope lands on that floor, so the two regimes
                          meet continuously)
        W > 400 lb        spectrum held at the 400 lb profile

    Returns (bp_freqs, bp_asd, details) at the registry baseline (acceptance)
    level; _recompute() applies the selected test level's dB offset afterwards.
    """
    bp = spec["baseline"]
    if len(bp) != 4:
        raise ValueError(
            f"_reduce_gevs requires exactly 4 baseline breakpoints, got {len(bp)}"
        )
    wa         = spec["weight_adjust"]
    threshold  = wa["threshold_lb"]        # 50 lb
    ref_flat   = wa["ref_flat"]            # 0.08 g²/Hz (acceptance plateau)
    ref_w      = wa["ref_weight_lb"]       # 50 lb
    slope_w    = wa["slope_hold_lb"]       # 130 lb — slopes maintained below this
    cap_w      = wa["cap_weight_lb"]       # 400 lb — spectrum frozen above this

    bp_freqs = np.array([p[0] for p in bp])
    bp_asd   = np.array([p[1] for p in bp])

    if weight_lb <= threshold:
        return bp_freqs.copy(), bp_asd.copy(), {
            "reduced": False, "weight_lb": weight_lb,
            "weight_kg": weight_lb * _KG_PER_LB,
        }

    # Above 400 lb the spectrum is frozen at the 400 lb profile.
    effective_w = min(weight_lb, cap_w)
    capped      = weight_lb > cap_w

    # Plateau reduction — ASD(50-800 Hz) = 0.08 * (50 / W) at acceptance level
    new_flat     = ref_flat * (ref_w / effective_w)
    reduction_db = 10.0 * math.log10(ref_flat / new_flat)

    f_lo_end, f_flat_start = bp_freqs[0], bp_freqs[1]     # 20, 50 Hz
    f_flat_end, f_hi_end   = bp_freqs[2], bp_freqs[3]     # 800, 2000 Hz

    n_up   = _db_oct_exponent(wa["ramp_up_db_oct"])       # +6 dB/oct -> ~+1.993
    n_down = _db_oct_exponent(wa["ramp_down_db_oct"])     # -6 dB/oct -> ~-1.993

    # The floor is a qualification-level value; express it at the acceptance
    # baseline so the +3 dB level scale puts it back on 0.01 g²/Hz.
    end_floor = _GEVS_END_ASD_QUAL / (10.0 ** (_GEVS_QUAL_DB / 10.0))

    slopes_held = effective_w <= slope_w
    if slopes_held:
        # Slopes maintained at +/-6 dB/oct; endpoints ride down with the plateau.
        new_lo_asd = new_flat * (f_lo_end / f_flat_start) ** n_up
        new_hi_asd = new_flat * (f_hi_end / f_flat_end) ** n_down
    else:
        # Slopes adjusted to hold the floor at 20 and 2000 Hz.
        new_lo_asd = end_floor
        new_hi_asd = end_floor

    new_bp_freqs = np.array([f_lo_end, f_flat_start, f_flat_end, f_hi_end])
    new_bp_asd   = np.array([new_lo_asd, new_flat, new_flat, new_hi_asd])

    # Effective slopes actually flown, for the details pane.
    eff_up_db_oct   = 10.0 * math.log10(new_flat / new_lo_asd) / math.log10(
        f_flat_start / f_lo_end) * math.log10(2.0)
    eff_down_db_oct = 10.0 * math.log10(new_hi_asd / new_flat) / math.log10(
        f_hi_end / f_flat_end) * math.log10(2.0)

    return new_bp_freqs, new_bp_asd, {
        "reduced":            True,
        "weight_lb":          weight_lb,
        "weight_kg":          weight_lb * _KG_PER_LB,
        "effective_weight_lb": effective_w,
        "new_flat":           new_flat,
        "reduction_db":       reduction_db,
        "capped":             capped,
        "slopes_held":        slopes_held,
        "new_lo_asd":         new_lo_asd,
        "new_hi_asd":         new_hi_asd,
        "end_floor":          end_floor,
        "eff_up_db_oct":      eff_up_db_oct,
        "eff_down_db_oct":    eff_down_db_oct,
        "f_lo_end":           f_lo_end,
        "f_flat_start":       f_flat_start,
        "f_flat_end":         f_flat_end,
        "f_hi_end":           f_hi_end,
    }


def _describe_smc_b9(spec, details, weight_raw, unit_str):
    """Weight-adjustment narrative for SMC-S-016. Returns a list of lines."""
    wa       = spec["weight_adjust"]
    w_lb     = details["weight_lb"]
    w_kg     = details["weight_kg"]
    new_flat = details["new_flat"]
    f_break  = details["f_break"]
    red_db   = details["reduction_db"]

    L = []
    L.append(f"  Input:  {weight_raw} {unit_str}  =  {w_lb:.1f} lb  ({w_kg:.1f} kg)")
    L.append("")
    L.append(f"  Eq. B.9:  Reduced flat = 0.04 × (50 / W)")
    if w_lb != details["effective_weight_lb"]:
        L.append(f"                        = 0.04 × (50 / {details['effective_weight_lb']:.0f})")
        L.append(f"                          [W capped at {details['effective_weight_lb']:.0f} lb")
        L.append(f"                           — max 6 dB reduction]")
    else:
        L.append(f"                        = 0.04 × (50 / {w_lb:.1f})")
    L.append(f"                        = {new_flat:.4f}  g²/Hz")
    L.append("")
    L.append(f"  Reduction: {red_db:.1f} dB"
             f"  (max: {wa['max_reduction_db']:.0f} dB)")
    L.append("")
    L.append("ADJUSTED BREAKPOINTS")
    L.append(f"  {details['anchor_f']:.0f} Hz:         "
             f"{details['anchor_asd']:.4f}  g²/Hz  (anchor, unchanged)")
    L.append(f"  {f_break:.1f} Hz:       "
             f"{new_flat:.4f}  g²/Hz  (ramp meets reduced flat)")
    L.append(f"  {details['flat_end_f']:.0f} Hz:        "
             f"{new_flat:.4f}  g²/Hz  (flat end)")
    L.append(f"  {details['end_f']:.0f} Hz:       "
             f"{details['new_end_asd']:.5f} g²/Hz")
    return L


def _describe_gevs(spec, details, weight_raw, unit_str):
    """Weight-adjustment narrative for GEVS Table 2.4-3. Returns a list of lines."""
    wa       = spec["weight_adjust"]
    w_lb     = details["weight_lb"]
    w_kg     = details["weight_kg"]
    new_flat = details["new_flat"]
    red_db   = details["reduction_db"]
    eff_w    = details["effective_weight_lb"]

    L = []
    L.append(f"  Input:  {weight_raw} {unit_str}  =  {w_lb:.1f} lb  ({w_kg:.1f} kg)")
    L.append("")
    L.append(f"  Table 2.4-3:  ASD(50–800 Hz) = 0.08 × (50 / W)   [acceptance]")
    L.append(f"                              = 0.16 × (50 / W)   [qual/protoflight]")
    if details["capped"]:
        L.append(f"                W held at {eff_w:.0f} lb — spectrum is frozen")
        L.append(f"                at the {eff_w:.0f} lb profile above that weight")
    L.append(f"                              = {new_flat:.4f}  g²/Hz  (acceptance)")
    L.append("")
    L.append(f"  dB reduction = 10 log(W / 50) = {red_db:.1f} dB")
    L.append("")

    if details["slopes_held"]:
        L.append(f"  Slopes MAINTAINED at ±6 dB/oct")
        L.append(f"  (W ≤ {wa['slope_hold_lb']:.0f} lb — endpoints fall with the plateau)")
    else:
        L.append(f"  Slopes ADJUSTED to hold {_GEVS_END_ASD_QUAL:.2f} g²/Hz at 20 and 2000 Hz")
        L.append(f"  (W > {wa['slope_hold_lb']:.0f} lb)")
        L.append(f"  Effective: {details['eff_up_db_oct']:+.2f} dB/oct  /  "
                 f"{details['eff_down_db_oct']:+.2f} dB/oct")
    L.append("")
    L.append("ADJUSTED BREAKPOINTS  (acceptance level)")
    L.append(f"  {details['f_lo_end']:.0f} Hz:        {details['new_lo_asd']:.5f} g²/Hz")
    L.append(f"  {details['f_flat_start']:.0f} Hz:        {new_flat:.4f}  g²/Hz  (plateau begins)")
    L.append(f"  {details['f_flat_end']:.0f} Hz:       {new_flat:.4f}  g²/Hz  (plateau ends)")
    L.append(f"  {details['f_hi_end']:.0f} Hz:      {details['new_hi_asd']:.5f} g²/Hz")
    return L


# ── spec registry ─────────────────────────────────────────────────────────────

_SPECS = {
    "SMC-S-016": {
        "label":         "SMC-S-016 — Unit Random Vibration (Acceptance)",
        "source":        "SMC-S-016 (2014), Appendix B",
        "baseline_ref":  "Figure 6.3.5-1",
        "reduction_ref": "Section B.2.1, Equation B.9",
        # Breakpoints: (freq_hz, asd_g2hz)
        # Segments: +3 dB/oct ramp  →  flat  →  -6 dB/oct rolloff
        "baseline": [
            (20.0,   0.0053),    # anchor — low-freq starting point
            (150.0,  0.04),      # flat level begins
            (800.0,  0.04),      # flat level ends
            (2000.0, 0.00644),   # rolloff endpoint
        ],
        "test_levels": [
            # (label, dB_offset_above_acceptance, duration)
            ("Acceptance",         0.0, "1 min/axis"),
            ("Protoqualification", 3.0, "2 min/axis"),
            ("Qualification",      6.0, "3 min/axis"),
        ],
        "weight_adjust": {
            "threshold_lb":    50.0,
            "threshold_kg":    23.0,
            "ref_flat":        0.04,    # baseline flat level (g²/Hz)
            "ref_weight_lb":   50.0,    # W_ref in Eq. B.9
            "max_reduction_db": 6.0,
            "ramp_up_db_oct":   3.0,
            "ramp_down_db_oct": -6.0,
        },
        # reduce(spec, weight_lb) → (bp_freqs, bp_asd, details)
        # details["reduced"] is False when weight ≤ threshold (arrays = baseline copy).
        "reduce": _reduce_smc_b9,
        # describe(spec, details, weight_raw, unit_str) → list of detail lines
        "describe": _describe_smc_b9,
        # labels for the 3 segments between the 4 baseline breakpoints
        "segment_labels": ["+3 dB/oct", "flat", "-6 dB/oct"],
    },
    "GEVS": {
        "label":         "GEVS — Generalized Random Vibration, Components (ELV)",
        "source":        "GSFC-STD-7000B, Table 2.4-3",
        "baseline_ref":  "Table 2.4-3, acceptance column",
        "reduction_ref": "Table 2.4-3 notes",
        # Baseline is the ACCEPTANCE column; the qualification/protoflight
        # column is exactly +3 dB above it (0.08→0.16, 0.013→0.026), so it
        # falls out of the test_levels offset rather than being duplicated.
        # Segments: +6 dB/oct ramp → plateau → -6 dB/oct rolloff
        "baseline": [
            (20.0,   0.013),     # low endpoint
            (50.0,   0.08),      # plateau begins
            (800.0,  0.08),      # plateau ends
            (2000.0, 0.013),     # high endpoint
        ],
        "test_levels": [
            # (label, dB_offset_above_acceptance, duration) — durations per Table 2.4-1
            ("Acceptance",              0.0,           "1 min/axis"),
            ("Protoflight Qual",        _GEVS_QUAL_DB, "1 min/axis"),
            ("Prototype Qual",          _GEVS_QUAL_DB, "2 min/axis"),
        ],
        "weight_adjust": {
            "threshold_lb":     50.0,    # reduction applies above this
            "threshold_kg":     22.7,
            "ref_flat":         0.08,    # acceptance plateau (g²/Hz)
            "ref_weight_lb":    50.0,
            "slope_hold_lb":    130.0,   # ±6 dB/oct maintained up to here (59 kg)
            "cap_weight_lb":    400.0,   # spectrum frozen above here (182 kg)
            "max_reduction_db": 10.0 * math.log10(400.0 / 50.0),   # ≈ 9.03 dB at the cap
            "ramp_up_db_oct":   6.0,
            "ramp_down_db_oct": -6.0,
        },
        "reduce":   _reduce_gevs,
        "describe": _describe_gevs,
        "segment_labels": ["+6 dB/oct", "plateau", "-6 dB/oct"],
    },
}


# ── help text ─────────────────────────────────────────────────────────────────

_HELP_SMC = """\
RANDOM VIBRATION ENVIRONMENT GENERATOR

Generates weight-adjusted random vibration (RV) test environments
from standard specifications.

WORKFLOW
  1. Select a specification from the Spec dropdown.
  2. Select a test level (Acceptance / Protoqual / Qual).
     The baseline spectrum appears immediately on the plot.
  3. Enter the component weight (lb or kg) and press Enter.
     The weight-adjusted (reduced) spectrum overlays the original.
  4. Export saves the reduced profile as a two-column text file.

SMC-S-016 REDUCTION (Section B.2.1)
  Applies to units weighing more than 50 lb (23 kg).

  Equation B.9:
    Reduced flat level (g²/Hz) = 0.04 × (50 / W)

    where W = unit weight in pounds.

  Constraints:
    • Max reduction: 6 dB  →  min flat = 0.01 g²/Hz (W = 200 lb)
    • Anchor (20 Hz, 0.0053 g²/Hz) is always unchanged
    • +3 dB/oct ramp meets new flat at a lower breakpoint frequency
    • -6 dB/oct rolloff starts from the new flat at 800 Hz

BASELINE SPECTRUM (Figure 6.3.5-1)
  20 Hz:        0.0053 g²/Hz  (anchor)
  20–150 Hz:    +3 dB/oct
  150–800 Hz:   0.04   g²/Hz  (flat)
  800–2000 Hz:  -6 dB/oct
  2000 Hz:      0.00644 g²/Hz

VALIDATION (from Figure B.2.2-1)
  Weight  Flat (g²/Hz)  Low break (Hz)  GRMS
  50 lb   0.040         150             6.90 g
  100 lb  0.020          ~75            4.87 g
  200 lb  0.010          ~38            3.52 g

SLOPE MATH (log-log space)
  dB/octave slopes are power laws: PSD(f) = PSD0 × (f/f0)^n
    n = S_dB_oct / (10 × log10(2))
  Breakpoint frequency from inverse power law:
    f_break = f_anchor × (new_flat / anchor_asd)^(1/n)
"""


_HELP_GEVS = """\
RANDOM VIBRATION ENVIRONMENT GENERATOR

GEVS component random vibration environments, weight-adjusted.

WORKFLOW
  1. Select GEVS from the Spec dropdown.
  2. Select a test level. The baseline spectrum appears immediately.
  3. Enter the component weight (lb or kg) and press Enter.
     The weight-adjusted spectrum overlays the original.
  4. Export saves the reduced profile as a two-column text file.

BASELINE SPECTRUM (Table 2.4-3, components 22.7 kg / 50 lb or less)
                        Qualification    Acceptance
  20 Hz:                  0.026            0.013    g²/Hz
  20-50 Hz:              +6 dB/oct        +6 dB/oct
  50-800 Hz:              0.16             0.08     g²/Hz  (plateau)
  800-2000 Hz:           -6 dB/oct        -6 dB/oct
  2000 Hz:                0.026            0.013    g²/Hz
  Overall:               14.1 Grms        10.0 Grms

  The acceptance column is stored as the baseline; the qualification
  column is exactly 2x it. GEVS labels that step "+3 dB" but the true
  ratio is 3.0103 dB, which is what the tool carries -- a literal 3.0
  would reproduce the table only to ~0.24%.

TEST LEVELS (Table 2.4-1)
  Acceptance          Limit level       1 min/axis
  Protoflight Qual    Limit + 3 dB      1 min/axis
  Prototype Qual      Limit + 3 dB      2 min/axis

WEIGHT REDUCTION (Table 2.4-3 notes)
  Applies to components weighing more than 22.7 kg (50 lb).

    dB reduction      = 10 log(W/22.7 kg)  =  10 log(W/50 lb)
    ASD(50-800 Hz)    = 0.16 x (50/W)      qualification/protoflight
                      = 0.08 x (50/W)      acceptance

  Slope handling has two regimes:
    W <= 130 lb (59 kg)   slopes MAINTAINED at +/-6 dB/oct, so the
                          20 and 2000 Hz endpoints fall with the plateau
    W >  130 lb           slopes ADJUSTED to hold 0.01 g²/Hz at 20 and
                          2000 Hz
    W >  400 lb (182 kg)  spectrum frozen at the 400 lb profile

  130 lb is where a +/-6 dB/oct slope naturally lands on 0.01 g²/Hz, so
  the two regimes very nearly meet. GEVS rounds 59 kg to 130 lb while
  the exact crossover is ~128.5 lb, leaving a ~1% step at the seam.
  That discontinuity is in the specification, not in this tool.

  NOTE ON THE 0.01 g²/Hz FLOOR
  GEVS quotes the floor without naming a test level. It is applied here
  as a QUALIFICATION-level requirement, since Table 2.4-3 is written
  around that spectrum. If your program reads it as an acceptance-level
  floor, change _GEVS_END_ASD_QUAL / the divisor in _reduce_gevs.

VALIDATION (computed against the published table)
  Level        Published   Computed
  Qualification  14.1 g     14.14 g
  Acceptance     10.0 g     10.00 g

SLOPE MATH (log-log space)
  dB/octave slopes are power laws: PSD(f) = PSD0 x (f/f0)^n
    n = S_dB_oct / (10 x log10(2))
"""


_HELP_TEXTS = {
    "SMC-S-016": _HELP_SMC,
    "GEVS":      _HELP_GEVS,
}


# ── module class ──────────────────────────────────────────────────────────────

class RandomVibeEnvModule:
    name = "RV Environment"

    def __init__(self, parent):
        self.frame = ctk.CTkFrame(parent)
        self._theme = "light"

        self._spec_key      = list(_SPECS.keys())[0]
        self._level_idx     = 0
        self._weight_lb     = None   # None = not entered yet
        self._weight_pending = False

        self._build_ui()
        self._on_spec_change(self._spec_key)

    # ── build UI ──────────────────────────────────────────────────────────────

    def _build_ui(self):
        # ── toolbar ──
        toolbar = ctk.CTkFrame(self.frame, height=44)
        toolbar.pack(fill=tk.X, side=tk.TOP)
        toolbar.pack_propagate(False)

        spec_keys = list(_SPECS.keys())

        ctk.CTkLabel(toolbar, text="Spec:").pack(side=tk.LEFT, padx=(10, 2))
        self._spec_var = tk.StringVar(value=spec_keys[0])
        self._spec_menu = ctk.CTkOptionMenu(
            toolbar, variable=self._spec_var,
            values=spec_keys, width=200,
            command=self._on_spec_change,
        )
        self._spec_menu.pack(side=tk.LEFT, padx=(0, 12))

        ctk.CTkLabel(toolbar, text="Level:").pack(side=tk.LEFT, padx=(0, 2))
        self._level_var = tk.StringVar()
        self._level_menu = ctk.CTkOptionMenu(
            toolbar, variable=self._level_var,
            values=[""], width=180,
            command=self._on_level_change,
        )
        self._level_menu.pack(side=tk.LEFT, padx=(0, 12))

        ctk.CTkLabel(toolbar, text="Weight:").pack(side=tk.LEFT, padx=(0, 2))
        self._weight_var = tk.StringVar()
        self._weight_entry = ctk.CTkEntry(
            toolbar, textvariable=self._weight_var, width=80,
            placeholder_text="e.g. 100",
        )
        self._weight_entry.pack(side=tk.LEFT, padx=(0, 4))
        self._weight_entry.bind("<Return>",   lambda _e: self._on_weight_submit())
        self._weight_entry.bind("<FocusOut>", lambda _e: self._on_weight_submit())

        self._unit_var = tk.StringVar(value="lb")
        ctk.CTkOptionMenu(
            toolbar, variable=self._unit_var,
            values=["lb", "kg"], width=60,
            command=lambda _: self._on_weight_submit(),
        ).pack(side=tk.LEFT, padx=(0, 14))

        ctk.CTkButton(
            toolbar, text="Export", width=70,
            command=self._export,
        ).pack(side=tk.LEFT, padx=(0, 6))

        ctk.CTkButton(
            toolbar, text="Copy Figure", width=100,
            command=self._copy_figure,
        ).pack(side=tk.LEFT, padx=(0, 6))

        ctk.CTkButton(
            toolbar, text="?", width=30,
            command=self._show_help,
        ).pack(side=tk.LEFT, padx=(0, 6))

        ctk.CTkButton(
            toolbar, text="Light/Dark", width=90,
            command=self._toggle_theme,
        ).pack(side=tk.LEFT, padx=(0, 6))

        # ── body ──
        body = ctk.CTkFrame(self.frame)
        body.pack(fill=tk.BOTH, expand=True)

        # Left: calculation details
        left = ctk.CTkFrame(body, width=450)
        left.pack(side=tk.LEFT, fill=tk.Y, padx=(4, 0), pady=4)
        left.pack_propagate(False)

        ctk.CTkLabel(
            left, text="Calculation Details",
            font=ctk.CTkFont(size=13, weight="bold"),
        ).pack(anchor=tk.W, padx=10, pady=(8, 4))

        self._details_box = ctk.CTkTextbox(
            left, wrap="none", state="disabled",
            font=ctk.CTkFont(family="Courier", size=12),
        )
        self._details_box.pack(fill=tk.BOTH, expand=True, padx=6, pady=(0, 6))

        # Right: matplotlib plot
        right = ctk.CTkFrame(body)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=4, pady=4)

        self._fig = Figure(figsize=(6, 4), dpi=100)
        self._ax  = self._fig.add_subplot(111)

        self._canvas = FigureCanvasTkAgg(self._fig, master=right)
        self._canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        tb_frame = tk.Frame(right)
        tb_frame.pack(fill=tk.X)
        NavigationToolbar2Tk(self._canvas, tb_frame)

        self._refresh_level_menu()

    def _refresh_level_menu(self):
        spec   = _SPECS[self._spec_key]
        labels = [lvl[0] for lvl in spec["test_levels"]]
        self._level_var.set(labels[0])
        self._level_menu.configure(values=labels)
        self._level_idx = 0

    # ── events ────────────────────────────────────────────────────────────────

    def _on_spec_change(self, key):
        self._spec_key  = key
        self._weight_lb = None
        self._weight_var.set("")
        self._refresh_level_menu()
        self._recompute()

    def _on_level_change(self, label):
        spec = _SPECS[self._spec_key]
        for i, (lname, _db, _dur) in enumerate(spec["test_levels"]):
            if lname == label:
                self._level_idx = i
                break
        self._recompute()

    def _on_weight_submit(self):
        # <Return> triggers <FocusOut> immediately after — debounce to fire once.
        if self._weight_pending:
            return
        self._weight_pending = True
        self.frame.after(10, self._process_weight_submit)

    def _process_weight_submit(self):
        self._weight_pending = False
        raw = self._weight_var.get().strip()
        if not raw:
            if self._weight_lb is not None:
                self._weight_lb = None
                self._recompute()
            return
        try:
            val = float(raw)
        except ValueError:
            return
        if not math.isfinite(val) or val <= 0:
            return
        weight_lb = val * _LB_PER_KG if self._unit_var.get() == "kg" else val
        if weight_lb != self._weight_lb:
            self._weight_lb = weight_lb
            self._recompute()

    # ── compute ───────────────────────────────────────────────────────────────

    def _recompute(self):
        spec = _SPECS[self._spec_key]
        bp   = spec["baseline"]

        bp_freqs = np.array([p[0] for p in bp])
        bp_asd   = np.array([p[1] for p in bp])

        _lname, db_offset, _dur = spec["test_levels"][self._level_idx]
        level_scale = 10.0 ** (db_offset / 10.0)

        spec_asd = bp_asd * level_scale

        # Reduced profile — only set when reduction actually applies
        red_freqs = red_asd = red_freqs_base = red_asd_base = details = None
        if self._weight_lb is not None:
            rf, ra, details = spec["reduce"](spec, self._weight_lb)
            if details.get("reduced"):
                red_freqs_base = rf
                red_asd_base   = ra
                red_freqs = rf
                red_asd   = ra * level_scale

        # Compute GRMS once and share between plot and details
        grms_spec = math.sqrt(max(grms_loglog(bp_freqs, spec_asd), 0.0))
        grms_red  = (math.sqrt(max(grms_loglog(red_freqs, red_asd), 0.0))
                     if red_freqs is not None else None)

        self._refresh_plot(bp_freqs, spec_asd, red_freqs, red_asd,
                           grms_spec, grms_red)
        self._refresh_details(spec, bp, bp_freqs, bp_asd, spec_asd,
                              red_freqs_base, red_asd_base, details,
                              db_offset, grms_spec, grms_red)

    # ── plot ──────────────────────────────────────────────────────────────────

    def _refresh_plot(self, spec_freqs, spec_asd, red_freqs, red_asd,
                      grms_spec, grms_red):
        t   = _THEMES[self._theme]
        ax  = self._ax
        fig = self._fig

        fig.patch.set_facecolor(t["fig_bg"])
        ax.clear()
        ax.set_facecolor(t["plot_bg"])

        # Spec level curve (pre-reduction)
        sf_d, sa_d = _expand_profile(spec_freqs, spec_asd)
        lname      = _SPECS[self._spec_key]["test_levels"][self._level_idx][0]
        ax.loglog(sf_d, sa_d, color=_COLOR_SPEC, lw=2,
                  label=f"{lname}  —  {grms_spec:.2f} Grms")

        # Reduced curve (post-reduction)
        if red_freqs is not None:
            rf_d, ra_d = _expand_profile(red_freqs, red_asd)
            raw      = self._weight_var.get().strip()
            unit_str = self._unit_var.get()
            ax.loglog(rf_d, ra_d, color=_COLOR_REDUCED, lw=2,
                      label=f"Reduced ({raw} {unit_str})  —  {grms_red:.2f} Grms")

        ax.grid(True, which='both',  color=t["grid"], linestyle='-',  linewidth=0.5)
        ax.grid(True, which='minor', color=t["grid"], linestyle=':', linewidth=0.3)

        ax.set_xlabel("Frequency (Hz)", color=t["text"])
        ax.set_ylabel("ASD (g²/Hz)",    color=t["text"])
        ax.set_title(
            f"{self._spec_key}  —  {_SPECS[self._spec_key]['baseline_ref']}",
            color=t["text"],
        )
        ax.tick_params(colors=t["text"])
        ax.xaxis.set_major_formatter(ticker.FuncFormatter(lambda x, _: f"{x:.0f}"))
        ax.xaxis.set_minor_formatter(ticker.NullFormatter())
        ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, _: f"{x:g}"))
        ax.yaxis.set_minor_formatter(ticker.NullFormatter())

        # Y limits: snap to nearest decade below min and above max
        all_asd = list(spec_asd)
        if red_asd is not None:
            all_asd.extend(red_asd)
        pos = [v for v in all_asd if v > 0]
        if pos:
            y_lo = 10 ** math.floor(math.log10(min(pos)))
            y_hi = 10 ** math.ceil(math.log10(max(pos)))
            if y_lo < y_hi:
                ax.set_ylim(y_lo, y_hi)

        for spine in ax.spines.values():
            spine.set_edgecolor(t["spine"])

        ax.legend(facecolor=t["legend_bg"], edgecolor=t["spine"],
                  labelcolor=t["text"], fontsize=10)

        self._canvas.draw_idle()

    # ── details text ──────────────────────────────────────────────────────────

    def _refresh_details(self, spec, bp, bp_freqs, bp_asd, spec_asd,
                         red_freqs_base, red_asd_base, details,
                         db_offset, grms_spec, grms_red):
        lname, _db, duration = spec["test_levels"][self._level_idx]
        wa = spec["weight_adjust"]

        L = []   # lines

        L.append(f"{self._spec_key}  —  Weight-Adjusted RV Environment")
        L.append("=" * 52)
        L.append(f"Source:  {spec['source']}")
        L.append("")

        # ── selected level
        db_tag = f"+{db_offset:.0f} dB" if db_offset > 0 else "+0 dB (acceptance level)"
        L.append(f"SELECTED TEST LEVEL")
        L.append(f"  {lname}  ({db_tag},  {duration})")
        L.append("")

        # ── baseline
        grms_base = math.sqrt(max(grms_loglog(bp_freqs, bp_asd), 0.0))

        L.append(f"BASELINE SPECTRUM  ({spec['baseline_ref']})")
        seg_labels = spec.get("segment_labels", ["ramp", "flat", "rolloff"])
        point_tags = ["(anchor)", "(flat begins)", "(flat ends)", ""]
        for i, (f_hz, asd_v) in enumerate(bp):
            tag = point_tags[i] if i < len(point_tags) else ""
            L.append(f"  {f_hz:>6.0f} Hz:  {asd_v:>9.5f}  g²/Hz  {tag}".rstrip())
            if i < len(bp) - 1 and i < len(seg_labels):
                L.append(f"  {f_hz:.0f}–{bp[i + 1][0]:.0f} Hz:  {seg_labels[i]}")
        L.append(f"  Baseline GRMS:   {grms_base:.2f} g")
        if db_offset != 0.0:
            L.append(f"  {lname} GRMS: {grms_spec:.2f} g")
        L.append("")

        # ── weight adjustment
        L.append(f"WEIGHT ADJUSTMENT  ({spec['reduction_ref']})")
        L.append(f"  Applies when W > {wa['threshold_lb']:.0f} lb"
                 f"  ({wa['threshold_kg']:.0f} kg)")
        L.append("")

        if details is None:
            L.append("  (enter weight above)")

        elif not details["reduced"]:
            w_lb = details["weight_lb"]
            w_kg = details["weight_kg"]
            L.append(f"  Weight: {w_lb:.1f} lb  ({w_kg:.1f} kg)")
            L.append(f"  W ≤ {wa['threshold_lb']:.0f} lb  —  no reduction applied")

        else:
            L.extend(spec["describe"](
                spec, details,
                self._weight_var.get().strip(),
                self._unit_var.get()))
            L.append("")

            # GRMS table — use the unscaled breakpoint arrays returned by reduce()
            base_area = grms_loglog(red_freqs_base, red_asd_base)

            L.append(f"TEST LEVELS  (on reduced profile)")
            L.append(f"  {'Level':<22}  {'GRMS':>6}  Duration")
            L.append("  " + "─" * 44)
            for lvl_name, lvl_db, lvl_dur in spec["test_levels"]:
                lvl_scale = 10.0 ** (lvl_db / 10.0)
                lvl_grms  = math.sqrt(max(base_area * lvl_scale, 0.0))
                marker    = "  ◀" if lvl_name == lname else ""
                L.append(f"  {lvl_name:<22}  {lvl_grms:>5.2f} g  {lvl_dur}{marker}")

        text = "\n".join(L)
        self._details_box.configure(state="normal")
        self._details_box.delete("1.0", "end")
        self._details_box.insert("1.0", text)
        self._details_box.configure(state="disabled")

    # ── export ────────────────────────────────────────────────────────────────

    def _export(self):
        if self._weight_lb is None:
            messagebox.showinfo(
                "Export", "Enter a weight first — export produces the reduced profile.")
            return

        spec = _SPECS[self._spec_key]
        rf, ra, details = spec["reduce"](spec, self._weight_lb)

        _lname, db_offset, _dur = spec["test_levels"][self._level_idx]
        ra_scaled = ra * 10.0 ** (db_offset / 10.0)

        raw      = self._weight_var.get().strip()
        unit_str = self._unit_var.get()
        default_name = f"{self._spec_key} {_lname} {raw}{unit_str}"

        name = simpledialog.askstring(
            "Export Name",
            "Environment name (first line of file):",
            initialvalue=default_name,
            parent=self.frame.winfo_toplevel(),
        )
        if name is None:
            return

        path = filedialog.asksaveasfilename(
            title="Export Reduced ASD Profile",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
        )
        if not path:
            return

        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(f"{name}\n")
                for freq, asd_val in zip(rf, ra_scaled):
                    f.write(f"{freq:.2f}  {asd_val:.6g}\n")
        except OSError as exc:
            messagebox.showerror("Export Error", str(exc))
            return

        messagebox.showinfo("Exported", f"Saved to:\n{path}")

    # ── misc ──────────────────────────────────────────────────────────────────

    def _copy_figure(self):
        import io, os, tempfile, subprocess
        buf = io.BytesIO()
        try:
            self._fig.savefig(buf, format='png', dpi=200, bbox_inches='tight',
                              facecolor=self._fig.get_facecolor())
        except Exception as exc:
            messagebox.showerror("Copy Error", f"Could not render figure:\n{exc}")
            return
        buf.seek(0)
        tmp = None
        try:
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as f:
                f.write(buf.getvalue())
                tmp = f.name
            if os.name == 'nt':
                ps = (
                    'Add-Type -Assembly System.Windows.Forms;'
                    '[Windows.Forms.Clipboard]::SetImage('
                    f'[System.Drawing.Image]::FromFile("{tmp}"))'
                )
                subprocess.run(['powershell', '-Command', ps], check=True)
            elif os.uname().sysname == 'Darwin':
                subprocess.run(
                    ['osascript', '-e',
                     f'set the clipboard to '
                     f'(read (POSIX file "{tmp}") as «class PNGf»)'],
                    check=True)
            else:
                subprocess.run(
                    ['xclip', '-selection', 'clipboard',
                     '-t', 'image/png', '-i', tmp],
                    check=True)
        except Exception as exc:
            messagebox.showerror("Copy Error", str(exc))
        finally:
            if tmp:
                try:
                    os.unlink(tmp)
                except Exception:
                    pass

    def _toggle_theme(self):
        self._theme = "light" if self._theme == "dark" else "dark"
        self._recompute()

    def _show_help(self):
        _show_popup(self.frame.winfo_toplevel(),
                    f"RV Environment — {self._spec_key} Help",
                    _HELP_TEXTS.get(self._spec_key, _HELP_SMC))
