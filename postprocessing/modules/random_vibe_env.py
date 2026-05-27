"""Random Vibration Environment Generator.

Generates weight-adjusted random vibration test environments per specification.

First supported spec: SMC-S-016 (2014), Appendix B.
  - Baseline:  Section 6.3.5.3, Figure 6.3.5-1
  - Reduction: Section B.2.1, Equation B.9 (units > 50 lb / 23 kg)
"""

import math
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk
import numpy as np
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure

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
    },
}


# ── help text ─────────────────────────────────────────────────────────────────

_HELP_TEXT = """\
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


# ── module class ──────────────────────────────────────────────────────────────

class RandomVibeEnvModule:
    name = "RV Environment"

    def __init__(self, parent):
        self.frame = ctk.CTkFrame(parent)
        self._theme = "dark"

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
        left = ctk.CTkFrame(body, width=370)
        left.pack(side=tk.LEFT, fill=tk.Y, padx=(4, 0), pady=4)
        left.pack_propagate(False)

        ctk.CTkLabel(
            left, text="Calculation Details",
            font=ctk.CTkFont(size=13, weight="bold"),
        ).pack(anchor=tk.W, padx=10, pady=(8, 4))

        self._details_box = ctk.CTkTextbox(
            left, wrap="none", state="disabled",
            font=ctk.CTkFont(family="Courier", size=11),
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
        L.append(f"  {bp[0][0]:.0f} Hz:          {bp[0][1]:.4f}  g²/Hz  (anchor)")
        L.append(f"  {bp[0][0]:.0f}–{bp[1][0]:.0f} Hz:    +3 dB/oct")
        L.append(f"  {bp[1][0]:.0f} Hz:         {bp[1][1]:.4f}  g²/Hz  (flat begins)")
        L.append(f"  {bp[1][0]:.0f}–{bp[2][0]:.0f} Hz:   flat")
        L.append(f"  {bp[2][0]:.0f} Hz:         {bp[2][1]:.4f}  g²/Hz  (flat ends)")
        L.append(f"  {bp[2][0]:.0f}–{bp[3][0]:.0f} Hz:  -6 dB/oct")
        L.append(f"  {bp[3][0]:.0f} Hz:        {bp[3][1]:.5f} g²/Hz")
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
            raw      = self._weight_var.get().strip()
            unit_str = self._unit_var.get()
            w_lb     = details["weight_lb"]
            w_kg     = details["weight_kg"]
            new_flat = details["new_flat"]
            f_break  = details["f_break"]
            red_db   = details["reduction_db"]

            L.append(f"  Input:  {raw} {unit_str}  =  {w_lb:.1f} lb  ({w_kg:.1f} kg)")
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

        path = filedialog.asksaveasfilename(
            title="Export Reduced ASD Profile",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
        )
        if not path:
            return

        grms_out = math.sqrt(max(grms_loglog(rf, ra_scaled), 0.0))
        raw      = self._weight_var.get().strip()
        unit_str = self._unit_var.get()

        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(f"# {self._spec_key} — Weight-Adjusted RV Environment\n")
                f.write(f"# Source: {spec['source']}\n")
                f.write(f"# Baseline:  {spec['baseline_ref']}\n")
                f.write(f"# Reduction: {spec['reduction_ref']}\n")
                f.write(f"# Weight:    {raw} {unit_str}"
                        f"  ({self._weight_lb:.1f} lb)\n")
                f.write(f"# Level:     {_lname}"
                        f"  ({'+' if db_offset >= 0 else ''}{db_offset:.0f} dB)\n")
                if details.get("reduced"):
                    f.write(f"# Reduction: {details['reduction_db']:.1f} dB"
                            f"  (Eq. B.9)\n")
                else:
                    f.write(f"# Reduction: none  (W <= 50 lb)\n")
                f.write(f"# GRMS:      {grms_out:.2f} g\n")
                f.write("#\n")
                f.write("# Freq (Hz)    ASD (g^2/Hz)\n")
                for freq, asd_val in zip(rf, ra_scaled):
                    f.write(f"{freq:10.2f}  {asd_val:.6g}\n")
        except OSError as exc:
            messagebox.showerror("Export Error", str(exc))
            return

        messagebox.showinfo("Exported", f"Saved to:\n{path}")

    # ── misc ──────────────────────────────────────────────────────────────────

    def _toggle_theme(self):
        self._theme = "light" if self._theme == "dark" else "dark"
        self._recompute()

    def _show_help(self):
        _show_popup(self.frame.winfo_toplevel(), "RV Environment — Help", _HELP_TEXT)
