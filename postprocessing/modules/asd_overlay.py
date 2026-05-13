"""ASD Overlay — Acceleration Spectral Density comparison across 1–2 OP2 files.

Reads PSD acceleration results from Nastran random response OP2 files (SOL 111).
Plots ASD (g²/Hz) vs frequency on log-log axes for selected nodes.  Up to two
OP2 files can be overlaid.  RMS is shown in the legend.

Note: this module deliberately avoids importing matplotlib.pyplot so that it
does not interfere with the Agg backend used by the Miles Equation popup.
"""

import csv
import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk
import numpy as np
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure

def _sc_int(key):
    """Normalize a pyNastran result dict key to a plain integer subcase ID."""
    return int(key[0]) if isinstance(key, tuple) else int(key)


def _unique_subcases(result_dict):
    """Return sorted unique integer subcase IDs from a pyNastran result dict."""
    seen, out = set(), []
    for key in sorted(result_dict.keys(), key=_sc_int):
        sc = _sc_int(key)
        if sc not in seen:
            seen.add(sc)
            out.append(sc)
    return out


def _lookup_subcase(result_dict, subcase_int):
    """Fetch a result table by integer subcase ID regardless of key format."""
    if subcase_int in result_dict:
        return result_dict[subcase_int]
    for key, val in result_dict.items():
        if _sc_int(key) == subcase_int:
            return val
    return None


_NODE_COLORS = (
    "#1f77b4", "#ff7f0e", "#2ca02c", "#d62728",
    "#9467bd", "#8c564b", "#e377c2", "#bcbd22",
    "#17becf", "#7f7f7f",
)
_SLOT_TAGS = ("A", "B")
_SLOT_LINES = ("-", "--")

_DARK_BG = "#2b2b2b"
_THEMES = {
    "dark": {
        "fig_bg":    "#2b2b2b",
        "plot_bg":   "#1e1e1e",
        "grid":      "#3a3a3a",
        "text":      "#c0c0c0",
        "spine":     "#505050",
        "legend_bg": "#383838",
    },
    "light": {
        "fig_bg":    "#f5f5f5",
        "plot_bg":   "white",
        "grid":      "#cccccc",
        "text":      "#222222",
        "spine":     "#888888",
        "legend_bg": "white",
    },
}


class AsdOverlayModule:
    name = "ASD Overlay"

    DOF_LABELS = ("T1 (X)", "T2 (Y)", "T3 (Z)")
    UNIT_OPTIONS = ("in/s²", "m/s²")
    UNIT_FACTORS = {"in/s²": 386.089, "m/s²": 9.80665}

    _REF_COLORS = (
        "#000000", "#7f4f24", "#5a189a",
        "#bb3e03", "#005f73", "#404040",
    )

    _GUIDE_TEXT = """\
ASD Overlay Tool — Quick Guide

PURPOSE
Compare Acceleration Spectral Density (ASD) curves from 1 or 2 Nastran
random response OP2 files.  Plots g²/Hz vs frequency on log-log axes.
The RMS value for each curve is shown in the legend.

REQUIREMENTS
OP2 must contain PSD acceleration output.  Required deck entries:
  RANDOM = <sid>               (random analysis flag)
  ACCELERATION(PLOT) = ALL     (or the specific node set)

WORKFLOW
1. Open OP2 A — required first file.  Select subcase and units.
2. Open OP2 B — optional second file for overlay comparison.
3. Add Nodes — paste grid IDs, one per line.  Optional label formats:
     1001
     1001 Tip
     1001, Tip
   Or use "Import" to load a file with columns: grid_id, label
   (header row is auto-detected and skipped).
4. Check/uncheck nodes to show or hide their curves on the plot.
5. Use the DOF dropdown to switch which response direction is plotted.
   All checked nodes across both OP2s update simultaneously.

REFERENCE ASDs
Load one or more reference ASD files (spec envelopes, qual levels) via
the "Reference ASDs" section.  They overlay on the plot for comparison
and do not drive any response calculation.

UNITS
Select the native acceleration unit from the OP2 (matches BDF units).
  in/s²  — divides by 386.089² to convert PSD to g²/Hz
  m/s²   — divides by 9.80665² to convert PSD to g²/Hz

RMS IN LEGEND
The tool first checks whether the OP2 contains a Nastran-integrated RMS
acceleration table (matches F06 GRMS output).  If not present, it falls
back to numerical integration of the displayed ASD curve using the
trapezoidal rule.

LINE STYLES
  Solid lines  — OP2 A
  Dashed lines — OP2 B
  Thick solid  — Reference ASDs

NAVIGATION
Use the matplotlib toolbar below the plot to zoom, pan, and save images.
"""

    def __init__(self, parent):
        self.frame = ctk.CTkFrame(parent)

        # Slot state: {0: slot_A, 1: slot_B}
        self._op2_slots = {0: self._empty_slot(), 1: self._empty_slot()}

        # Node rows
        self._nodes = []

        # Reference ASD rows
        self._refs = []

        self._dof_var = ctk.StringVar(value="T3 (Z)")

        # Per-slot UI widgets populated in _build_ui
        self._open_btn = [None, None]
        self._file_label = [None, None]
        self._unit_var = [ctk.StringVar(value="in/s²"), ctk.StringVar(value="in/s²")]
        self._sc_var = [tk.StringVar(value="(none)"), tk.StringVar(value="(none)")]
        self._sc_menu = [None, None]
        self._mode_var = [ctk.StringVar(value="PSD (RANDOM)"),
                          ctk.StringVar(value="PSD (RANDOM)")]
        self._frf_row = [None, None]
        self._input_asd_btn = [None, None]
        self._input_asd_label = [None, None]

        # Per-slot analysis name (auto-fills from OP2 stem, user-editable)
        self._name_var = [ctk.StringVar(value=""), ctk.StringVar(value="")]
        self._name_user_edited = [False, False]
        self._suppress_name_trace = [False, False]
        for i in range(2):
            self._name_var[i].trace_add(
                "write", lambda *_a, idx=i: self._on_name_var_write(idx))

        self._same_input_asd_var = tk.BooleanVar(value=False)

        self._plot_theme = "dark"
        self._theme_btn = None
        self._view_mode_var = ctk.StringVar(value="Manual")
        self._cycle_index = 0
        self._prev_btn = None
        self._next_btn = None
        self._cycle_label = None
        self._title_var = ctk.StringVar(value="")
        self._env_var = ctk.StringVar(value="")
        self._env_user_edited = False
        self._suppress_env_trace = False
        self._title_var.trace_add("write", self._on_title_var_write)
        self._env_var.trace_add("write", self._on_env_var_write)

        # Plot mode and y-axis scale
        self._plot_mode_var = ctk.StringVar(value="ASD")
        self._yscale_var = ctk.StringVar(value="Log")

        # Peak picking
        self._pick_peaks_mode = False
        self._pick_btn = None
        self._peak_label_style = ctk.StringVar(value="Freq only")
        self._picked_peaks = []
        self._last_drawn_curves = []
        self._mpl_cid_click = None

        self._build_ui()

    @staticmethod
    def _empty_slot():
        return {
            "op2": None, "path": None, "subcase": None,
            "mode": "PSD",
            "input_asd_path": None,
            "input_asd_freqs": None,
            "input_asd_g2hz": None,
        }

    # ── UI construction ──────────────────────────────────────────────────────

    def _build_ui(self):
        # ── Toolbar ──────────────────────────────────────────────────────────
        toolbar = ctk.CTkFrame(self.frame, fg_color="transparent")
        toolbar.pack(fill=tk.X, padx=5, pady=(5, 0))

        # Shared grid so both slot rows have aligned columns
        slot_grid = ctk.CTkFrame(toolbar, fg_color="transparent")
        slot_grid.pack(fill=tk.X, pady=1)
        slot_grid.grid_columnconfigure(1, weight=1)  # file label absorbs slack

        for i, tag in enumerate(_SLOT_TAGS):
            btn = ctk.CTkButton(
                slot_grid, text=f"Open OP2 {tag}…", width=120,
                command=lambda idx=i: self._open_op2(idx),
            )
            btn.grid(row=i, column=0, sticky="w", padx=(0, 6), pady=2)
            self._open_btn[i] = btn

            lbl = ctk.CTkLabel(slot_grid, text="(no file)", text_color="gray",
                               anchor=tk.W)
            lbl.grid(row=i, column=1, sticky="ew", padx=(0, 8))
            self._file_label[i] = lbl

            ctk.CTkLabel(slot_grid, text="Name:").grid(
                row=i, column=2, padx=(0, 2))
            ctk.CTkEntry(slot_grid, textvariable=self._name_var[i], width=130,
                         placeholder_text="Analysis name",
                         ).grid(row=i, column=3, padx=(0, 10))

            ctk.CTkLabel(slot_grid, text="Units:").grid(
                row=i, column=4, padx=(0, 2))
            ctk.CTkOptionMenu(
                slot_grid, variable=self._unit_var[i],
                values=list(self.UNIT_OPTIONS),
                command=lambda _v, idx=i: self._refresh_plot(),
                width=80,
            ).grid(row=i, column=5, padx=(0, 10))

            ctk.CTkLabel(slot_grid, text="Subcase:").grid(
                row=i, column=6, padx=(0, 2))
            scmenu = ctk.CTkOptionMenu(
                slot_grid, variable=self._sc_var[i], values=["(none)"],
                command=lambda _v, idx=i: self._on_sc_select(idx),
                width=100,
            )
            scmenu.grid(row=i, column=7, padx=(0, 10))
            self._sc_menu[i] = scmenu

            ctk.CTkLabel(slot_grid, text="Mode:").grid(
                row=i, column=8, padx=(0, 2))
            ctk.CTkOptionMenu(
                slot_grid, variable=self._mode_var[i],
                values=["PSD (RANDOM)", "FRF + Input ASD"],
                command=lambda _v, idx=i: self._on_mode_change(idx),
                width=140,
            ).grid(row=i, column=9, sticky="w")

            # FRF subrow — hidden until mode is switched
            frf_row = ctk.CTkFrame(toolbar, fg_color="transparent")
            self._frf_row[i] = frf_row

            if i == 1:
                ctk.CTkCheckBox(frf_row, text="Same as A",
                                variable=self._same_input_asd_var,
                                command=self._on_same_input_asd_toggle,
                                ).pack(side=tk.LEFT, padx=(0, 8))

            asd_btn = ctk.CTkButton(
                frf_row, text="Load Input ASD…", width=140,
                command=lambda idx=i: self._load_input_asd(idx),
            )
            asd_btn.pack(side=tk.LEFT, padx=(0, 6))
            self._input_asd_btn[i] = asd_btn
            asd_lbl = ctk.CTkLabel(frf_row, text="(no file)", text_color="gray",
                                   anchor=tk.W, width=200)
            asd_lbl.pack(side=tk.LEFT, padx=(0, 12))
            self._input_asd_label[i] = asd_lbl
            ctk.CTkLabel(frf_row, text="Assumes unit-g base input.",
                         text_color="gray",
                         font=ctk.CTkFont(size=11, slant="italic")).pack(side=tk.LEFT)

        # ── DOF row ──────────────────────────────────────────────────────────
        dof_row = ctk.CTkFrame(toolbar, fg_color="transparent")
        dof_row.pack(fill=tk.X, pady=(4, 0))

        ctk.CTkLabel(dof_row, text="DOF:").pack(side=tk.LEFT, padx=(0, 2))
        ctk.CTkOptionMenu(
            dof_row, variable=self._dof_var, values=list(self.DOF_LABELS),
            command=lambda _: self._refresh_plot(), width=100,
        ).pack(side=tk.LEFT, padx=(0, 12))

        ctk.CTkButton(
            dof_row, text="?", width=30, font=ctk.CTkFont(weight="bold"),
            command=self._show_guide,
        ).pack(side=tk.LEFT, padx=(0, 6))

        self._status_label = ctk.CTkLabel(
            dof_row, text="Open an OP2 to begin", text_color="gray")
        self._status_label.pack(side=tk.LEFT, padx=(10, 0))

        # ── Cycle / View row ─────────────────────────────────────────────────
        view_row = ctk.CTkFrame(toolbar, fg_color="transparent")
        view_row.pack(fill=tk.X, pady=(2, 0))

        ctk.CTkLabel(view_row, text="View:").pack(side=tk.LEFT, padx=(0, 2))
        ctk.CTkOptionMenu(
            view_row, variable=self._view_mode_var,
            values=["Manual", "All grids, cycle DOF",
                    "One grid, cycle DOF×grid", "One grid all DOFs, cycle grid"],
            command=self._on_view_mode_change,
            width=200,
        ).pack(side=tk.LEFT, padx=(0, 10))

        self._prev_btn = ctk.CTkButton(
            view_row, text="◀ Prev", width=70, state=tk.DISABLED,
            command=lambda: self._step_cycle(-1),
        )
        self._prev_btn.pack(side=tk.LEFT, padx=(0, 4))

        self._cycle_label = ctk.CTkLabel(view_row, text="", text_color="gray",
                                         width=260, anchor=tk.W)
        self._cycle_label.pack(side=tk.LEFT, padx=(0, 4))

        self._next_btn = ctk.CTkButton(
            view_row, text="Next ▶", width=70, state=tk.DISABLED,
            command=lambda: self._step_cycle(1),
        )
        self._next_btn.pack(side=tk.LEFT)

        # ── Title / Environment row ───────────────────────────────────────────
        title_row = ctk.CTkFrame(toolbar, fg_color="transparent")
        title_row.pack(fill=tk.X, pady=(2, 4))

        ctk.CTkLabel(title_row, text="Title:").pack(side=tk.LEFT, padx=(0, 2))
        ctk.CTkEntry(title_row, textvariable=self._title_var, width=280,
                     placeholder_text="Plot title").pack(side=tk.LEFT, padx=(0, 16))
        ctk.CTkLabel(title_row, text="Environment:").pack(side=tk.LEFT, padx=(0, 2))
        ctk.CTkEntry(title_row, textvariable=self._env_var, width=220,
                     placeholder_text="Auto-fills on file load").pack(side=tk.LEFT)

        # ── Annotations row ──────────────────────────────────────────────────
        annot_row = ctk.CTkFrame(toolbar, fg_color="transparent")
        annot_row.pack(fill=tk.X, pady=(0, 4))

        ctk.CTkLabel(annot_row, text="Plot:").pack(side=tk.LEFT, padx=(0, 2))
        ctk.CTkOptionMenu(
            annot_row, variable=self._plot_mode_var,
            values=["ASD", "Cumulative GRMS"],
            command=self._on_plot_mode_change,
            width=160,
        ).pack(side=tk.LEFT, padx=(0, 14))

        ctk.CTkLabel(annot_row, text="Y-axis:").pack(side=tk.LEFT, padx=(0, 2))
        ctk.CTkOptionMenu(
            annot_row, variable=self._yscale_var,
            values=["Log", "Linear"],
            command=lambda _: self._refresh_plot(),
            width=80,
        ).pack(side=tk.LEFT, padx=(0, 14))

        self._pick_btn = ctk.CTkButton(
            annot_row, text="Pick Peaks", width=100,
            command=self._toggle_pick_mode,
        )
        self._pick_btn.pack(side=tk.LEFT, padx=(0, 4))

        ctk.CTkButton(annot_row, text="Clear Peaks", width=90,
                      command=self._clear_peaks,
                      ).pack(side=tk.LEFT, padx=(0, 14))

        ctk.CTkLabel(annot_row, text="Label:").pack(side=tk.LEFT, padx=(0, 2))
        ctk.CTkOptionMenu(annot_row, variable=self._peak_label_style,
                          values=["Freq only", "Freq + value"],
                          command=lambda _: self._refresh_plot(),
                          width=120,
                          ).pack(side=tk.LEFT)

        # ── Body: node panel + plot ───────────────────────────────────────────
        body = ctk.CTkFrame(self.frame, fg_color="transparent")
        body.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Left: node management panel
        node_panel = ctk.CTkFrame(body, width=260)
        node_panel.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 5))
        node_panel.pack_propagate(False)

        ctk.CTkLabel(
            node_panel, text="Nodes",
            font=ctk.CTkFont(size=13, weight="bold"),
        ).pack(pady=(8, 4))

        btn_row1 = ctk.CTkFrame(node_panel, fg_color="transparent")
        btn_row1.pack(fill=tk.X, padx=6)
        ctk.CTkButton(btn_row1, text="Add…", width=60,
                      command=self._add_nodes_dialog).pack(side=tk.LEFT)
        ctk.CTkButton(btn_row1, text="Import", width=60,
                      command=self._import_nodes).pack(side=tk.LEFT, padx=3)
        ctk.CTkButton(btn_row1, text="Clear", width=55,
                      command=self._clear_nodes).pack(side=tk.LEFT)

        btn_row2 = ctk.CTkFrame(node_panel, fg_color="transparent")
        btn_row2.pack(fill=tk.X, padx=6, pady=(3, 5))
        ctk.CTkButton(btn_row2, text="All", width=60,
                      command=lambda: self._select_all(True)).pack(side=tk.LEFT)
        ctk.CTkButton(btn_row2, text="None", width=60,
                      command=lambda: self._select_all(False)).pack(side=tk.LEFT, padx=3)

        self._node_scroll = ctk.CTkScrollableFrame(
            node_panel, fg_color="transparent", label_text="")
        self._node_scroll.pack(fill=tk.BOTH, expand=True, padx=4, pady=(0, 5))

        # Reference ASDs section
        ctk.CTkLabel(
            node_panel, text="Reference ASDs",
            font=ctk.CTkFont(size=13, weight="bold"),
        ).pack(pady=(6, 4))

        ref_btn_row = ctk.CTkFrame(node_panel, fg_color="transparent")
        ref_btn_row.pack(fill=tk.X, padx=6)
        ctk.CTkButton(ref_btn_row, text="Load…", width=70,
                      command=self._load_reference_asd).pack(side=tk.LEFT)
        ctk.CTkButton(ref_btn_row, text="Clear", width=55,
                      command=self._clear_references).pack(side=tk.LEFT, padx=3)

        self._ref_scroll = ctk.CTkScrollableFrame(
            node_panel, fg_color="transparent", label_text="", height=120)
        self._ref_scroll.pack(fill=tk.X, padx=4, pady=(2, 6))

        # Right: matplotlib plot
        plot_container = ctk.CTkFrame(body)
        plot_container.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Thin header for theme toggle
        plot_header = ctk.CTkFrame(plot_container, fg_color="transparent",
                                   height=28)
        plot_header.pack(side=tk.TOP, fill=tk.X)
        plot_header.pack_propagate(False)
        self._theme_btn = ctk.CTkButton(
            plot_header, text="☀ Light", width=80,
            command=self._toggle_theme,
        )
        self._theme_btn.pack(side=tk.RIGHT, padx=4, pady=2)

        self._fig = Figure(figsize=(8, 5), dpi=100, facecolor=_DARK_BG)
        self._ax = self._fig.add_subplot(111)

        self._canvas = FigureCanvasTkAgg(self._fig, master=plot_container)
        self._canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        toolbar_tk = tk.Frame(plot_container)
        toolbar_tk.pack(side=tk.BOTTOM, fill=tk.X)
        self._mpl_toolbar = NavigationToolbar2Tk(self._canvas, toolbar_tk)
        self._mpl_toolbar.update()

        self._mpl_cid_click = self._canvas.mpl_connect(
            "button_press_event", self._on_canvas_click)

        self._draw_empty_axes()

    def _toggle_theme(self):
        self._plot_theme = "light" if self._plot_theme == "dark" else "dark"
        self._theme_btn.configure(
            text="☀ Light" if self._plot_theme == "dark" else "☾ Dark")
        self._refresh_plot()

    def _draw_empty_axes(self):
        t = _THEMES[self._plot_theme]
        ax = self._ax
        ax.clear()
        ax.set_facecolor(t["plot_bg"])
        ax.set_xlabel("Frequency (Hz)", color=t["text"])
        ax.set_ylabel("ASD (g²/Hz)", color=t["text"])
        ax.tick_params(colors=t["text"], which="both")
        for spine in ax.spines.values():
            spine.set_edgecolor(t["spine"])
        self._fig.set_facecolor(t["fig_bg"])
        ax.text(0.5, 0.5, "Load an OP2 file and add nodes to begin",
                transform=ax.transAxes,
                ha="center", va="center", color="gray", fontsize=11)
        self._canvas.draw_idle()

    # ── Guide ────────────────────────────────────────────────────────────────

    def _show_guide(self):
        try:
            from structures_tools import show_guide
        except ImportError:
            return
        show_guide(self.frame.winfo_toplevel(), "ASD Overlay Guide",
                   self._GUIDE_TEXT)

    # ── Background threading ─────────────────────────────────────────────────

    def _run_in_background(self, label, work_fn, done_fn):
        self._status_label.configure(text=label, text_color="gray")
        for btn in self._open_btn:
            btn.configure(state=tk.DISABLED)

        container = {}

        def _worker():
            try:
                container['result'] = work_fn()
            except Exception as exc:
                container['error'] = exc

        def _poll():
            if thread.is_alive():
                self.frame.after(50, _poll)
            else:
                for btn in self._open_btn:
                    btn.configure(state=tk.NORMAL)
                done_fn(container.get('result'), container.get('error'))

        thread = threading.Thread(target=_worker, daemon=True)
        thread.start()
        self.frame.after(50, _poll)

    # ── OP2 loading ──────────────────────────────────────────────────────────

    def _open_op2(self, slot_idx):
        tag = _SLOT_TAGS[slot_idx]
        path = filedialog.askopenfilename(
            title=f"Open OP2 {tag}",
            filetypes=[("OP2 files", "*.op2"), ("All files", "*.*")])
        if not path:
            return

        def _work():
            from pyNastran.op2.op2 import OP2
            op2 = OP2(mode='nx', debug=False)
            op2.read_op2(path)
            return op2

        def _done(op2, error):
            if error is not None:
                messagebox.showerror("Error", f"Could not read OP2:\n{error}")
                self._status_label.configure(text="Load failed", text_color="red")
                return

            psd_dict = op2.op2_results.psd.accelerations
            frf_dict = op2.accelerations
            if psd_dict:
                mode = "PSD"
                result_dict = psd_dict
            elif frf_dict:
                mode = "FRF"
                result_dict = frf_dict
            else:
                messagebox.showwarning(
                    "No Acceleration Data",
                    f"OP2 {tag} contains no acceleration results.\n\n"
                    "Ensure the deck includes:\n"
                    "  ACCELERATION(PLOT) = ALL\n\n"
                    "For PSD output also add:\n"
                    "  RANDOM = <sid>")
                self._file_label[slot_idx].configure(
                    text="(no data)", text_color="orange")
                return

            self._op2_slots[slot_idx]['mode'] = mode
            mode_label = "PSD (RANDOM)" if mode == "PSD" else "FRF + Input ASD"
            self._mode_var[slot_idx].set(mode_label)
            if mode == "FRF":
                self._frf_row[slot_idx].pack(fill=tk.X, pady=1)
            else:
                self._frf_row[slot_idx].pack_forget()

            self._op2_slots[slot_idx]['op2'] = op2
            self._op2_slots[slot_idx]['path'] = path

            subcases = _unique_subcases(result_dict)
            sc_strs = [str(s) for s in subcases]
            self._sc_menu[slot_idx].configure(values=sc_strs)
            self._sc_var[slot_idx].set(sc_strs[0])
            self._op2_slots[slot_idx]['subcase'] = subcases[0]

            stem = os.path.splitext(os.path.basename(path))[0]
            self._maybe_autofill_name(slot_idx, stem)
            if mode == "PSD":
                self._maybe_autofill_env(stem)

            self._file_label[slot_idx].configure(
                text=os.path.basename(path), text_color=("gray10", "gray90"))
            self._status_label.configure(
                text=f"OP2 {tag}: {os.path.basename(path)} "
                     f"({len(subcases)} subcase{'s' if len(subcases) != 1 else ''})",
                text_color=("gray10", "gray90"))
            self._refresh_plot()

        self._run_in_background(f"Loading OP2 {tag}…", _work, _done)

    def _on_sc_select(self, slot_idx):
        val = self._sc_var[slot_idx].get()
        try:
            self._op2_slots[slot_idx]['subcase'] = int(val) if val != "(none)" else None
        except ValueError:
            self._op2_slots[slot_idx]['subcase'] = None
        self._refresh_plot()

    def _on_mode_change(self, slot_idx):
        mode_label = self._mode_var[slot_idx].get()
        mode = "FRF" if mode_label == "FRF + Input ASD" else "PSD"
        self._op2_slots[slot_idx]['mode'] = mode
        if mode == "FRF":
            self._frf_row[slot_idx].pack(fill=tk.X, pady=1)
        else:
            self._frf_row[slot_idx].pack_forget()
        self._op2_slots[slot_idx]['op2'] = None
        self._op2_slots[slot_idx]['subcase'] = None
        self._sc_var[slot_idx].set("(none)")
        self._sc_menu[slot_idx].configure(values=["(none)"])
        self._file_label[slot_idx].configure(text="(no file)", text_color="gray")
        self._refresh_plot()

    # ── Input ASD loading ────────────────────────────────────────────────────

    @staticmethod
    def _parse_asd_text_file(path):
        """Parse a 2-column ASD text file (freq, g²/Hz). Returns (freqs, asds) or (None, None)."""
        freqs, asds = [], []
        try:
            with open(path, encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if not line or line.startswith('#') or line.startswith('$'):
                        continue
                    parts = line.replace(',', ' ').split()
                    if len(parts) < 2:
                        continue
                    try:
                        freqs.append(float(parts[0]))
                        asds.append(float(parts[1]))
                    except ValueError:
                        continue
        except Exception as exc:
            messagebox.showerror("Load Error", str(exc))
            return None, None

        if len(freqs) < 2:
            messagebox.showerror("Load Error", "Need at least 2 frequency points.")
            return None, None

        freqs_arr = np.array(freqs)
        asds_arr = np.array(asds)
        order = np.argsort(freqs_arr)
        return freqs_arr[order], asds_arr[order]

    def _load_input_asd(self, slot_idx):
        path = filedialog.askopenfilename(
            title="Load Input ASD",
            filetypes=[("Text files", "*.txt"), ("CSV files", "*.csv"),
                       ("All files", "*.*")])
        if not path:
            return

        freqs_arr, asds_arr = self._parse_asd_text_file(path)
        if freqs_arr is None:
            return

        slot = self._op2_slots[slot_idx]
        slot['input_asd_path'] = path
        slot['input_asd_freqs'] = freqs_arr
        slot['input_asd_g2hz'] = asds_arr

        self._input_asd_label[slot_idx].configure(
            text=f"{os.path.basename(path)} "
                 f"({len(freqs_arr)} pts, "
                 f"{freqs_arr[0]:.1f}–{freqs_arr[-1]:.1f} Hz)",
            text_color=("gray10", "gray90"))
        self._maybe_autofill_env(os.path.splitext(os.path.basename(path))[0])
        self._refresh_plot()

    def _on_same_input_asd_toggle(self):
        use_same = self._same_input_asd_var.get()
        self._input_asd_btn[1].configure(
            state=tk.DISABLED if use_same else tk.NORMAL)
        if use_same:
            a = self._op2_slots[0]
            if a['input_asd_freqs'] is None:
                self._input_asd_label[1].configure(
                    text="(slot A has no Input ASD)", text_color="orange")
            else:
                self._input_asd_label[1].configure(
                    text=f"← same as A: {os.path.basename(a['input_asd_path'])}",
                    text_color=("gray10", "gray90"))
        else:
            b = self._op2_slots[1]
            if b['input_asd_freqs'] is not None:
                self._input_asd_label[1].configure(
                    text=f"{os.path.basename(b['input_asd_path'])} "
                         f"({len(b['input_asd_freqs'])} pts)",
                    text_color=("gray10", "gray90"))
            else:
                self._input_asd_label[1].configure(
                    text="(no file)", text_color="gray")
        self._refresh_plot()

    # ── Reference ASD management ─────────────────────────────────────────────

    def _load_reference_asd(self):
        path = filedialog.askopenfilename(
            title="Load Reference ASD",
            filetypes=[("Text/CSV files", "*.txt *.csv"),
                       ("All files", "*.*")])
        if not path:
            return
        freqs, asds = self._parse_asd_text_file(path)
        if freqs is None:
            return
        name = os.path.splitext(os.path.basename(path))[0]
        var = tk.BooleanVar(value=True)
        name_var = tk.StringVar(value=name)
        row_frame = ctk.CTkFrame(self._ref_scroll, fg_color="transparent")
        row_frame.pack(fill=tk.X, pady=1)
        ref = {"path": path, "name": name, "freqs": freqs, "g2hz": asds,
               "checked": var, "name_var": name_var, "row_frame": row_frame}
        self._refs.append(ref)

        ctk.CTkCheckBox(row_frame, text="", variable=var, width=24,
                        command=self._refresh_plot).pack(side=tk.LEFT, padx=(2, 0))
        name_entry = ctk.CTkEntry(row_frame, textvariable=name_var, width=152)
        name_entry.pack(side=tk.LEFT, padx=(2, 2))
        name_entry.bind("<Return>",   lambda _e, r=ref: self._commit_ref_name(r))
        name_entry.bind("<FocusOut>", lambda _e, r=ref: self._commit_ref_name(r))
        ctk.CTkButton(row_frame, text="✕", width=22,
                      command=lambda r=ref: self._remove_reference(r),
                      fg_color="transparent", hover_color=("gray75", "gray30"),
                      text_color=("gray40", "gray60"),
                      ).pack(side=tk.LEFT)
        self._refresh_plot()

    def _remove_reference(self, ref):
        ref['row_frame'].destroy()
        self._refs.remove(ref)
        self._refresh_plot()

    def _clear_references(self):
        for r in self._refs:
            r['row_frame'].destroy()
        self._refs.clear()
        self._refresh_plot()

    def _commit_ref_name(self, ref):
        new_name = ref['name_var'].get().strip() or ref['name']
        if new_name == ref['name']:
            return
        ref['name'] = new_name
        ref['name_var'].set(new_name)
        self._refresh_plot()

    # ── Node management ──────────────────────────────────────────────────────

    def _add_nodes_dialog(self):
        dlg = ctk.CTkToplevel(self.frame.winfo_toplevel())
        dlg.title("Add Nodes")
        dlg.geometry("380x310")
        dlg.transient(self.frame.winfo_toplevel())
        dlg.grab_set()

        ctk.CTkLabel(
            dlg,
            text="Enter one node per line.  Optional label after the ID:\n"
                 "  1001        1001 Tip        1001, Tip",
            justify=tk.LEFT,
            anchor=tk.W,
        ).pack(padx=12, pady=(12, 4), fill=tk.X)

        tb = ctk.CTkTextbox(dlg, wrap="none")
        tb.pack(fill=tk.BOTH, expand=True, padx=12)

        def _ok():
            self._parse_and_add_nodes(tb.get("1.0", "end"))
            dlg.destroy()

        btn_row = ctk.CTkFrame(dlg, fg_color="transparent")
        btn_row.pack(fill=tk.X, padx=12, pady=8)
        ctk.CTkButton(btn_row, text="Add", command=_ok).pack(side=tk.LEFT)
        ctk.CTkButton(btn_row, text="Cancel",
                      command=dlg.destroy).pack(side=tk.LEFT, padx=5)
        dlg.bind("<Return>", lambda _: _ok())

    def _parse_and_add_nodes(self, text):
        existing_ids = {n['id'] for n in self._nodes}
        added = False
        for raw in text.splitlines():
            line = raw.strip()
            if not line:
                continue
            parts = [p.strip() for p in line.split(',', 1)] if ',' in line \
                else line.split(None, 1)
            try:
                gid = int(parts[0])
            except (ValueError, IndexError):
                continue
            if gid in existing_ids:
                continue
            existing_ids.add(gid)
            label = parts[1].strip() if len(parts) > 1 and parts[1].strip() \
                else f"Node {gid}"
            self._add_node_row(gid, label)
            added = True

        if added:
            self._refresh_plot()

    def _import_nodes(self):
        path = filedialog.askopenfilename(
            title="Import Nodes",
            filetypes=[("Text files", "*.txt"),
                       ("CSV files", "*.csv"),
                       ("All files", "*.*")])
        if not path:
            return
        ext = os.path.splitext(path)[1].lower()
        try:
            if ext == '.csv':
                text = self._read_csv_as_text(path)
            else:
                text = self._read_text_node_file(path)
        except Exception as exc:
            messagebox.showerror("Import Error", str(exc))
            return
        self._parse_and_add_nodes(text)

    @staticmethod
    def _read_csv_as_text(path):
        """Parse a CSV node file; return 'gid,label' lines joined by newline."""
        with open(path, newline='', encoding='utf-8-sig') as f:
            sample = f.read(4096)
            f.seek(0)
            try:
                dialect = csv.Sniffer().sniff(sample)
                has_header = csv.Sniffer().has_header(sample)
            except csv.Error:
                dialect = csv.excel
                has_header = False
            reader = csv.reader(f, dialect)
            if has_header:
                next(reader, None)
            lines = []
            for row in reader:
                if not row:
                    continue
                try:
                    gid = int(str(row[0]).strip())
                except (ValueError, IndexError):
                    continue
                label = row[1].strip() if len(row) > 1 and str(row[1]).strip() \
                    else ""
                lines.append(f"{gid},{label}" if label else str(gid))
        return "\n".join(lines)

    @staticmethod
    def _read_text_node_file(path):
        """Parse a loose-format text node file; return 'gid,label' lines."""
        lines = []
        with open(path, encoding='utf-8') as f:
            for raw in f:
                line = raw.strip()
                if not line or line.startswith('#') or line.startswith('$'):
                    continue
                parts = [p.strip() for p in line.split(',', 1)] if ',' in line \
                    else line.split(None, 1)
                try:
                    gid = int(parts[0])
                except (ValueError, IndexError):
                    continue
                label = parts[1].strip() if len(parts) > 1 and parts[1].strip() \
                    else ""
                lines.append(f"{gid},{label}" if label else str(gid))
        return "\n".join(lines)

    def _add_node_row(self, gid, label):
        var = tk.BooleanVar(value=True)
        gid_var = tk.StringVar(value=str(gid))
        label_var = tk.StringVar(value=label)

        row_frame = ctk.CTkFrame(self._node_scroll, fg_color="transparent")
        row_frame.pack(fill=tk.X, pady=1)

        node = {"id": gid, "label": label,
                "checked": var, "row_frame": row_frame,
                "gid_var": gid_var, "label_var": label_var}
        self._nodes.append(node)

        ctk.CTkCheckBox(row_frame, text="", variable=var, width=24,
                        command=self._refresh_plot).pack(side=tk.LEFT, padx=(2, 0))

        gid_entry = ctk.CTkEntry(row_frame, textvariable=gid_var, width=90)
        gid_entry.pack(side=tk.LEFT, padx=(2, 2))
        gid_entry.bind("<Return>",   lambda _e, n=node: self._commit_node_gid(n))
        gid_entry.bind("<FocusOut>", lambda _e, n=node: self._commit_node_gid(n))

        lbl_entry = ctk.CTkEntry(row_frame, textvariable=label_var, width=100)
        lbl_entry.pack(side=tk.LEFT, padx=(0, 2))
        lbl_entry.bind("<Return>",   lambda _e, n=node: self._commit_node_label(n))
        lbl_entry.bind("<FocusOut>", lambda _e, n=node: self._commit_node_label(n))

        ctk.CTkButton(row_frame, text="✕", width=22,
                      command=lambda n=node: self._remove_node(n),
                      fg_color="transparent", hover_color=("gray75", "gray30"),
                      text_color=("gray40", "gray60"),
                      ).pack(side=tk.LEFT)

    def _clear_nodes(self):
        for n in self._nodes:
            n['row_frame'].destroy()
        self._nodes.clear()
        self._refresh_plot()

    def _select_all(self, state):
        for n in self._nodes:
            n['checked'].set(state)
        self._refresh_plot()

    def _remove_node(self, node):
        node['row_frame'].destroy()
        self._nodes.remove(node)
        self._refresh_plot()

    def _commit_node_gid(self, node):
        raw = node['gid_var'].get().strip()
        try:
            new_gid = int(raw)
        except ValueError:
            node['gid_var'].set(str(node['id']))
            return
        if new_gid == node['id']:
            return
        if any(n is not node and n['id'] == new_gid for n in self._nodes):
            node['gid_var'].set(str(node['id']))
            return
        if node['label'] == f"Node {node['id']}":
            node['label'] = f"Node {new_gid}"
            node['label_var'].set(node['label'])
        node['id'] = new_gid
        self._refresh_plot()

    def _commit_node_label(self, node):
        new_label = node['label_var'].get().strip() or f"Node {node['id']}"
        if new_label == node['label']:
            return
        node['label'] = new_label
        node['label_var'].set(new_label)
        self._refresh_plot()

    # ── Title / Environment / Name helpers ───────────────────────────────────

    def _on_title_var_write(self, *_):
        self._refresh_plot()

    def _on_env_var_write(self, *_):
        if self._suppress_env_trace:
            return
        self._env_user_edited = True
        self._refresh_plot()

    def _maybe_autofill_env(self, name):
        if self._env_user_edited or not name:
            return
        self._suppress_env_trace = True
        try:
            self._env_var.set(name)
        finally:
            self._suppress_env_trace = False
        self._refresh_plot()

    def _on_name_var_write(self, slot_idx):
        if self._suppress_name_trace[slot_idx]:
            return
        self._name_user_edited[slot_idx] = True
        self._refresh_plot()

    def _maybe_autofill_name(self, slot_idx, name):
        if self._name_user_edited[slot_idx] or not name:
            return
        self._suppress_name_trace[slot_idx] = True
        try:
            self._name_var[slot_idx].set(name)
        finally:
            self._suppress_name_trace[slot_idx] = False

    # ── RMS helpers ──────────────────────────────────────────────────────────

    @staticmethod
    def _grms_loglog(freqs, asd):
        """Area under an ASD curve using analytical log-log segment integration (FEMCI)."""
        area = 0.0
        for i in range(len(freqs) - 1):
            fl, fh = float(freqs[i]), float(freqs[i + 1])
            al, ah = float(asd[i]), float(asd[i + 1])
            if fl <= 0 or fh <= 0 or al <= 0 or ah <= 0:
                continue
            log_f = np.log(fh / fl)
            b = np.log(ah / al) / log_f if log_f != 0 else 0.0
            if abs(b + 1.0) < 1e-6:
                area += al * fl * log_f
            else:
                area += (ah * fh - al * fl) / (b + 1.0)
        return area

    @staticmethod
    def _cumulative_grms_loglog(freqs, asd):
        """Cumulative GRMS array using FEMCI log-log integration. cum[0] = 0."""
        cum_area = np.zeros(len(freqs))
        running = 0.0
        for i in range(len(freqs) - 1):
            fl, fh = float(freqs[i]), float(freqs[i + 1])
            al, ah = float(asd[i]), float(asd[i + 1])
            if fl <= 0 or fh <= 0 or al <= 0 or ah <= 0:
                cum_area[i + 1] = running
                continue
            log_f = np.log(fh / fl)
            b = np.log(ah / al) / log_f if log_f != 0 else 0.0
            if abs(b + 1.0) < 1e-6:
                running += al * fl * log_f
            else:
                running += (ah * fh - al * fl) / (b + 1.0)
            cum_area[i + 1] = running
        return np.sqrt(np.maximum(cum_area, 0.0))

    @staticmethod
    def _interp_input_asd_to_grid(slot, op2_freqs):
        """Log-log interpolate slot's input ASD onto op2_freqs. Out-of-range → 0."""
        f_in = slot['input_asd_freqs']
        a_in = slot['input_asd_g2hz']
        result = np.zeros(len(op2_freqs))
        for i, f in enumerate(op2_freqs):
            if f < f_in[0] or f > f_in[-1]:
                result[i] = 0.0
                continue
            idx = int(np.searchsorted(f_in, f, side='right')) - 1
            idx = min(idx, len(f_in) - 2)
            fl, fh = f_in[idx], f_in[idx + 1]
            al, ah = a_in[idx], a_in[idx + 1]
            if fl <= 0 or fh <= 0 or al <= 0 or ah <= 0 or fl == fh:
                result[i] = al
            else:
                b = np.log(ah / al) / np.log(fh / fl)
                result[i] = al * (f / fl) ** b
        return result

    def _get_response_psd(self, slot_idx, subcase, nid, idof, unit_factor):
        """Return (freqs, data, is_psd) for one node/DOF."""
        slot = self._op2_slots[slot_idx]
        op2 = slot['op2']

        if slot['mode'] == "PSD":
            psd_tbl = _lookup_subcase(op2.op2_results.psd.accelerations, subcase)
            if psd_tbl is None:
                return None, None, None
            freqs = psd_tbl._times
            op2_nids = psd_tbl.node_gridtype[:, 0]
            hits = np.where(op2_nids == nid)[0]
            if not len(hits):
                return None, None, None
            raw_psd = psd_tbl.data[:, hits[0], idof]
            return freqs, raw_psd / (unit_factor ** 2), True
        else:  # FRF
            if not op2.accelerations:
                return None, None, None
            frf_tbl = _lookup_subcase(op2.accelerations, subcase)
            if frf_tbl is None:
                return None, None, None
            freqs = frf_tbl._times
            op2_nids = frf_tbl.node_gridtype[:, 0]
            hits = np.where(op2_nids == nid)[0]
            if not len(hits):
                return None, None, None
            H_native = frf_tbl.data[:, hits[0], idof]
            H_g = H_native / unit_factor

            # Borrow slot A's input ASD for slot B when "Same as A" is checked
            if slot_idx == 1 and self._same_input_asd_var.get():
                slot_asd = self._op2_slots[0]
            else:
                slot_asd = slot

            if slot_asd['input_asd_freqs'] is not None:
                H_mag2 = H_g.real ** 2 + H_g.imag ** 2
                S_in = self._interp_input_asd_to_grid(slot_asd, freqs)
                return freqs, H_mag2 * S_in, True
            else:
                return freqs, np.abs(H_g), False

    @staticmethod
    def _get_rms_g(op2, subcase, nid, idof, freqs, psd_curve_g2hz, unit_factor):
        """Return RMS in g; prefers Nastran RMS table, falls back to FEMCI integration."""
        try:
            rms_tbl = _lookup_subcase(op2.op2_results.rms.accelerations, subcase)
            if rms_tbl is not None:
                nids_arr = rms_tbl.node_gridtype[:, 0]
                hits = np.where(nids_arr == nid)[0]
                if len(hits):
                    rms_native = float(rms_tbl.data[0, hits[0], idof])
                    return rms_native / unit_factor
        except Exception:
            pass
        area = AsdOverlayModule._grms_loglog(freqs, psd_curve_g2hz)
        return float(np.sqrt(area))

    # ── Cycle helpers ────────────────────────────────────────────────────────

    def _get_plot_frames(self):
        checked = [(n['id'], n['label']) for n in self._nodes if n['checked'].get()]
        cur_dof = self.DOF_LABELS.index(self._dof_var.get())
        mode = self._view_mode_var.get()

        if mode == "Manual":
            return [[(nid, lbl, cur_dof) for nid, lbl in checked]], [""]

        if mode == "All grids, cycle DOF":
            return (
                [[(nid, lbl, d) for nid, lbl in checked] for d in range(3)],
                [self.DOF_LABELS[d] for d in range(3)],
            )

        if mode == "One grid, cycle DOF×grid":
            frames, descs = [], []
            for nid, lbl in checked:
                for d in range(3):
                    frames.append([(nid, lbl, d)])
                    descs.append(f"Node {nid} {lbl} — {self.DOF_LABELS[d]}")
            return frames, descs

        if mode == "One grid all DOFs, cycle grid":
            return (
                [[(nid, lbl, d) for d in range(3)] for nid, lbl in checked],
                [f"Node {nid} {lbl}" for nid, lbl in checked],
            )

        return [[(nid, lbl, cur_dof) for nid, lbl in checked]], [""]

    def _update_cycle_controls(self, frames, descs):
        mode = self._view_mode_var.get()
        n = len(frames)
        i = self._cycle_index

        if mode == "Manual" or n <= 1:
            self._cycle_label.configure(text="")
            self._prev_btn.configure(state=tk.DISABLED)
            self._next_btn.configure(state=tk.DISABLED)
            return

        desc = descs[i] if i < len(descs) else ""
        self._cycle_label.configure(text=f"{i + 1} of {n}: {desc}")
        self._prev_btn.configure(state=tk.NORMAL if i > 0 else tk.DISABLED)
        self._next_btn.configure(state=tk.NORMAL if i < n - 1 else tk.DISABLED)

    def _on_view_mode_change(self, _val=None):
        self._cycle_index = 0
        self._refresh_plot()

    def _step_cycle(self, delta):
        self._cycle_index += delta
        self._refresh_plot()

    # ── Peak picking ─────────────────────────────────────────────────────────

    def _toggle_pick_mode(self):
        self._pick_peaks_mode = not self._pick_peaks_mode
        if self._pick_peaks_mode:
            self._pick_btn.configure(fg_color=("#1f6aa5", "#1f6aa5"),
                                     text="Pick Peaks (ON)")
            self._canvas.get_tk_widget().configure(cursor="cross")
        else:
            self._pick_btn.configure(fg_color=("gray75", "gray25"),
                                     text="Pick Peaks")
            self._canvas.get_tk_widget().configure(cursor="")

    def _clear_peaks(self):
        self._picked_peaks.clear()
        self._refresh_plot()

    def _on_canvas_click(self, event):
        if not self._pick_peaks_mode:
            return
        if self._plot_mode_var.get() == "Cumulative GRMS":
            return
        if event.inaxes is None or event.xdata is None or event.ydata is None:
            return
        if event.xdata <= 0 or event.ydata <= 0:
            return

        from scipy.signal import find_peaks

        click_lx = np.log10(event.xdata)
        click_ly = np.log10(event.ydata)

        best = None
        for curve in self._last_drawn_curves:
            if not curve["is_psd"]:
                continue
            freqs = curve["freqs"]
            data = curve["data"]
            peak_idx, _ = find_peaks(data)
            if not len(peak_idx):
                continue
            for i in peak_idx:
                f, v = float(freqs[i]), float(data[i])
                if f <= 0 or v <= 0:
                    continue
                d = ((np.log10(f) - click_lx) ** 2
                     + (np.log10(v) - click_ly) ** 2) ** 0.5
                if best is None or d < best[0]:
                    best = (d, curve["slot_idx"], curve["nid"],
                            curve["idof"], f, v)

        if best is None or best[0] > 0.5:
            return
        self._picked_peaks.append({
            "slot_idx": best[1], "nid": best[2], "idof": best[3],
            "freq": best[4], "value": best[5],
        })
        self._refresh_plot()

    # ── Plot mode ─────────────────────────────────────────────────────────────

    def _on_plot_mode_change(self, _val=None):
        self._refresh_plot()

    # ── Plot ─────────────────────────────────────────────────────────────────

    def _refresh_plot(self):
        t = _THEMES[self._plot_theme]
        ax = self._ax
        ax.clear()
        ax.set_facecolor(t["plot_bg"])

        plot_mode = self._plot_mode_var.get()   # "ASD" | "Cumulative GRMS"
        yscale = self._yscale_var.get().lower() # "log" | "linear"
        is_cum = (plot_mode == "Cumulative GRMS")

        ax.set_xscale("log")
        ax.set_yscale(yscale)

        frames, descs = self._get_plot_frames()
        if not frames:
            frames, descs = [[]], [""]
        self._cycle_index = max(0, min(self._cycle_index, len(frames) - 1))
        curves = frames[self._cycle_index]

        view_mode = self._view_mode_var.get()
        color_by_dof = (view_mode == "One grid all DOFs, cycle grid")

        if view_mode in ("All grids, cycle DOF", "One grid, cycle DOF×grid"):
            if curves:
                active_dof = curves[0][2]
                self._dof_var.set(self.DOF_LABELS[active_dof])

        has_curves = False
        has_psd = False
        has_frf_mag = False
        self._last_drawn_curves = []

        for slot_idx, slot in self._op2_slots.items():
            op2 = slot['op2']
            subcase = slot['subcase']
            if op2 is None or subcase is None:
                continue

            unit_factor = self.UNIT_FACTORS[self._unit_var[slot_idx].get()]
            tag = _SLOT_TAGS[slot_idx]
            ls = _SLOT_LINES[slot_idx]
            name = self._name_var[slot_idx].get().strip() or tag

            for curve_idx, (nid, lbl, idof) in enumerate(curves):
                color = (_NODE_COLORS[idof % len(_NODE_COLORS)] if color_by_dof
                         else _NODE_COLORS[curve_idx % len(_NODE_COLORS)])

                freqs, data, is_psd = self._get_response_psd(
                    slot_idx, subcase, nid, idof, unit_factor)
                if freqs is None:
                    continue

                dof_label = self.DOF_LABELS[idof]

                if is_psd:
                    if is_cum:
                        plot_data = self._cumulative_grms_loglog(freqs, data)
                        final_g = float(plot_data[-1]) if len(plot_data) else 0.0
                        label = f"{name}: {lbl} {dof_label}  (final GRMS = {final_g:.3g} g)"
                    else:
                        plot_data = data
                        rms_g = self._get_rms_g(
                            op2, subcase, nid, idof, freqs, data, unit_factor)
                        label = f"{name}: {lbl} {dof_label}  (RMS = {rms_g:.3g} g)"
                    has_psd = True
                else:
                    if is_cum:
                        continue  # cumulative only for PSD
                    plot_data = data
                    label = f"{name}: {lbl} {dof_label}  (FRF magnitude)"
                    has_frf_mag = True

                ax.plot(freqs, plot_data, label=label, color=color, linestyle=ls)
                has_curves = True
                self._last_drawn_curves.append({
                    "slot_idx": slot_idx, "nid": nid, "idof": idof,
                    "freqs": np.asarray(freqs), "data": np.asarray(data),
                    "is_psd": is_psd, "color": color,
                })

        # ── Reference ASD overlays ────────────────────────────────────────────
        for ref_idx, ref in enumerate(self._refs):
            if not ref['checked'].get():
                continue
            color = self._REF_COLORS[ref_idx % len(self._REF_COLORS)]
            rname = ref['name_var'].get().strip() or ref['name']
            if is_cum:
                ref_plot = self._cumulative_grms_loglog(ref['freqs'], ref['g2hz'])
                final = float(ref_plot[-1]) if len(ref_plot) else 0.0
                ref_label = f"[Ref] {rname}  (final GRMS = {final:.3g} g)"
            else:
                ref_plot = ref['g2hz']
                grms = float(np.sqrt(max(self._grms_loglog(ref['freqs'], ref['g2hz']), 0.0)))
                ref_label = f"[Ref] {rname}  (GRMS = {grms:.3g} g)"
            ax.plot(ref['freqs'], ref_plot, label=ref_label, color=color,
                    linestyle="-", linewidth=2.0, alpha=0.65)
            has_curves = True

        ax.set_xlabel("Frequency (Hz)", color=t["text"])
        if is_cum:
            ax.set_ylabel("Cumulative GRMS (g)", color=t["text"])
        elif has_psd and has_frf_mag:
            ax.set_ylabel("ASD (g²/Hz) / FRF (g/g)", color=t["text"])
        elif has_frf_mag:
            ax.set_ylabel("FRF Magnitude (g/g)", color=t["text"])
        else:
            ax.set_ylabel("ASD (g²/Hz)", color=t["text"])

        ax.tick_params(colors=t["text"], which="both")
        for spine in ax.spines.values():
            spine.set_edgecolor(t["spine"])
        ax.grid(True, which="both", alpha=0.3, color=t["grid"])
        self._fig.set_facecolor(t["fig_bg"])

        if has_curves:
            ax.legend(loc="best", fontsize=8,
                      facecolor=t["legend_bg"], labelcolor=t["text"],
                      edgecolor=t["spine"])
        else:
            ax.text(0.5, 0.5,
                    "No data — check OP2 loaded, nodes added, and boxes checked",
                    transform=ax.transAxes,
                    ha="center", va="center", color="gray", fontsize=10)

        # ── Picked peak markers and labels (ASD mode only) ────────────────────
        if not is_cum:
            label_style = self._peak_label_style.get()
            for pk in self._picked_peaks:
                curve = next((c for c in self._last_drawn_curves
                              if c["slot_idx"] == pk["slot_idx"]
                              and c["nid"] == pk["nid"]
                              and c["idof"] == pk["idof"]), None)
                if curve is None:
                    continue
                f, v = pk["freq"], pk["value"]
                ax.plot([f], [v], marker="o", markersize=6,
                        markerfacecolor=curve["color"],
                        markeredgecolor=t["text"], linestyle="none", zorder=5)
                if label_style == "Freq + value":
                    ann_text = f"{f:.0f} Hz\n{v:.3g} g²/Hz"
                else:
                    ann_text = f"{f:.0f} Hz"
                ax.annotate(ann_text, xy=(f, v),
                            xytext=(6, 6), textcoords="offset points",
                            color=t["text"], fontsize=9,
                            bbox=dict(boxstyle="round,pad=0.2",
                                      fc=t["legend_bg"], ec=t["spine"], alpha=0.85))

        # ── Title and environment ─────────────────────────────────────────────
        plot_title = self._title_var.get().strip()
        env_name = self._env_var.get().strip()
        ax.set_title(
            plot_title or "",
            color=t["text"],
            fontsize=13,
            weight="bold",
            pad=20 if env_name else 6,
        )
        if env_name:
            ax.text(0.5, 1.01, env_name,
                    transform=ax.transAxes,
                    ha="center", va="bottom",
                    color=t["text"], fontsize=10,
                    style="italic", alpha=0.75)

        self._canvas.draw_idle()
        self._update_cycle_controls(frames, descs)
