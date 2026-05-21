"""FRF Response Limiting — compute a notched drive ASD from FRF data.

Loads a Nastran SOL 111 frequency-response OP2, an input environment ASD,
and a response limit ASD.  Node/DOF rows are added to a list; rows flagged
as 'Limit' drive the notch calculation while all 'Show' rows have their
response curves plotted before and after notching.
"""

import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk
import numpy as np
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure
from tksheet import Sheet

from .asd_common import (
    RESPONSE_TYPES,
    subcase_options,
    lookup_subcase,
    parse_asd_text,
    parse_asd_file,
    interp_loglog,
    grms_loglog,
)


_DOF_NAMES = ["X", "Y", "Z"]

_NODE_COLORS = (
    "#1f77b4", "#2ca02c", "#9467bd", "#17becf", "#bcbd22",
    "#d62728", "#ff7f0e", "#8c564b", "#e377c2", "#7f7f7f",
)

_THEMES = {
    "dark": {
        "fig_bg": "#2b2b2b", "plot_bg": "#1e1e1e", "grid": "#3a3a3a",
        "text": "#c0c0c0", "spine": "#505050", "legend_bg": "#383838",
    },
    "light": {
        "fig_bg": "#f5f5f5", "plot_bg": "white", "grid": "#cccccc",
        "text": "#222222", "spine": "#888888", "legend_bg": "white",
    },
}


class ResponseLimitingModule:
    name = "Response Limiting"

    _GUIDE_TEXT = """\
Response Limiting — Quick Guide

PURPOSE
  Compute a notched drive ASD so that the response at selected nodes/DOFs
  stays at or below a specified limit spectrum.  All added nodes are plotted;
  nodes flagged 'Limit' drive the notch calculation.

WORKFLOW
  1. Open a SOL 111 FRF OP2 (ACCELERATION(PLOT,PHASE)=ALL required).
  2. Load or paste the Input ASD (freq vs g²/Hz) — the baseline drive spec.
  3. Load or paste the Response Limit ASD (freq vs g²/Hz).
  4. Paste or import (node_id  direction) rows, e.g.:
         1001 X
         1001 Z
         1042 Y
     Check 'Limit' on the rows you want to notch against.
  5. Click Compute Notch.
  6. Use the View radio buttons to explore results.
  7. Export Excel or CSV for the notched ASD.

MATH
  At each frequency f, for each Limit DOF j:
    S_allowed_j(f) = S_limit(f) / |H_j(f)|²
  Notched input = min(S_original, min_j S_allowed_j)
  Optional floor: notched ≥ S_original × 10^(-D/10)

NODE ROWS
  Show ☑  — include this curve in response plots
  Limit ☑ — include this DOF in the notch calculation
  A node can be Show-only (no Limit) to observe its response under the
  notched input without contributing to the notch.

UNITS
  Set Units to match the OP2 output (in/s² for slinch/inch models).
  All plotted ASDs are in g²/Hz; GRMS in g.
"""

    def __init__(self, parent):
        self.frame = ctk.CTkFrame(parent)

        # ── State ─────────────────────────────────────────────────────────
        self._op2 = None
        self._op2_path = None
        self._subcase_opts = []
        self._subcase_var = ctk.StringVar(value="—")
        self._units_var = ctk.StringVar(value="in/s²")
        self._rt_var = ctk.StringVar(value="Acceleration")

        self._input_asd_freqs = None
        self._input_asd_vals = None
        self._limit_asd_freqs = None
        self._limit_asd_vals = None

        # Node rows: list of dicts with nid, show_var, dir_var, limit_var, color, row_frame
        self._node_rows = []

        self._notch_enabled_var = ctk.BooleanVar(value=False)
        self._notch_db_var = ctk.StringVar(value="6.0")

        # Computed results (on FRF frequency grid)
        self._frf_freqs = None
        self._orig_asd_interp = None
        self._limit_asd_interp = None
        self._notched_asd = None
        self._response_curves = {}   # {(nid, dof_idx): (resp_before, resp_after)}
        self._limit_dofs_set = set() # (nid, dof_idx) pairs that were limit nodes at compute time

        self._view_var = ctk.StringVar(value="input")
        self._dof_picker_var = ctk.StringVar(value="—")

        self._build_ui()

    # ── UI construction ────────────────────────────────────────────────────

    def _build_ui(self):
        toolbar = ctk.CTkFrame(self.frame, fg_color="transparent")
        toolbar.pack(fill=tk.X, padx=6, pady=(6, 2))

        # Row 0: OP2 controls + help
        row0 = ctk.CTkFrame(toolbar, fg_color="transparent")
        row0.pack(fill=tk.X, pady=1)

        self._open_btn = ctk.CTkButton(row0, text="Open OP2…", width=110,
                                       command=self._open_op2)
        self._open_btn.pack(side=tk.LEFT, padx=(0, 4))
        ctk.CTkButton(row0, text="Clear", width=60,
                      command=self._clear_op2).pack(side=tk.LEFT, padx=(0, 10))

        ctk.CTkLabel(row0, text="Run:").pack(side=tk.LEFT)
        self._file_label = ctk.CTkLabel(row0, text="(none)", text_color="gray",
                                        width=180, anchor=tk.W)
        self._file_label.pack(side=tk.LEFT, padx=4)

        ctk.CTkLabel(row0, text="Units:").pack(side=tk.LEFT, padx=(8, 2))
        self._units_menu = ctk.CTkOptionMenu(row0, variable=self._units_var,
                                             values=["in/s²", "m/s²"], width=90)
        self._units_menu.pack(side=tk.LEFT)

        ctk.CTkLabel(row0, text="Subcase:").pack(side=tk.LEFT, padx=(8, 2))
        self._sc_menu = ctk.CTkOptionMenu(row0, variable=self._subcase_var,
                                          values=["—"],
                                          command=self._on_subcase_change,
                                          width=160)
        self._sc_menu.pack(side=tk.LEFT)

        self._status_label = ctk.CTkLabel(row0, text="Load an FRF OP2 to begin",
                                          text_color="gray")
        self._status_label.pack(side=tk.LEFT, padx=(12, 0))

        ctk.CTkButton(row0, text="?", width=28,
                      command=self._open_help).pack(side=tk.RIGHT)

        # Row 1: ASD loaders
        row1 = ctk.CTkFrame(toolbar, fg_color="transparent")
        row1.pack(fill=tk.X, pady=1)

        input_group = ctk.CTkFrame(row1, fg_color="transparent")
        input_group.pack(side=tk.LEFT, padx=(0, 16))
        ctk.CTkLabel(input_group, text="Input ASD:", width=72, anchor=tk.W).pack(side=tk.LEFT)
        ctk.CTkButton(input_group, text="Load…", width=60,
                      command=self._load_input_asd).pack(side=tk.LEFT, padx=2)
        ctk.CTkButton(input_group, text="Paste…", width=60,
                      command=self._paste_input_asd).pack(side=tk.LEFT, padx=2)
        ctk.CTkButton(input_group, text="Clear", width=50,
                      command=self._clear_input_asd).pack(side=tk.LEFT, padx=2)
        self._input_status = ctk.CTkLabel(input_group, text="(none)", text_color="gray",
                                          width=200, anchor=tk.W)
        self._input_status.pack(side=tk.LEFT, padx=4)

        limit_group = ctk.CTkFrame(row1, fg_color="transparent")
        limit_group.pack(side=tk.LEFT)
        ctk.CTkLabel(limit_group, text="Response Limit:", width=104, anchor=tk.W).pack(side=tk.LEFT)
        ctk.CTkButton(limit_group, text="Load…", width=60,
                      command=self._load_limit_asd).pack(side=tk.LEFT, padx=2)
        ctk.CTkButton(limit_group, text="Paste…", width=60,
                      command=self._paste_limit_asd).pack(side=tk.LEFT, padx=2)
        ctk.CTkButton(limit_group, text="Clear", width=50,
                      command=self._clear_limit_asd).pack(side=tk.LEFT, padx=2)
        self._limit_status = ctk.CTkLabel(limit_group, text="(none)", text_color="gray",
                                          width=200, anchor=tk.W)
        self._limit_status.pack(side=tk.LEFT, padx=4)

        # Row 2: notch floor + action buttons
        row2 = ctk.CTkFrame(toolbar, fg_color="transparent")
        row2.pack(fill=tk.X, pady=1)

        ctk.CTkCheckBox(row2, text="Max notch depth (dB):",
                        variable=self._notch_enabled_var,
                        command=self._on_notch_toggle).pack(side=tk.LEFT)
        self._notch_entry = ctk.CTkEntry(row2, textvariable=self._notch_db_var,
                                          width=54, state=tk.DISABLED)
        self._notch_entry.pack(side=tk.LEFT, padx=(4, 16))

        self._compute_btn = ctk.CTkButton(row2, text="Compute Notch", width=130,
                                          command=self._compute_notch)
        self._compute_btn.pack(side=tk.LEFT, padx=(0, 8))
        self._export_excel_btn = ctk.CTkButton(row2, text="Export Excel…", width=110,
                                                command=self._export_excel,
                                                state=tk.DISABLED)
        self._export_excel_btn.pack(side=tk.LEFT, padx=(0, 4))
        self._export_csv_btn = ctk.CTkButton(row2, text="Export CSV…", width=90,
                                              command=self._export_csv,
                                              state=tk.DISABLED)
        self._export_csv_btn.pack(side=tk.LEFT)

        # ── Body ──────────────────────────────────────────────────────────
        body = ctk.CTkFrame(self.frame, fg_color="transparent")
        body.pack(fill=tk.BOTH, expand=True, padx=6, pady=4)

        # Left panel — scrollable
        self._left = ctk.CTkScrollableFrame(body, width=300)
        self._left.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 6))

        # Nodes section header + buttons
        ctk.CTkLabel(self._left, text="Nodes",
                     font=ctk.CTkFont(weight="bold")).pack(anchor=tk.W, pady=(4, 2))

        node_btns = ctk.CTkFrame(self._left, fg_color="transparent")
        node_btns.pack(fill=tk.X, pady=(0, 2))
        ctk.CTkButton(node_btns, text="Paste…", width=68,
                      command=self._paste_nodes).pack(side=tk.LEFT, padx=(0, 2))
        ctk.CTkButton(node_btns, text="Import CSV…", width=90,
                      command=self._import_nodes_csv).pack(side=tk.LEFT, padx=2)
        ctk.CTkButton(node_btns, text="Clear All", width=76,
                      command=self._clear_all_nodes).pack(side=tk.LEFT, padx=2)

        # Column header for node rows
        hdr = ctk.CTkFrame(self._left, fg_color="transparent")
        hdr.pack(fill=tk.X)
        ctk.CTkLabel(hdr, text="Show", width=44, anchor=tk.CENTER,
                     font=ctk.CTkFont(size=11), text_color="gray").pack(side=tk.LEFT)
        ctk.CTkLabel(hdr, text="Node", width=46, anchor=tk.CENTER,
                     font=ctk.CTkFont(size=11), text_color="gray").pack(side=tk.LEFT)
        ctk.CTkLabel(hdr, text="Dir", width=74, anchor=tk.CENTER,
                     font=ctk.CTkFont(size=11), text_color="gray").pack(side=tk.LEFT)
        ctk.CTkLabel(hdr, text="Limit", width=50, anchor=tk.CENTER,
                     font=ctk.CTkFont(size=11), text_color="gray").pack(side=tk.LEFT)

        # Node rows will be packed here (between header and view section)
        self._nodes_container = ctk.CTkFrame(self._left, fg_color="transparent")
        self._nodes_container.pack(fill=tk.X)

        # Separator
        ctk.CTkFrame(self._left, height=1, fg_color="gray40").pack(
            fill=tk.X, pady=(6, 6))

        # View section
        ctk.CTkLabel(self._left, text="View",
                     font=ctk.CTkFont(weight="bold")).pack(anchor=tk.W, pady=(0, 2))

        for val, label in [
            ("input",        "Input ASDs (orig / notched)"),
            ("response_dof", "Response — selected DOF"),
            ("response_all", "All responses overlay"),
            ("grms",         "GRMS Summary"),
        ]:
            ctk.CTkRadioButton(self._left, text=label, variable=self._view_var,
                               value=val, command=self._redraw).pack(anchor=tk.W, pady=1)

        dof_pick_row = ctk.CTkFrame(self._left, fg_color="transparent")
        dof_pick_row.pack(fill=tk.X, pady=(4, 0))
        ctk.CTkLabel(dof_pick_row, text="DOF:").pack(side=tk.LEFT)
        self._dof_picker_menu = ctk.CTkOptionMenu(
            dof_pick_row, variable=self._dof_picker_var,
            values=["—"], command=lambda _: self._redraw(), width=150)
        self._dof_picker_menu.pack(side=tk.LEFT, padx=4)

        # Right — matplotlib canvas + GRMS table overlay
        right = ctk.CTkFrame(body, fg_color="transparent")
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._fig = Figure(figsize=(8, 5))
        self._ax = self._fig.add_subplot(111)
        self._canvas = FigureCanvasTkAgg(self._fig, master=right)
        self._canvas_widget = self._canvas.get_tk_widget()
        self._canvas_widget.pack(fill=tk.BOTH, expand=True)
        NavigationToolbar2Tk(self._canvas, right).update()

        self._grms_frame = ctk.CTkFrame(right)
        self._grms_sheet = Sheet(
            self._grms_frame,
            headers=["DOF", "Orig Input GRMS (g)", "Notched Input GRMS (g)",
                     "Resp Before (g)", "Resp After (g)", "Max Notch (dB)"],
            height=300,
            show_row_index=False,
        )
        self._grms_sheet.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)
        self._grms_sheet.enable_bindings("column_width_resize")

        self._draw_idle_plot()

    # ── Help ──────────────────────────────────────────────────────────────

    def _open_help(self):
        win = ctk.CTkToplevel(self.frame.winfo_toplevel())
        win.title("Response Limiting — Guide")
        win.geometry("560x480")
        win.resizable(True, True)
        win.transient(self.frame.winfo_toplevel())
        tb = ctk.CTkTextbox(win, wrap="word")
        tb.pack(fill="both", expand=True, padx=10, pady=(10, 5))
        tb.insert("1.0", self._GUIDE_TEXT)
        tb.configure(state="disabled")
        ctk.CTkButton(win, text="Close", width=80,
                      command=win.destroy).pack(pady=(0, 10))

    # ── Notch floor toggle ────────────────────────────────────────────────

    def _on_notch_toggle(self):
        self._notch_entry.configure(
            state=tk.NORMAL if self._notch_enabled_var.get() else tk.DISABLED)

    # ── Background threading ───────────────────────────────────────────────

    def _run_in_background(self, label, work_fn, done_fn):
        self._status_label.configure(text=label, text_color="gray")
        self._open_btn.configure(state=tk.DISABLED)
        self._compute_btn.configure(state=tk.DISABLED)
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
                self._open_btn.configure(state=tk.NORMAL)
                self._compute_btn.configure(state=tk.NORMAL)
                done_fn(container.get('result'), container.get('error'))

        thread = threading.Thread(target=_worker, daemon=True)
        thread.start()
        self.frame.after(50, _poll)

    # ── OP2 loading ────────────────────────────────────────────────────────

    def _open_op2(self):
        path = filedialog.askopenfilename(
            title="Open FRF OP2",
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

            cfg = RESPONSE_TYPES['Acceleration']
            frf_dict = getattr(op2, cfg['frf_attr'], None) or {}
            if not frf_dict:
                messagebox.showwarning(
                    "No FRF Data",
                    "This OP2 has no frequency-response acceleration data.\n\n"
                    "The deck must include:\n"
                    "  ACCELERATION(PLOT,PHASE) = ALL")
                self._file_label.configure(text="(no FRF data)", text_color="orange")
                return

            self._op2 = op2
            self._op2_path = path

            sc_pairs = subcase_options(frf_dict)
            self._subcase_opts = sc_pairs
            labels = [lbl for _, lbl in sc_pairs]
            self._sc_menu.configure(values=labels)
            self._subcase_var.set(labels[0] if labels else "—")

            stem = os.path.splitext(os.path.basename(path))[0]
            self._file_label.configure(text=stem, text_color=("gray10", "gray90"))
            self._status_label.configure(
                text=f"{os.path.basename(path)} — {len(sc_pairs)} subcase(s)",
                text_color=("gray10", "gray90"))
            self._clear_results()

        self._run_in_background("Loading OP2…", _work, _done)

    def _clear_op2(self):
        self._op2 = None
        self._op2_path = None
        self._subcase_opts = []
        self._subcase_var.set("—")
        self._sc_menu.configure(values=["—"])
        self._file_label.configure(text="(none)", text_color="gray")
        self._status_label.configure(text="Load an FRF OP2 to begin", text_color="gray")
        self._clear_results()
        self._draw_idle_plot()

    def _on_subcase_change(self, _=None):
        self._clear_results()

    def _get_subcase_int(self):
        label = self._subcase_var.get()
        for sc_id, lbl in self._subcase_opts:
            if lbl == label:
                return sc_id
        return None

    # ── ASD loaders ───────────────────────────────────────────────────────

    @staticmethod
    def _asd_status_text(label, freqs):
        return f"{label} ({len(freqs)} pts, {freqs[0]:.1f}–{freqs[-1]:.1f} Hz)"

    def _load_input_asd(self):
        path = filedialog.askopenfilename(
            title="Load Input ASD",
            filetypes=[("Text/CSV", "*.txt *.csv"), ("All files", "*.*")])
        if not path:
            return
        try:
            freqs, asds = parse_asd_file(path)
        except (OSError, ValueError) as exc:
            messagebox.showerror("Load Error", str(exc))
            return
        self._input_asd_freqs, self._input_asd_vals = freqs, asds
        self._input_status.configure(
            text=self._asd_status_text(os.path.basename(path), freqs),
            text_color=("gray10", "gray90"))
        self._clear_results()

    def _paste_input_asd(self):
        text = self._paste_dialog("Paste Input ASD",
                                  "Paste 2-column data (freq  g²/Hz), one row per line:")
        if text is None:
            return
        freqs, asds = parse_asd_text(text)
        if freqs is None:
            messagebox.showerror("Parse Error", "Need at least 2 valid frequency rows.")
            return
        self._input_asd_freqs, self._input_asd_vals = freqs, asds
        self._input_status.configure(
            text=self._asd_status_text("(pasted)", freqs), text_color=("gray10", "gray90"))
        self._clear_results()

    def _clear_input_asd(self):
        self._input_asd_freqs = self._input_asd_vals = None
        self._input_status.configure(text="(none)", text_color="gray")
        self._clear_results()

    def _load_limit_asd(self):
        path = filedialog.askopenfilename(
            title="Load Response Limit ASD",
            filetypes=[("Text/CSV", "*.txt *.csv"), ("All files", "*.*")])
        if not path:
            return
        try:
            freqs, asds = parse_asd_file(path)
        except (OSError, ValueError) as exc:
            messagebox.showerror("Load Error", str(exc))
            return
        self._limit_asd_freqs, self._limit_asd_vals = freqs, asds
        self._limit_status.configure(
            text=self._asd_status_text(os.path.basename(path), freqs),
            text_color=("gray10", "gray90"))
        self._clear_results()

    def _paste_limit_asd(self):
        text = self._paste_dialog("Paste Response Limit ASD",
                                  "Paste 2-column data (freq  g²/Hz), one row per line:")
        if text is None:
            return
        freqs, asds = parse_asd_text(text)
        if freqs is None:
            messagebox.showerror("Parse Error", "Need at least 2 valid frequency rows.")
            return
        self._limit_asd_freqs, self._limit_asd_vals = freqs, asds
        self._limit_status.configure(
            text=self._asd_status_text("(pasted)", freqs), text_color=("gray10", "gray90"))
        self._clear_results()

    def _clear_limit_asd(self):
        self._limit_asd_freqs = self._limit_asd_vals = None
        self._limit_status.configure(text="(none)", text_color="gray")
        self._clear_results()

    # ── Node row management ───────────────────────────────────────────────

    def _add_node_row(self, nid, dir_name="X", show=True, limit=True):
        color = _NODE_COLORS[len(self._node_rows) % len(_NODE_COLORS)]

        show_var = ctk.BooleanVar(value=show)
        dir_var = ctk.StringVar(value=dir_name)
        limit_var = ctk.BooleanVar(value=limit)

        row_frame = ctk.CTkFrame(self._nodes_container, fg_color="transparent")
        row_frame.pack(fill=tk.X, pady=1)

        # Colour swatch
        swatch = tk.Canvas(row_frame, width=10, height=16,
                           bg=ctk.ThemeManager.theme["CTkFrame"]["fg_color"][1],
                           highlightthickness=0)
        swatch.pack(side=tk.LEFT, padx=(2, 0))
        swatch.create_rectangle(1, 3, 9, 13, fill=color, outline="")

        ctk.CTkCheckBox(row_frame, text="", variable=show_var, width=24,
                        command=self._redraw).pack(side=tk.LEFT)
        ctk.CTkLabel(row_frame, text=str(nid), width=46,
                     anchor=tk.E).pack(side=tk.LEFT)
        ctk.CTkOptionMenu(row_frame, variable=dir_var, values=_DOF_NAMES,
                          width=72, command=lambda _: self._clear_results()).pack(side=tk.LEFT, padx=2)
        ctk.CTkCheckBox(row_frame, text="Limit", variable=limit_var,
                        width=70).pack(side=tk.LEFT, padx=(4, 2))

        entry = {
            'nid': nid,
            'show_var': show_var,
            'dir_var': dir_var,
            'limit_var': limit_var,
            'color': color,
            'row_frame': row_frame,
        }
        self._node_rows.append(entry)

    def _paste_nodes(self):
        text = self._paste_dialog(
            "Paste Nodes",
            "Paste one node per line:  node_id  direction\n"
            "Direction is X, Y, or Z  (e.g.  1001 Z)\n"
            "Direction defaults to X if omitted.")
        if text is None:
            return
        added = 0
        for line in text.splitlines():
            line = line.strip()
            if not line or line.startswith('#'):
                continue
            parts = line.replace(',', ' ').split()
            try:
                nid = int(parts[0])
                dir_name = parts[1].upper() if len(parts) > 1 else "X"
                if dir_name not in _DOF_NAMES:
                    dir_name = "X"
            except (ValueError, IndexError):
                continue
            self._add_node_row(nid, dir_name)
            added += 1
        if added:
            self._status_label.configure(text=f"Added {added} node row(s).",
                                         text_color=("gray10", "gray90"))

    def _import_nodes_csv(self):
        path = filedialog.askopenfilename(
            title="Import Nodes",
            filetypes=[("Text/CSV", "*.txt *.csv"), ("All files", "*.*")])
        if not path:
            return
        try:
            with open(path, encoding='utf-8') as f:
                text = f.read()
        except OSError as exc:
            messagebox.showerror("Error", str(exc))
            return
        added = 0
        for line in text.splitlines():
            line = line.strip()
            if not line or line.startswith('#') or line.startswith('$'):
                continue
            parts = line.replace(',', ' ').split()
            try:
                nid = int(parts[0])
                dir_name = parts[1].upper() if len(parts) > 1 else "X"
                if dir_name not in _DOF_NAMES:
                    dir_name = "X"
            except (ValueError, IndexError):
                continue
            self._add_node_row(nid, dir_name)
            added += 1
        self._status_label.configure(text=f"Imported {added} node row(s).",
                                     text_color=("gray10", "gray90"))

    def _clear_all_nodes(self):
        for entry in self._node_rows:
            entry['row_frame'].destroy()
        self._node_rows.clear()
        self._clear_results()

    # ── Paste dialog helper ────────────────────────────────────────────────

    def _paste_dialog(self, title, prompt):
        result = {}
        win = ctk.CTkToplevel(self.frame.winfo_toplevel())
        win.title(title)
        win.geometry("440x320")
        win.resizable(True, True)
        win.transient(self.frame.winfo_toplevel())
        win.grab_set()

        ctk.CTkLabel(win, text=prompt, anchor=tk.W).pack(fill=tk.X, padx=10, pady=(10, 4))
        tb = ctk.CTkTextbox(win, wrap="none")
        tb.pack(fill=tk.BOTH, expand=True, padx=10, pady=4)

        btn_row = ctk.CTkFrame(win, fg_color="transparent")
        btn_row.pack(fill=tk.X, padx=10, pady=(4, 10))

        def _ok():
            result['text'] = tb.get("1.0", tk.END)
            win.destroy()

        ctk.CTkButton(btn_row, text="OK", width=80, command=_ok).pack(side=tk.LEFT, padx=4)
        ctk.CTkButton(btn_row, text="Cancel", width=80,
                      command=win.destroy).pack(side=tk.LEFT, padx=4)
        win.wait_window()
        return result.get('text')

    # ── Core computation ──────────────────────────────────────────────────

    def _compute_notch(self):
        if self._op2 is None:
            messagebox.showwarning("No OP2", "Load an FRF OP2 first.")
            return
        if self._input_asd_freqs is None:
            messagebox.showwarning("No Input ASD", "Load an Input ASD first.")
            return
        if self._limit_asd_freqs is None:
            messagebox.showwarning("No Limit ASD", "Load a Response Limit ASD first.")
            return

        plot_rows = [r for r in self._node_rows if r['show_var'].get()]
        limit_rows = [r for r in self._node_rows if r['limit_var'].get()]
        if not plot_rows and not limit_rows:
            messagebox.showwarning("No Nodes", "Add at least one node row.")
            return
        if not limit_rows:
            messagebox.showwarning("No Limit Nodes",
                                   "Check 'Limit' on at least one node row.")
            return

        sc = self._get_subcase_int()
        if sc is None:
            messagebox.showwarning("No Subcase", "Select a valid subcase.")
            return

        cfg = RESPONSE_TYPES['Acceleration']
        unit_factor = cfg['unit_factors'].get(self._units_var.get(), 386.089)
        id_attr = cfg['id_attr']

        frf_dict = getattr(self._op2, cfg['frf_attr'], None) or {}
        frf_tbl = lookup_subcase(frf_dict, sc)
        if frf_tbl is None:
            messagebox.showerror("Error", f"Subcase {sc} not found in FRF results.")
            return

        freqs = frf_tbl._times
        arr = getattr(frf_tbl, id_attr)
        entity_ids = arr[:, 0] if id_attr == 'node_gridtype' else arr

        orig_interp = interp_loglog(self._input_asd_freqs, self._input_asd_vals, freqs)
        limit_interp = interp_loglog(self._limit_asd_freqs, self._limit_asd_vals, freqs)

        # Compute minimum allowed input across all limit DOFs
        min_allowed = np.full(len(freqs), np.inf)
        missing = []
        limit_dofs_set = set()

        for r in limit_rows:
            nid = r['nid']
            dof_idx = _DOF_NAMES.index(r['dir_var'].get())
            hits = np.where(entity_ids == nid)[0]
            if not len(hits):
                missing.append(f"{nid} {_DOF_NAMES[dof_idx]}")
                continue
            H_g = frf_tbl.data[:, hits[0], dof_idx] / unit_factor
            H_mag2 = H_g.real**2 + H_g.imag**2
            with np.errstate(divide='ignore', invalid='ignore'):
                allowed = np.where(H_mag2 > 0, limit_interp / H_mag2, np.inf)
            min_allowed = np.minimum(min_allowed, allowed)
            limit_dofs_set.add((nid, dof_idx))

        notched = np.minimum(orig_interp, min_allowed)
        notched = np.where(np.isinf(notched), orig_interp, notched)

        if self._notch_enabled_var.get():
            try:
                db = float(self._notch_db_var.get())
            except ValueError:
                db = 6.0
            notched = np.maximum(notched, orig_interp * 10**(-db / 10.0))

        # Compute response curves for all show rows
        response_curves = {}
        for r in plot_rows:
            nid = r['nid']
            dof_idx = _DOF_NAMES.index(r['dir_var'].get())
            hits = np.where(entity_ids == nid)[0]
            if not len(hits):
                if f"{nid} {_DOF_NAMES[dof_idx]}" not in missing:
                    missing.append(f"{nid} {_DOF_NAMES[dof_idx]}")
                continue
            H_g = frf_tbl.data[:, hits[0], dof_idx] / unit_factor
            H_mag2 = H_g.real**2 + H_g.imag**2
            response_curves[(nid, dof_idx)] = (H_mag2 * orig_interp, H_mag2 * notched)

        self._frf_freqs = freqs
        self._orig_asd_interp = orig_interp
        self._limit_asd_interp = limit_interp
        self._notched_asd = notched
        self._response_curves = response_curves
        self._limit_dofs_set = limit_dofs_set

        # Update DOF picker with show-row labels
        labels = []
        for r in plot_rows:
            nid = r['nid']
            dof_idx = _DOF_NAMES.index(r['dir_var'].get())
            if (nid, dof_idx) in response_curves:
                tag = " ★" if (nid, dof_idx) in limit_dofs_set else ""
                labels.append(f"Node {nid} {_DOF_NAMES[dof_idx]}{tag}")
        self._dof_picker_menu.configure(values=labels if labels else ["—"])
        if labels:
            self._dof_picker_var.set(labels[0])

        self._export_excel_btn.configure(state=tk.NORMAL)
        self._export_csv_btn.configure(state=tk.NORMAL)

        parts = [f"Notch computed — {len(response_curves)} curve(s)."]
        if missing:
            parts.append(f"Not found: {', '.join(missing[:4])}")
        self._status_label.configure(text="  ".join(parts),
                                     text_color=("gray10", "gray90"))
        self._redraw()

    def _clear_results(self):
        self._frf_freqs = None
        self._orig_asd_interp = None
        self._limit_asd_interp = None
        self._notched_asd = None
        self._response_curves = {}
        self._limit_dofs_set = set()
        self._export_excel_btn.configure(state=tk.DISABLED)
        self._export_csv_btn.configure(state=tk.DISABLED)
        self._dof_picker_menu.configure(values=["—"])
        self._dof_picker_var.set("—")

    # ── Plot helpers ──────────────────────────────────────────────────────

    def _get_theme(self):
        return _THEMES.get(ctk.get_appearance_mode().lower(), _THEMES["light"])

    def _format_plot(self, ax, title="", ylabel="ASD (g²/Hz)"):
        th = self._get_theme()
        self._fig.set_facecolor(th["fig_bg"])
        ax.set_facecolor(th["plot_bg"])
        ax.tick_params(colors=th["text"])
        ax.xaxis.label.set_color(th["text"])
        ax.yaxis.label.set_color(th["text"])
        ax.title.set_color(th["text"])
        for spine in ax.spines.values():
            spine.set_edgecolor(th["spine"])
        ax.grid(True, which="both", color=th["grid"], linewidth=0.5)
        ax.set_xlabel("Frequency (Hz)")
        ax.set_ylabel(ylabel)
        if title:
            ax.set_title(title)

    def _draw_idle_plot(self):
        th = self._get_theme()
        self._fig.clf()
        ax = self._fig.add_subplot(111)
        self._ax = ax
        self._format_plot(ax, title="Response Limiting")
        ax.text(0.5, 0.5,
                "1. Open FRF OP2\n2. Load Input ASD + Response Limit\n"
                "3. Paste node rows and check Limit\n4. Compute Notch",
                ha='center', va='center', transform=ax.transAxes,
                color=th["text"], fontsize=11)
        self._canvas.draw_idle()

    # ── View dispatch ──────────────────────────────────────────────────────

    def _redraw(self):
        if self._notched_asd is None:
            return
        view = self._view_var.get()
        if view == "grms":
            self._canvas_widget.pack_forget()
            self._grms_frame.pack(fill=tk.BOTH, expand=True)
            self._draw_grms_view()
        else:
            self._grms_frame.pack_forget()
            self._canvas_widget.pack(fill=tk.BOTH, expand=True)
            self._fig.clf()
            self._ax = self._fig.add_subplot(111)
            if view == "input":
                self._draw_input_view()
            elif view == "response_dof":
                self._draw_response_dof_view()
            elif view == "response_all":
                self._draw_response_all_view()
            self._canvas.draw_idle()

    def _draw_input_view(self):
        ax = self._ax
        freqs = self._frf_freqs
        if self._input_asd_freqs is not None:
            ax.loglog(self._input_asd_freqs, self._input_asd_vals,
                      color="#1f77b4", label="Original Input")
        ax.loglog(freqs, self._notched_asd,
                  color="#d62728", label="Notched Input")
        orig_grms = np.sqrt(max(0.0, grms_loglog(self._input_asd_freqs,
                                                   self._input_asd_vals))) \
            if self._input_asd_freqs is not None else 0.0
        notch_grms = np.sqrt(max(0.0, grms_loglog(freqs, self._notched_asd)))
        ax.legend(title=f"Orig:   {orig_grms:.3g} g GRMS\nNotched: {notch_grms:.3g} g GRMS")
        self._format_plot(ax, title="Input ASD — Original vs Notched")
        ax.set_xlim(freqs[0], freqs[-1])

    def _draw_response_dof_view(self):
        ax = self._ax
        freqs = self._frf_freqs
        sel = self._dof_picker_var.get()

        # Resolve selected key from picker label
        selected_key = None
        for nid, dof_idx in self._response_curves:
            tag = " ★" if (nid, dof_idx) in self._limit_dofs_set else ""
            if f"Node {nid} {_DOF_NAMES[dof_idx]}{tag}" == sel:
                selected_key = (nid, dof_idx)
                break
        if selected_key is None and self._response_curves:
            selected_key = next(iter(self._response_curves))

        if selected_key is not None:
            rb, ra = self._response_curves[selected_key]
            ax.loglog(freqs, rb, color="#1f77b4", label="Before notch")
            ax.loglog(freqs, ra, color="#d62728", label="After notch")
            rb_grms = np.sqrt(max(0.0, grms_loglog(freqs, rb)))
            ra_grms = np.sqrt(max(0.0, grms_loglog(freqs, ra)))

        ax.loglog(freqs, self._limit_asd_interp,
                  color="#ff7f0e", linestyle="--", linewidth=1.5, label="Response Limit")

        if selected_key is not None:
            ax.legend(title=f"Before: {rb_grms:.3g} g\nAfter:  {ra_grms:.3g} g")
        else:
            ax.legend()

        nid, dof_idx = selected_key if selected_key else (0, 0)
        self._format_plot(ax, title=f"Response: Node {nid} {_DOF_NAMES[dof_idx]}")
        ax.set_xlim(freqs[0], freqs[-1])

    def _draw_response_all_view(self):
        ax = self._ax
        freqs = self._frf_freqs

        # Colour map from node_rows so colours match the swatch
        color_map = {
            (r['nid'], _DOF_NAMES.index(r['dir_var'].get())): r['color']
            for r in self._node_rows if r['show_var'].get()
        }
        for (nid, dof_idx), (_, resp_after) in self._response_curves.items():
            col = color_map.get((nid, dof_idx), "#aaaaaa")
            tag = " ★" if (nid, dof_idx) in self._limit_dofs_set else ""
            ax.loglog(freqs, resp_after, color=col,
                      label=f"Node {nid} {_DOF_NAMES[dof_idx]}{tag}")

        ax.loglog(freqs, self._limit_asd_interp,
                  color="#ff7f0e", linestyle="--", linewidth=2, label="Response Limit")
        ax.legend(fontsize=8)
        self._format_plot(ax, title="All Responses (notched) vs Limit — ★ = limit node")
        ax.set_xlim(freqs[0], freqs[-1])

    def _draw_grms_view(self):
        freqs = self._frf_freqs
        orig_grms = np.sqrt(max(0.0, grms_loglog(freqs, self._orig_asd_interp)))
        notch_grms = np.sqrt(max(0.0, grms_loglog(freqs, self._notched_asd)))
        with np.errstate(divide='ignore', invalid='ignore'):
            ratio = np.where(self._notched_asd > 0,
                             self._orig_asd_interp / self._notched_asd, 1.0)
        max_notch_db = 10.0 * np.log10(max(float(ratio.max()), 1.0))

        rows = []
        for (nid, dof_idx), (rb, ra) in self._response_curves.items():
            tag = " ★" if (nid, dof_idx) in self._limit_dofs_set else ""
            rows.append([
                f"Node {nid} {_DOF_NAMES[dof_idx]}{tag}",
                f"{orig_grms:.4g}",
                f"{notch_grms:.4g}",
                f"{np.sqrt(max(0.0, grms_loglog(freqs, rb))):.4g}",
                f"{np.sqrt(max(0.0, grms_loglog(freqs, ra))):.4g}",
                f"{max_notch_db:.2f}",
            ])
        self._grms_sheet.set_sheet_data(rows)

    # ── Export ────────────────────────────────────────────────────────────

    def _export_excel(self):
        if self._notched_asd is None:
            messagebox.showwarning("No Data", "Compute notch first.")
            return
        path = filedialog.asksaveasfilename(
            title="Save Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("All files", "*.*")])
        if not path:
            return
        try:
            import openpyxl
        except ImportError:
            messagebox.showerror("Missing Library", "openpyxl is required for Excel export.")
            return

        freqs = self._frf_freqs
        wb = openpyxl.Workbook()

        ws = wb.active
        ws.title = "Notched ASD"
        ws.append(["Freq (Hz)", "Original ASD (g²/Hz)", "Notched ASD (g²/Hz)"])
        for f, orig, notch in zip(freqs, self._orig_asd_interp, self._notched_asd):
            ws.append([float(f), float(orig), float(notch)])

        ws2 = wb.create_sheet("Limit ASD")
        ws2.append(["Freq (Hz)", "Limit ASD (g²/Hz)"])
        for f, lim in zip(freqs, self._limit_asd_interp):
            ws2.append([float(f), float(lim)])

        for (nid, dof_idx), (rb, ra) in self._response_curves.items():
            sname = f"N{nid}_{_DOF_NAMES[dof_idx]}"[:31]
            ws_d = wb.create_sheet(sname)
            ws_d.append(["Freq (Hz)", "Resp Before (g²/Hz)", "Resp After (g²/Hz)"])
            for f, r_b, r_a in zip(freqs, rb, ra):
                ws_d.append([float(f), float(r_b), float(r_a)])

        ws_g = wb.create_sheet("GRMS Summary")
        ws_g.append(["DOF", "Orig Input GRMS (g)", "Notched Input GRMS (g)",
                      "Resp Before (g)", "Resp After (g)", "Max Notch (dB)"])
        orig_grms = np.sqrt(max(0.0, grms_loglog(freqs, self._orig_asd_interp)))
        notch_grms = np.sqrt(max(0.0, grms_loglog(freqs, self._notched_asd)))
        with np.errstate(divide='ignore', invalid='ignore'):
            ratio = np.where(self._notched_asd > 0,
                             self._orig_asd_interp / self._notched_asd, 1.0)
        max_notch_db = 10.0 * np.log10(max(float(ratio.max()), 1.0))
        for (nid, dof_idx), (rb, ra) in self._response_curves.items():
            tag = " ★" if (nid, dof_idx) in self._limit_dofs_set else ""
            ws_g.append([
                f"Node {nid} {_DOF_NAMES[dof_idx]}{tag}",
                round(orig_grms, 6), round(notch_grms, 6),
                round(np.sqrt(max(0.0, grms_loglog(freqs, rb))), 6),
                round(np.sqrt(max(0.0, grms_loglog(freqs, ra))), 6),
                round(max_notch_db, 4),
            ])

        wb.save(path)
        self._status_label.configure(text=f"Saved: {os.path.basename(path)}",
                                     text_color=("gray10", "gray90"))

    def _export_csv(self):
        if self._notched_asd is None:
            messagebox.showwarning("No Data", "Compute notch first.")
            return
        path = filedialog.asksaveasfilename(
            title="Save CSV",
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv"), ("All files", "*.*")])
        if not path:
            return
        import csv as _csv
        with open(path, 'w', newline='', encoding='utf-8') as f:
            writer = _csv.writer(f)
            writer.writerow(["Freq (Hz)", "Notched ASD (g²/Hz)"])
            for freq, asd in zip(self._frf_freqs, self._notched_asd):
                writer.writerow([float(freq), float(asd)])
        self._status_label.configure(text=f"Saved: {os.path.basename(path)}",
                                     text_color=("gray10", "gray90"))
