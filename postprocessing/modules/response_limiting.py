"""FRF Response Limiting: compute a notched drive ASD from FRF data.

Loads a Nastran SOL 111 frequency-response OP2, an input environment ASD,
and a response limit ASD.  Node/direction rows are added to a grid; rows
with X/Y/Z set to 'S' (Show) or 'L' (Limit) are plotted, rows with 'L'
drive the notch calculation.  Results update live as inputs change.
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
_DOF_COLS  = [2, 3, 4]   # column indices in the node grid

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

# Cell colours for S / L states
_COL_S  = {"bg": "#1f6aa5", "fg": "white"}   # show only
_COL_L  = {"bg": "#cc6600", "fg": "white"}   # limit (implies show)
_VALID  = {"", "S", "L"}


class ResponseLimitingModule:
    name = "Response Limiting"

    _GUIDE_TEXT = """\
Response Limiting: Quick Guide

PURPOSE
  Compute a notched drive ASD so that the response at selected nodes/DOFs
  stays at or below a specified limit spectrum.

WORKFLOW
  1. Open a SOL 111 FRF OP2 (ACCELERATION(PLOT,PHASE)=ALL required).
  2. Load or paste the Input ASD (freq vs g²/Hz).
  3. Load or paste the Response Limit ASD (freq vs g²/Hz).
     The Response Limit Paste button pre-fills with existing data for editing.
  4. Add nodes via Paste…, Import CSV…, or Add… and set X/Y/Z states.
     Type or use Tab/Enter to cycle states in the grid:
       S = Show:  response curve is plotted
       L = Limit: drives the notch AND is plotted (marked [L])
       (empty)  = not used
  5. Results update live; use Export to save.

NODE GRID
  Add / import format, one per line:  node_id  [label]
    e.g.  1001   or   1001 Tip mass   or   1001, Tip mass
    Nodes arrive with all directions set to L.
    Edit individual X / Y / Z cells in the grid to adjust.

PLOT THEME
  ☾ / ☀ button toggles plot background independent of the system theme.

CURVE PICKER  (Response: single curve view)
  Selects which node/direction to display. [L] marks Limit DOFs.

UNITS
  Set to match OP2 output units (in/s² for slinch/inch models).
  All ASDs are displayed in g²/Hz; GRMS in g.
"""

    def __init__(self, parent):
        self.frame = ctk.CTkFrame(parent)

        # ── State ─────────────────────────────────────────────────────────
        self._op2 = None
        self._op2_path = None
        self._subcase_opts = []
        self._subcase_var = ctk.StringVar(value="(none)")
        self._units_var = ctk.StringVar(value="in/s²")

        self._input_asd_freqs = None
        self._input_asd_vals  = None
        self._limit_asd_freqs = None
        self._limit_asd_vals  = None

        self._plot_theme = "light"   # "light" or "dark"

        # Computed results (on FRF frequency grid)
        self._frf_freqs       = None
        self._orig_asd_interp = None
        self._limit_asd_interp = None
        self._notched_asd     = None
        self._response_curves  = {}   # {(nid, dof_idx): (resp_before, resp_after)}
        self._limit_dofs_set   = set()

        self._notch_enabled_var = ctk.BooleanVar(value=False)
        self._notch_db_var      = ctk.StringVar(value="6.0")

        self._view_var       = ctk.StringVar(value="input")
        self._curve_var      = ctk.StringVar(value="(none)")   # single-curve picker

        self._debounce_id    = None

        self._build_ui()

    # ── UI construction ────────────────────────────────────────────────────

    def _build_ui(self):
        toolbar = ctk.CTkFrame(self.frame, fg_color="transparent")
        toolbar.pack(fill=tk.X, padx=6, pady=(6, 2))

        # Row 0: OP2 + units + subcase + status + help
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
        ctk.CTkOptionMenu(row0, variable=self._units_var,
                          values=["in/s²", "m/s²"], width=90,
                          command=lambda _: self._schedule_recompute()).pack(side=tk.LEFT)

        ctk.CTkLabel(row0, text="Subcase:").pack(side=tk.LEFT, padx=(8, 2))
        self._sc_menu = ctk.CTkOptionMenu(row0, variable=self._subcase_var,
                                          values=["(none)"],
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

        inp = ctk.CTkFrame(row1, fg_color="transparent")
        inp.pack(side=tk.LEFT, padx=(0, 16))
        ctk.CTkLabel(inp, text="Input ASD:", width=72, anchor=tk.W).pack(side=tk.LEFT)
        ctk.CTkButton(inp, text="Load…",  width=60, command=self._load_input_asd).pack(side=tk.LEFT, padx=2)
        ctk.CTkButton(inp, text="Paste…", width=60, command=self._paste_input_asd).pack(side=tk.LEFT, padx=2)
        ctk.CTkButton(inp, text="Clear",  width=50, command=self._clear_input_asd).pack(side=tk.LEFT, padx=2)
        self._input_status = ctk.CTkLabel(inp, text="(none)", text_color="gray",
                                          width=200, anchor=tk.W)
        self._input_status.pack(side=tk.LEFT, padx=4)

        lim = ctk.CTkFrame(row1, fg_color="transparent")
        lim.pack(side=tk.LEFT)
        ctk.CTkLabel(lim, text="Response Limit:", width=104, anchor=tk.W).pack(side=tk.LEFT)
        ctk.CTkButton(lim, text="Load…",  width=60, command=self._load_limit_asd).pack(side=tk.LEFT, padx=2)
        ctk.CTkButton(lim, text="Paste…", width=60, command=self._paste_limit_asd).pack(side=tk.LEFT, padx=2)
        ctk.CTkButton(lim, text="Clear",  width=50, command=self._clear_limit_asd).pack(side=tk.LEFT, padx=2)
        self._limit_status = ctk.CTkLabel(lim, text="(none)", text_color="gray",
                                          width=200, anchor=tk.W)
        self._limit_status.pack(side=tk.LEFT, padx=4)

        # Row 2: notch floor
        row2 = ctk.CTkFrame(toolbar, fg_color="transparent")
        row2.pack(fill=tk.X, pady=1)

        ctk.CTkCheckBox(row2, text="Max notch depth (dB):",
                        variable=self._notch_enabled_var,
                        command=self._on_notch_toggle).pack(side=tk.LEFT)
        self._notch_entry = ctk.CTkEntry(row2, textvariable=self._notch_db_var,
                                          width=54, state=tk.DISABLED)
        self._notch_entry.pack(side=tk.LEFT, padx=(4, 16))
        self._notch_db_var.trace_add("write", lambda *_: self._schedule_recompute())

        self._export_excel_btn = ctk.CTkButton(row2, text="Export Excel…", width=110,
                                                command=self._export_excel, state=tk.DISABLED)
        self._export_excel_btn.pack(side=tk.LEFT, padx=(0, 4))
        self._export_csv_btn = ctk.CTkButton(row2, text="Export CSV…", width=90,
                                              command=self._export_csv, state=tk.DISABLED)
        self._export_csv_btn.pack(side=tk.LEFT)

        # ── Body ──────────────────────────────────────────────────────────
        body = ctk.CTkFrame(self.frame, fg_color="transparent")
        body.pack(fill=tk.BOTH, expand=True, padx=6, pady=4)

        # Left panel
        self._left = ctk.CTkScrollableFrame(body, width=320)
        self._left.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 6))

        # Node grid section
        ctk.CTkLabel(self._left, text="Nodes  (S = Show, L = Limit)",
                     font=ctk.CTkFont(weight="bold")).pack(anchor=tk.W, pady=(4, 2))

        node_btns = ctk.CTkFrame(self._left, fg_color="transparent")
        node_btns.pack(fill=tk.X, pady=(0, 4))
        ctk.CTkButton(node_btns, text="Add…",       width=62,
                      command=self._add_node_dialog).pack(side=tk.LEFT, padx=(0, 2))
        ctk.CTkButton(node_btns, text="Import CSV…",width=90,
                      command=self._import_nodes_csv).pack(side=tk.LEFT, padx=2)
        ctk.CTkButton(node_btns, text="Remove",     width=68,
                      command=self._remove_selected_nodes).pack(side=tk.LEFT, padx=2)
        ctk.CTkButton(node_btns, text="Clear All",  width=72,
                      command=self._clear_all_nodes).pack(side=tk.LEFT, padx=2)

        self._node_sheet = Sheet(
            self._left,
            headers=["Node", "Label", "X", "Y", "Z"],
            height=200,
            show_row_index=False,
            theme="dark" if ctk.get_appearance_mode() == "Dark" else "light",
        )
        self._node_sheet.pack(fill=tk.X, pady=(0, 6))
        self._node_sheet.enable_bindings(
            "single_select", "row_select", "edit_cell", "column_width_resize")
        self._node_sheet.column_width(0, 56)
        self._node_sheet.column_width(1, 100)
        self._node_sheet.column_width(2, 40)
        self._node_sheet.column_width(3, 40)
        self._node_sheet.column_width(4, 40)
        self._node_sheet.extra_bindings([("end_edit_cell", self._on_grid_edit)])

        # Separator
        ctk.CTkFrame(self._left, height=1, fg_color="gray40").pack(
            fill=tk.X, pady=(2, 6))

        # View section
        ctk.CTkLabel(self._left, text="View",
                     font=ctk.CTkFont(weight="bold")).pack(anchor=tk.W, pady=(0, 2))

        for val, label in [
            ("input",        "Input ASDs (orig / notched)"),
            ("response_dof", "Response: single curve"),
            ("response_all", "All responses overlay"),
            ("grms",         "GRMS Summary"),
        ]:
            ctk.CTkRadioButton(self._left, text=label, variable=self._view_var,
                               value=val, command=self._redraw).pack(anchor=tk.W, pady=1)

        curve_row = ctk.CTkFrame(self._left, fg_color="transparent")
        curve_row.pack(fill=tk.X, pady=(4, 0))
        ctk.CTkLabel(curve_row, text="Curve:").pack(side=tk.LEFT)
        self._curve_menu = ctk.CTkOptionMenu(
            curve_row, variable=self._curve_var,
            values=["(none)"], command=lambda _: self._redraw(), width=150)
        self._curve_menu.pack(side=tk.LEFT, padx=4)

        # Right: canvas + theme button + GRMS overlay
        right = ctk.CTkFrame(body, fg_color="transparent")
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Plot header with theme toggle
        plot_hdr = ctk.CTkFrame(right, fg_color="transparent")
        plot_hdr.pack(fill=tk.X)
        self._theme_btn = ctk.CTkButton(
            plot_hdr, text="☾ Dark", width=80,
            command=self._toggle_theme)
        self._theme_btn.pack(side=tk.RIGHT, padx=4, pady=2)

        self._fig = Figure(figsize=(8, 5))
        self._ax  = self._fig.add_subplot(111)
        self._canvas = FigureCanvasTkAgg(self._fig, master=right)
        self._canvas_widget = self._canvas.get_tk_widget()
        self._canvas_widget.pack(fill=tk.BOTH, expand=True)
        NavigationToolbar2Tk(self._canvas, right).update()

        self._grms_frame = ctk.CTkFrame(right)
        self._grms_sheet = Sheet(
            self._grms_frame,
            headers=["DOF", "Orig Input GRMS (g)", "Notched Input GRMS (g)",
                     "Resp Before (g)", "Resp After (g)", "Max Notch (dB)"],
            height=300, show_row_index=False,
        )
        self._grms_sheet.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)
        self._grms_sheet.enable_bindings("column_width_resize")

        self._draw_idle_plot()

    # ── Theme toggle ──────────────────────────────────────────────────────

    def _toggle_theme(self):
        self._plot_theme = "light" if self._plot_theme == "dark" else "dark"
        self._theme_btn.configure(
            text="☾ Dark" if self._plot_theme == "light" else "☀ Light")
        self._redraw()

    # ── Help ──────────────────────────────────────────────────────────────

    def _open_help(self):
        win = ctk.CTkToplevel(self.frame.winfo_toplevel())
        win.title("Response Limiting: Guide")
        win.geometry("560x500")
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
        self._schedule_recompute()

    # ── Live compute ──────────────────────────────────────────────────────

    def _schedule_recompute(self, delay=400):
        if self._debounce_id is not None:
            self.frame.after_cancel(self._debounce_id)
        self._debounce_id = self.frame.after(delay, self._auto_compute)

    def _auto_compute(self):
        self._debounce_id = None
        if (self._op2 is not None
                and self._input_asd_freqs is not None
                and self._limit_asd_freqs is not None
                and self._any_limit_dofs()):
            self._compute_notch(silent=True)

    def _any_limit_dofs(self):
        data = self._node_sheet.get_sheet_data()
        for row in data:
            for col in _DOF_COLS:
                if str(row[col]).upper() == "L":
                    return True
        return False

    # ── Background threading ───────────────────────────────────────────────

    def _run_in_background(self, label, work_fn, done_fn):
        self._status_label.configure(text=label, text_color="gray")
        self._open_btn.configure(state=tk.DISABLED)
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
                    "The deck must include:\n  ACCELERATION(PLOT,PHASE) = ALL")
                self._file_label.configure(text="(no FRF data)", text_color="orange")
                return

            self._op2 = op2
            self._op2_path = path

            sc_pairs = subcase_options(frf_dict)
            self._subcase_opts = sc_pairs
            labels = [lbl for _, lbl in sc_pairs]
            self._sc_menu.configure(values=labels)
            self._subcase_var.set(labels[0] if labels else "(none)")

            stem = os.path.splitext(os.path.basename(path))[0]
            self._file_label.configure(text=stem, text_color=("gray10", "gray90"))
            self._status_label.configure(
                text=f"{os.path.basename(path)}: {len(sc_pairs)} subcase(s)",
                text_color=("gray10", "gray90"))
            self._clear_results()
            self._schedule_recompute()

        self._run_in_background("Loading OP2…", _work, _done)

    def _clear_op2(self):
        self._op2 = None
        self._op2_path = None
        self._subcase_opts = []
        self._subcase_var.set("(none)")
        self._sc_menu.configure(values=["(none)"])
        self._file_label.configure(text="(none)", text_color="gray")
        self._status_label.configure(text="Load an FRF OP2 to begin", text_color="gray")
        self._clear_results()
        self._draw_idle_plot()

    def _on_subcase_change(self, _=None):
        self._clear_results()
        self._schedule_recompute()

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
        self._schedule_recompute()

    def _paste_input_asd(self):
        text = self._paste_dialog(
            "Paste Input ASD", "Paste 2-column data (freq  g²/Hz), one row per line:")
        if text is None:
            return
        freqs, asds = parse_asd_text(text)
        if freqs is None:
            messagebox.showerror("Parse Error", "Need at least 2 valid frequency rows.")
            return
        self._input_asd_freqs, self._input_asd_vals = freqs, asds
        self._input_status.configure(
            text=self._asd_status_text("(pasted)", freqs),
            text_color=("gray10", "gray90"))
        self._clear_results()
        self._schedule_recompute()

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
        self._schedule_recompute()

    def _paste_limit_asd(self):
        # Pre-populate with current limit data so the user can edit in place
        initial = ""
        if self._limit_asd_freqs is not None:
            lines = [f"{f:.5g}  {a:.6g}"
                     for f, a in zip(self._limit_asd_freqs, self._limit_asd_vals)]
            initial = "\n".join(lines)
        text = self._paste_dialog(
            "Paste / Edit Response Limit ASD",
            "2-column data (freq  g²/Hz), one row per line:",
            initial=initial)
        if text is None:
            return
        freqs, asds = parse_asd_text(text)
        if freqs is None:
            messagebox.showerror("Parse Error", "Need at least 2 valid frequency rows.")
            return
        self._limit_asd_freqs, self._limit_asd_vals = freqs, asds
        self._limit_status.configure(
            text=self._asd_status_text("(pasted)", freqs),
            text_color=("gray10", "gray90"))
        self._clear_results()
        self._schedule_recompute()

    def _clear_limit_asd(self):
        self._limit_asd_freqs = self._limit_asd_vals = None
        self._limit_status.configure(text="(none)", text_color="gray")
        self._clear_results()

    # ── Node grid management ──────────────────────────────────────────────

    def _add_node_to_grid(self, nid, label="", x="", y="", z=""):
        """Insert a row in the node sheet. x/y/z should be '', 'S', or 'L'."""
        nrows = len(self._node_sheet.get_sheet_data())
        self._node_sheet.insert_row(row=[str(nid), label,
                                        x.upper(), y.upper(), z.upper()])
        for ci, val in zip(_DOF_COLS, [x, y, z]):
            self._apply_cell_color(nrows, ci, val.upper())

    def _apply_cell_color(self, row, col, val):
        if val == "S":
            self._node_sheet.highlight_cells(
                row=row, column=col,
                bg=_COL_S["bg"], fg=_COL_S["fg"])
        elif val == "L":
            self._node_sheet.highlight_cells(
                row=row, column=col,
                bg=_COL_L["bg"], fg=_COL_L["fg"])
        else:
            self._node_sheet.dehighlight_cells(row=row, column=col)

    def _on_grid_edit(self, event):
        r, c = event.row, event.column
        if c not in _DOF_COLS:
            # Label or Node ID changed; just schedule recompute
            self._schedule_recompute()
            return
        raw = str(self._node_sheet.get_cell_data(r, c)).strip().upper()
        if raw not in _VALID:
            raw = ""
        self._node_sheet.set_cell_data(r, c, raw)
        self._apply_cell_color(r, c, raw)
        self._schedule_recompute()

    def _add_node_dialog(self):
        dlg = ctk.CTkToplevel(self.frame.winfo_toplevel())
        dlg.title("Add Nodes")
        dlg.geometry("380x310")
        dlg.transient(self.frame.winfo_toplevel())
        dlg.grab_set()

        ctk.CTkLabel(
            dlg,
            text="Enter one node per line.  Optional label:\n"
                 "  1001        1001 Tip mass        1001, Tip mass",
            justify=tk.LEFT, anchor=tk.W,
        ).pack(padx=12, pady=(12, 4), fill=tk.X)

        tb = ctk.CTkTextbox(dlg, wrap="none")
        tb.pack(fill=tk.BOTH, expand=True, padx=12)

        def _ok():
            self._parse_and_add_nodes(tb.get("1.0", "end"))
            dlg.destroy()

        btn_row = ctk.CTkFrame(dlg, fg_color="transparent")
        btn_row.pack(fill=tk.X, padx=12, pady=8)
        ctk.CTkButton(btn_row, text="Add", command=_ok).pack(side=tk.LEFT)
        ctk.CTkButton(btn_row, text="Cancel", command=dlg.destroy).pack(side=tk.LEFT, padx=5)
        dlg.bind("<Return>", lambda _: _ok())

    def _parse_and_add_nodes(self, text):
        existing = set()
        for r in self._node_sheet.get_sheet_data():
            try:
                existing.add(int(str(r[0]).strip()))
            except (ValueError, TypeError):
                pass
        added = 0
        for raw in text.splitlines():
            line = raw.strip()
            if not line or line.startswith('#') or line.startswith('$'):
                continue
            parts = [p.strip() for p in line.split(',', 1)] if ',' in line \
                else line.split(None, 1)
            try:
                nid = int(parts[0])
            except (ValueError, IndexError):
                continue
            if nid in existing:
                continue
            existing.add(nid)
            label = parts[1].strip() if len(parts) > 1 and parts[1].strip() else ""
            self._add_node_to_grid(nid, label, x="L", y="L", z="L")
            added += 1
        if added:
            self._status_label.configure(text=f"Added {added} node(s).",
                                         text_color=("gray10", "gray90"))
            self._schedule_recompute()

    def _import_nodes_csv(self):
        path = filedialog.askopenfilename(
            title="Import Nodes",
            filetypes=[("Text files", "*.txt"), ("CSV files", "*.csv"),
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
        import csv as _csv
        with open(path, newline='', encoding='utf-8-sig') as f:
            sample = f.read(4096)
            f.seek(0)
            try:
                dialect = _csv.Sniffer().sniff(sample)
                has_header = _csv.Sniffer().has_header(sample)
            except _csv.Error:
                dialect = _csv.excel
                has_header = False
            reader = _csv.reader(f, dialect)
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
                label = row[1].strip() if len(row) > 1 and str(row[1]).strip() else ""
                lines.append(f"{gid},{label}" if label else str(gid))
        return "\n".join(lines)

    @staticmethod
    def _read_text_node_file(path):
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
                label = parts[1].strip() if len(parts) > 1 and parts[1].strip() else ""
                lines.append(f"{gid},{label}" if label else str(gid))
        return "\n".join(lines)

    def _remove_selected_nodes(self):
        selected = self._node_sheet.get_selected_rows()
        if not selected:
            return
        for r in sorted(selected, reverse=True):
            self._node_sheet.delete_rows([r])
        self._schedule_recompute()

    def _clear_all_nodes(self):
        n = len(self._node_sheet.get_sheet_data())
        if n:
            self._node_sheet.delete_rows(list(range(n)))
        self._clear_results()

    def _get_grid_dofs(self):
        """Return (plot_dofs, limit_dofs, color_map, label_map) from current grid."""
        data     = self._node_sheet.get_sheet_data()
        plot_dofs  = []   # [(nid, dof_idx), ...]
        limit_dofs = []
        color_map  = {}   # (nid, dof_idx) → hex color
        label_map  = {}   # (nid, dof_idx) → display label
        for row_idx, row in enumerate(data):
            try:
                nid = int(str(row[0]).strip())
            except (ValueError, TypeError):
                continue
            user_label = str(row[1]).strip()
            color = _NODE_COLORS[row_idx % len(_NODE_COLORS)]
            for ci, dof_idx in zip(_DOF_COLS, range(3)):
                val = str(row[ci]).strip().upper()
                if val in ("S", "L"):
                    key = (nid, dof_idx)
                    plot_dofs.append(key)
                    color_map[key] = color
                    base = user_label if user_label else str(nid)
                    label_map[key] = f"{base} {_DOF_NAMES[dof_idx]}"
                if val == "L":
                    limit_dofs.append((nid, dof_idx))
        return plot_dofs, limit_dofs, color_map, label_map

    # ── Paste dialog ──────────────────────────────────────────────────────

    def _paste_dialog(self, title, prompt, initial=""):
        result = {}
        win = ctk.CTkToplevel(self.frame.winfo_toplevel())
        win.title(title)
        win.geometry("460x340")
        win.resizable(True, True)
        win.transient(self.frame.winfo_toplevel())
        win.grab_set()

        ctk.CTkLabel(win, text=prompt, anchor=tk.W).pack(fill=tk.X, padx=10, pady=(10, 4))
        tb = ctk.CTkTextbox(win, wrap="none")
        tb.pack(fill=tk.BOTH, expand=True, padx=10, pady=4)
        if initial:
            tb.insert("1.0", initial)

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

    def _compute_notch(self, silent=False):
        if self._op2 is None:
            if not silent:
                messagebox.showwarning("No OP2", "Load an FRF OP2 first.")
            return
        if self._input_asd_freqs is None:
            if not silent:
                messagebox.showwarning("No Input ASD", "Load an Input ASD first.")
            return
        if self._limit_asd_freqs is None:
            if not silent:
                messagebox.showwarning("No Limit ASD", "Load a Response Limit ASD first.")
            return

        plot_dofs, limit_dofs, color_map, label_map = self._get_grid_dofs()
        if not limit_dofs:
            if not silent:
                messagebox.showwarning("No Limit DOFs",
                                       "Set at least one X/Y/Z cell to L in the node grid.")
            return

        sc = self._get_subcase_int()
        if sc is None:
            return

        cfg = RESPONSE_TYPES['Acceleration']
        unit_factor = cfg['unit_factors'].get(self._units_var.get(), 386.089)
        id_attr = cfg['id_attr']

        frf_dict = getattr(self._op2, cfg['frf_attr'], None) or {}
        frf_tbl  = lookup_subcase(frf_dict, sc)
        if frf_tbl is None:
            return

        freqs = frf_tbl._times
        arr   = getattr(frf_tbl, id_attr)
        entity_ids = arr[:, 0] if id_attr == 'node_gridtype' else arr

        orig_interp  = interp_loglog(self._input_asd_freqs, self._input_asd_vals, freqs)
        limit_interp = interp_loglog(self._limit_asd_freqs, self._limit_asd_vals, freqs)

        min_allowed = np.full(len(freqs), np.inf)
        missing     = []
        limit_dofs_set = set()

        for nid, dof_idx in limit_dofs:
            hits = np.where(entity_ids == nid)[0]
            if not len(hits):
                missing.append(f"{nid} {_DOF_NAMES[dof_idx]}")
                continue
            H_g = frf_tbl.data[:, hits[0], dof_idx] / unit_factor
            H2  = H_g.real**2 + H_g.imag**2
            with np.errstate(divide='ignore', invalid='ignore'):
                min_allowed = np.minimum(
                    min_allowed, np.where(H2 > 0, limit_interp / H2, np.inf))
            limit_dofs_set.add((nid, dof_idx))

        notched = np.minimum(orig_interp, min_allowed)
        notched = np.where(np.isinf(notched), orig_interp, notched)

        if self._notch_enabled_var.get():
            try:
                db = float(self._notch_db_var.get())
            except ValueError:
                db = 6.0
            notched = np.maximum(notched, orig_interp * 10**(-db / 10.0))

        response_curves = {}
        for nid, dof_idx in plot_dofs:
            hits = np.where(entity_ids == nid)[0]
            if not len(hits):
                if f"{nid} {_DOF_NAMES[dof_idx]}" not in missing:
                    missing.append(f"{nid} {_DOF_NAMES[dof_idx]}")
                continue
            H_g = frf_tbl.data[:, hits[0], dof_idx] / unit_factor
            H2  = H_g.real**2 + H_g.imag**2
            response_curves[(nid, dof_idx)] = (H2 * orig_interp, H2 * notched)

        self._frf_freqs        = freqs
        self._orig_asd_interp  = orig_interp
        self._limit_asd_interp = limit_interp
        self._notched_asd      = notched
        self._response_curves  = response_curves
        self._limit_dofs_set   = limit_dofs_set
        self._color_map        = color_map
        self._label_map        = label_map

        # Update curve picker
        labels = []
        for key in response_curves:
            tag = " [L]" if key in limit_dofs_set else ""
            labels.append(label_map.get(key, f"Node {key[0]} {_DOF_NAMES[key[1]]}") + tag)
        self._curve_menu.configure(values=labels if labels else ["(none)"])
        if labels and (self._curve_var.get() == "(none)" or
                       self._curve_var.get() not in labels):
            self._curve_var.set(labels[0])

        self._export_excel_btn.configure(state=tk.NORMAL)
        self._export_csv_btn.configure(state=tk.NORMAL)

        parts = [f"Notched: {len(response_curves)} curve(s)."]
        if missing:
            parts.append(f"Not found: {', '.join(missing[:4])}")
        self._status_label.configure(text="  ".join(parts),
                                     text_color=("gray10", "gray90"))
        self._redraw()

    def _clear_results(self):
        self._frf_freqs        = None
        self._orig_asd_interp  = None
        self._limit_asd_interp = None
        self._notched_asd      = None
        self._response_curves  = {}
        self._limit_dofs_set   = set()
        self._color_map        = {}
        self._label_map        = {}
        self._export_excel_btn.configure(state=tk.DISABLED)
        self._export_csv_btn.configure(state=tk.DISABLED)
        self._curve_menu.configure(values=["(none)"])
        self._curve_var.set("(none)")

    # ── Plot helpers ──────────────────────────────────────────────────────

    def _get_theme(self):
        return _THEMES[self._plot_theme]

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

    def _legend(self, ax):
        th = self._get_theme()
        leg = ax.legend(facecolor=th["legend_bg"],
                        edgecolor=th["spine"],
                        labelcolor=th["text"],
                        fontsize=8)
        return leg

    def _draw_idle_plot(self):
        th = self._get_theme()
        self._fig.clf()
        ax = self._fig.add_subplot(111)
        self._ax = ax
        self._format_plot(ax, title="Response Limiting")
        ax.text(0.5, 0.5,
                "1. Open FRF OP2\n2. Load Input ASD + Response Limit\n"
                "3. Add nodes, set X/Y/Z to L\n"
                "Results update automatically",
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
        ax.loglog(freqs, self._notched_asd, color="#d62728", label="Notched Input")
        og = (np.sqrt(max(0.0, grms_loglog(self._input_asd_freqs, self._input_asd_vals)))
              if self._input_asd_freqs is not None else 0.0)
        ng = np.sqrt(max(0.0, grms_loglog(freqs, self._notched_asd)))
        self._legend(ax).set_title(
            f"Orig:    {og:.3g} g GRMS\nNotched: {ng:.3g} g GRMS")
        self._format_plot(ax, title="Input ASD: Original vs Notched")
        ax.set_xlim(freqs[0], freqs[-1])

    def _draw_response_dof_view(self):
        ax = self._ax
        freqs = self._frf_freqs
        sel = self._curve_var.get()

        selected_key = None
        for key in self._response_curves:
            tag = " [L]" if key in self._limit_dofs_set else ""
            lbl = self._label_map.get(key, f"Node {key[0]} {_DOF_NAMES[key[1]]}") + tag
            if lbl == sel:
                selected_key = key
                break
        if selected_key is None and self._response_curves:
            selected_key = next(iter(self._response_curves))

        rb_grms = ra_grms = 0.0
        if selected_key is not None:
            rb, ra = self._response_curves[selected_key]
            col = self._color_map.get(selected_key, "#1f77b4")
            ax.loglog(freqs, rb, color=col, linestyle="--", label="Before notch")
            ax.loglog(freqs, ra, color=col,              label="After notch")
            rb_grms = np.sqrt(max(0.0, grms_loglog(freqs, rb)))
            ra_grms = np.sqrt(max(0.0, grms_loglog(freqs, ra)))

        ax.loglog(freqs, self._limit_asd_interp,
                  color="#ff7f0e", linestyle=":", linewidth=2, label="Response Limit")
        leg = self._legend(ax)
        leg.set_title(f"Before: {rb_grms:.3g} g\nAfter:  {ra_grms:.3g} g")

        if selected_key:
            lbl = self._label_map.get(selected_key,
                                       f"Node {selected_key[0]} {_DOF_NAMES[selected_key[1]]}")
        else:
            lbl = "(none)"
        self._format_plot(ax, title=f"Response: {lbl}")
        ax.set_xlim(freqs[0], freqs[-1])

    def _draw_response_all_view(self):
        ax = self._ax
        freqs = self._frf_freqs
        for key, (_, ra) in self._response_curves.items():
            col = self._color_map.get(key, "#aaaaaa")
            tag = " [L]" if key in self._limit_dofs_set else ""
            lbl = self._label_map.get(key, f"Node {key[0]} {_DOF_NAMES[key[1]]}") + tag
            ax.loglog(freqs, ra, color=col, label=lbl)
        ax.loglog(freqs, self._limit_asd_interp,
                  color="#ff7f0e", linestyle=":", linewidth=2, label="Response Limit")
        self._legend(ax)
        self._format_plot(ax, title="All Responses (notched) vs Limit  ([L] = limit DOF)")
        ax.set_xlim(freqs[0], freqs[-1])

    def _draw_grms_view(self):
        freqs = self._frf_freqs
        og = np.sqrt(max(0.0, grms_loglog(freqs, self._orig_asd_interp)))
        ng = np.sqrt(max(0.0, grms_loglog(freqs, self._notched_asd)))
        with np.errstate(divide='ignore', invalid='ignore'):
            ratio = np.where(self._notched_asd > 0,
                             self._orig_asd_interp / self._notched_asd, 1.0)
        max_db = 10.0 * np.log10(max(float(ratio.max()), 1.0))
        rows = []
        for key, (rb, ra) in self._response_curves.items():
            tag = " [L]" if key in self._limit_dofs_set else ""
            lbl = self._label_map.get(key, f"Node {key[0]} {_DOF_NAMES[key[1]]}") + tag
            rows.append([
                lbl,
                f"{og:.4g}", f"{ng:.4g}",
                f"{np.sqrt(max(0.0, grms_loglog(freqs, rb))):.4g}",
                f"{np.sqrt(max(0.0, grms_loglog(freqs, ra))):.4g}",
                f"{max_db:.2f}",
            ])
        self._grms_sheet.set_sheet_data(rows)

    # ── Export ────────────────────────────────────────────────────────────

    def _export_excel(self):
        if self._notched_asd is None:
            messagebox.showwarning("No Data", "Compute notch first.")
            return
        path = filedialog.asksaveasfilename(
            title="Save Excel", defaultextension=".xlsx",
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
        for f, o, n in zip(freqs, self._orig_asd_interp, self._notched_asd):
            ws.append([float(f), float(o), float(n)])

        ws2 = wb.create_sheet("Limit ASD")
        ws2.append(["Freq (Hz)", "Limit ASD (g²/Hz)"])
        for f, lim in zip(freqs, self._limit_asd_interp):
            ws2.append([float(f), float(lim)])

        for key, (rb, ra) in self._response_curves.items():
            lbl = self._label_map.get(key, f"N{key[0]}_{_DOF_NAMES[key[1]]}")
            sname = lbl[:31]
            ws_d = wb.create_sheet(sname)
            ws_d.append(["Freq (Hz)", "Resp Before (g²/Hz)", "Resp After (g²/Hz)"])
            for f, r_b, r_a in zip(freqs, rb, ra):
                ws_d.append([float(f), float(r_b), float(r_a)])

        ws_g = wb.create_sheet("GRMS Summary")
        ws_g.append(["DOF", "Orig Input GRMS (g)", "Notched Input GRMS (g)",
                      "Resp Before (g)", "Resp After (g)", "Max Notch (dB)"])
        og = np.sqrt(max(0.0, grms_loglog(freqs, self._orig_asd_interp)))
        ng = np.sqrt(max(0.0, grms_loglog(freqs, self._notched_asd)))
        with np.errstate(divide='ignore', invalid='ignore'):
            ratio = np.where(self._notched_asd > 0,
                             self._orig_asd_interp / self._notched_asd, 1.0)
        max_db = 10.0 * np.log10(max(float(ratio.max()), 1.0))
        for key, (rb, ra) in self._response_curves.items():
            tag = " [L]" if key in self._limit_dofs_set else ""
            lbl = self._label_map.get(key, f"Node {key[0]} {_DOF_NAMES[key[1]]}") + tag
            ws_g.append([lbl, round(og, 6), round(ng, 6),
                          round(np.sqrt(max(0.0, grms_loglog(freqs, rb))), 6),
                          round(np.sqrt(max(0.0, grms_loglog(freqs, ra))), 6),
                          round(max_db, 4)])

        wb.save(path)
        self._status_label.configure(text=f"Saved: {os.path.basename(path)}",
                                     text_color=("gray10", "gray90"))

    def _export_csv(self):
        if self._notched_asd is None:
            messagebox.showwarning("No Data", "Compute notch first.")
            return
        path = filedialog.asksaveasfilename(
            title="Save CSV", defaultextension=".csv",
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
