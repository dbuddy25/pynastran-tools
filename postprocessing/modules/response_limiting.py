"""FRF Response Limiting — compute a notched drive ASD from FRF data.

Loads a Nastran frequency-response OP2 (SOL 108/111), an input environment
ASD, a response limit ASD, and a list of (node, direction) DOFs.  Computes
the notched input ASD such that the response at every limited DOF stays at or
below the limit.
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


_DARK_BG = "#2b2b2b"
_DOF_NAMES = ["X", "Y", "Z"]  # maps to dof_index 0, 1, 2

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


class ResponseLimitingModule:
    name = "Response Limiting"

    def __init__(self, parent):
        self.frame = ctk.CTkFrame(parent)

        # ── State ─────────────────────────────────────────────────────────
        self._op2 = None
        self._op2_path = None
        self._subcase_opts = []   # [(sc_int, label), ...]
        self._subcase_var = ctk.StringVar(value="—")
        self._units_var = ctk.StringVar(value="in/s²")
        self._rt_var = ctk.StringVar(value="Acceleration")

        self._input_asd_freqs = None
        self._input_asd_vals = None
        self._limit_asd_freqs = None
        self._limit_asd_vals = None

        self._limited_dofs = []   # [(node_id: int, dof_idx: int), ...]
        self._notch_enabled_var = ctk.BooleanVar(value=False)
        self._notch_db_var = ctk.StringVar(value="6.0")

        # Computed results (all on FRF frequency grid)
        self._frf_freqs = None
        self._orig_asd_interp = None
        self._limit_asd_interp = None
        self._notched_asd = None
        self._response_curves = {}  # {(nid, dof_idx): (resp_before, resp_after)}

        self._view_var = ctk.StringVar(value="input")
        self._dof_picker_var = ctk.StringVar(value="—")

        self._build_ui()

    # ── UI construction ────────────────────────────────────────────────────

    def _build_ui(self):
        toolbar = ctk.CTkFrame(self.frame, fg_color="transparent")
        toolbar.pack(fill=tk.X, padx=6, pady=(6, 2))

        # Row 0: OP2 controls
        row0 = ctk.CTkFrame(toolbar, fg_color="transparent")
        row0.pack(fill=tk.X, pady=1)

        self._open_btn = ctk.CTkButton(row0, text="Open OP2…", width=110,
                                       command=self._open_op2)
        self._open_btn.pack(side=tk.LEFT, padx=(0, 4))

        ctk.CTkButton(row0, text="Clear", width=60,
                      command=self._clear_op2).pack(side=tk.LEFT, padx=(0, 10))

        ctk.CTkLabel(row0, text="Run:").pack(side=tk.LEFT)
        self._file_label = ctk.CTkLabel(row0, text="(none)", text_color="gray",
                                        width=200, anchor=tk.W)
        self._file_label.pack(side=tk.LEFT, padx=4)

        ctk.CTkLabel(row0, text="Type:").pack(side=tk.LEFT, padx=(10, 2))
        ctk.CTkOptionMenu(row0, variable=self._rt_var,
                          values=["Acceleration"],
                          width=120).pack(side=tk.LEFT)

        ctk.CTkLabel(row0, text="Units:").pack(side=tk.LEFT, padx=(8, 2))
        self._units_menu = ctk.CTkOptionMenu(row0, variable=self._units_var,
                                             values=["in/s²", "m/s²"],
                                             width=90)
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

        # Row 2: notch controls + action buttons
        row2 = ctk.CTkFrame(toolbar, fg_color="transparent")
        row2.pack(fill=tk.X, pady=1)

        self._notch_cb = ctk.CTkCheckBox(row2, text="Max notch depth (dB):",
                                          variable=self._notch_enabled_var,
                                          command=self._on_notch_toggle)
        self._notch_cb.pack(side=tk.LEFT)
        self._notch_entry = ctk.CTkEntry(row2, textvariable=self._notch_db_var,
                                          width=54, state=tk.DISABLED)
        self._notch_entry.pack(side=tk.LEFT, padx=(4, 16))

        self._compute_btn = ctk.CTkButton(row2, text="Compute Notch", width=120,
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
        left = ctk.CTkScrollableFrame(body, width=280)
        left.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 6))

        # DOF section
        ctk.CTkLabel(left, text="Limited DOFs",
                     font=ctk.CTkFont(weight="bold")).pack(anchor=tk.W, pady=(4, 2))

        dof_btn_row = ctk.CTkFrame(left, fg_color="transparent")
        dof_btn_row.pack(fill=tk.X, pady=(0, 4))
        ctk.CTkButton(dof_btn_row, text="Add…", width=60,
                      command=self._add_dof_dialog).pack(side=tk.LEFT, padx=(0, 2))
        ctk.CTkButton(dof_btn_row, text="Remove", width=70,
                      command=self._remove_selected_dofs).pack(side=tk.LEFT, padx=2)
        ctk.CTkButton(dof_btn_row, text="Paste…", width=60,
                      command=self._paste_dofs).pack(side=tk.LEFT, padx=2)
        ctk.CTkButton(dof_btn_row, text="Clear All", width=70,
                      command=self._clear_dofs).pack(side=tk.LEFT, padx=2)

        self._dof_sheet = Sheet(
            left,
            headers=["Node ID", "Dir"],
            height=160,
            show_row_index=False,
            theme="dark" if ctk.get_appearance_mode() == "Dark" else "light",
        )
        self._dof_sheet.pack(fill=tk.X, pady=(0, 8))
        self._dof_sheet.enable_bindings("single_select", "row_select", "column_width_resize")
        self._dof_sheet.column_width(0, 100)
        self._dof_sheet.column_width(1, 60)

        # View section
        ctk.CTkLabel(left, text="View",
                     font=ctk.CTkFont(weight="bold")).pack(anchor=tk.W, pady=(8, 2))

        for val, label in [
            ("input", "Input ASDs (orig / notched)"),
            ("response_dof", "Response — selected DOF"),
            ("response_all", "All notched responses overlay"),
            ("grms", "GRMS Summary"),
        ]:
            ctk.CTkRadioButton(left, text=label, variable=self._view_var,
                               value=val, command=self._redraw).pack(anchor=tk.W, pady=1)

        # DOF picker for response_dof view
        self._dof_picker_row = ctk.CTkFrame(left, fg_color="transparent")
        self._dof_picker_row.pack(fill=tk.X, pady=(4, 0))
        ctk.CTkLabel(self._dof_picker_row, text="DOF:").pack(side=tk.LEFT)
        self._dof_picker_menu = ctk.CTkOptionMenu(
            self._dof_picker_row, variable=self._dof_picker_var,
            values=["—"], command=lambda _: self._redraw(), width=140)
        self._dof_picker_menu.pack(side=tk.LEFT, padx=4)

        # Right — matplotlib canvas + GRMS table overlay
        right = ctk.CTkFrame(body, fg_color="transparent")
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._fig = Figure(figsize=(8, 5))
        self._ax = self._fig.add_subplot(111)
        self._canvas = FigureCanvasTkAgg(self._fig, master=right)
        self._canvas_widget = self._canvas.get_tk_widget()
        self._canvas_widget.pack(fill=tk.BOTH, expand=True)

        toolbar_mpl = NavigationToolbar2Tk(self._canvas, right)
        toolbar_mpl.update()

        # GRMS table frame (hidden by default, shown in grms view)
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

    # ── Notch toggle ──────────────────────────────────────────────────────

    def _on_notch_toggle(self):
        state = tk.NORMAL if self._notch_enabled_var.get() else tk.DISABLED
        self._notch_entry.configure(state=state)

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

            rt = self._rt_var.get()
            cfg = RESPONSE_TYPES.get(rt, RESPONSE_TYPES['Acceleration'])
            frf_dict = getattr(op2, cfg['frf_attr'], None) or {}

            if not frf_dict:
                messagebox.showwarning(
                    "No FRF Data",
                    "This OP2 has no FRF data.\n\n"
                    "Load a SOL 108/111 frequency-response run with:\n"
                    "  ACCELERATION(PLOT) = ALL")
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
                text=f"{os.path.basename(path)} — {len(sc_pairs)} subcase(s), FRF mode",
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
    def _asd_status_text(path_or_label, freqs):
        n = len(freqs)
        f0 = freqs[0]
        fn = freqs[-1]
        return f"{path_or_label} ({n} pts, {f0:.1f}–{fn:.1f} Hz)"

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
        self._input_asd_freqs = freqs
        self._input_asd_vals = asds
        stem = os.path.basename(path)
        self._input_status.configure(
            text=self._asd_status_text(stem, freqs), text_color=("gray10", "gray90"))
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
        self._input_asd_freqs = freqs
        self._input_asd_vals = asds
        self._input_status.configure(
            text=self._asd_status_text("(pasted)", freqs), text_color=("gray10", "gray90"))
        self._clear_results()

    def _clear_input_asd(self):
        self._input_asd_freqs = None
        self._input_asd_vals = None
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
        self._limit_asd_freqs = freqs
        self._limit_asd_vals = asds
        stem = os.path.basename(path)
        self._limit_status.configure(
            text=self._asd_status_text(stem, freqs), text_color=("gray10", "gray90"))
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
        self._limit_asd_freqs = freqs
        self._limit_asd_vals = asds
        self._limit_status.configure(
            text=self._asd_status_text("(pasted)", freqs), text_color=("gray10", "gray90"))
        self._clear_results()

    def _clear_limit_asd(self):
        self._limit_asd_freqs = None
        self._limit_asd_vals = None
        self._limit_status.configure(text="(none)", text_color="gray")
        self._clear_results()

    # ── DOF management ────────────────────────────────────────────────────

    def _add_dof_dialog(self):
        win = ctk.CTkToplevel(self.frame.winfo_toplevel())
        win.title("Add Limited DOF")
        win.geometry("280x140")
        win.resizable(False, False)
        win.transient(self.frame.winfo_toplevel())
        win.grab_set()

        ctk.CTkLabel(win, text="Node ID:").grid(row=0, column=0, padx=10, pady=(12, 4), sticky=tk.W)
        nid_var = ctk.StringVar()
        ctk.CTkEntry(win, textvariable=nid_var, width=120).grid(row=0, column=1, padx=10, pady=(12, 4))

        ctk.CTkLabel(win, text="Direction:").grid(row=1, column=0, padx=10, pady=4, sticky=tk.W)
        dir_var = ctk.StringVar(value="X")
        ctk.CTkOptionMenu(win, variable=dir_var, values=_DOF_NAMES,
                          width=120).grid(row=1, column=1, padx=10, pady=4)

        def _add():
            try:
                nid = int(nid_var.get().strip())
            except ValueError:
                messagebox.showerror("Invalid", "Node ID must be an integer.", parent=win)
                return
            dof_idx = _DOF_NAMES.index(dir_var.get())
            if (nid, dof_idx) not in self._limited_dofs:
                self._limited_dofs.append((nid, dof_idx))
                self._refresh_dof_sheet()
            win.destroy()

        btn_row = ctk.CTkFrame(win, fg_color="transparent")
        btn_row.grid(row=2, column=0, columnspan=2, pady=(8, 4))
        ctk.CTkButton(btn_row, text="Add", width=80, command=_add).pack(side=tk.LEFT, padx=4)
        ctk.CTkButton(btn_row, text="Cancel", width=80, command=win.destroy).pack(side=tk.LEFT, padx=4)

    def _remove_selected_dofs(self):
        rows = sorted(self._dof_sheet.get_selected_rows(), reverse=True)
        for r in rows:
            if 0 <= r < len(self._limited_dofs):
                del self._limited_dofs[r]
        self._refresh_dof_sheet()

    def _paste_dofs(self):
        text = self._paste_dialog(
            "Paste DOFs",
            "Paste DOFs, one per line: node_id  direction\n"
            "Direction: X, Y, or Z  (e.g.  1001 X)")
        if text is None:
            return
        added = 0
        for line in text.splitlines():
            line = line.strip()
            if not line or line.startswith('#'):
                continue
            parts = line.replace(',', ' ').split()
            if len(parts) < 2:
                continue
            try:
                nid = int(parts[0])
                direction = parts[1].upper()
                if direction not in _DOF_NAMES:
                    continue
                dof_idx = _DOF_NAMES.index(direction)
                if (nid, dof_idx) not in self._limited_dofs:
                    self._limited_dofs.append((nid, dof_idx))
                    added += 1
            except (ValueError, IndexError):
                continue
        self._refresh_dof_sheet()
        if added:
            self._status_label.configure(text=f"Added {added} DOF(s).",
                                         text_color=("gray10", "gray90"))

    def _clear_dofs(self):
        self._limited_dofs.clear()
        self._refresh_dof_sheet()

    def _refresh_dof_sheet(self):
        data = [[str(nid), _DOF_NAMES[dof_idx]]
                for nid, dof_idx in self._limited_dofs]
        self._dof_sheet.set_sheet_data(data)

    # ── Paste dialog helper ────────────────────────────────────────────────

    def _paste_dialog(self, title, prompt):
        """Show a modal textbox dialog. Returns text string or None on cancel."""
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

        def _cancel():
            win.destroy()

        ctk.CTkButton(btn_row, text="OK", width=80, command=_ok).pack(side=tk.LEFT, padx=4)
        ctk.CTkButton(btn_row, text="Cancel", width=80, command=_cancel).pack(side=tk.LEFT, padx=4)
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
        if not self._limited_dofs:
            messagebox.showwarning("No DOFs", "Add at least one limited DOF.")
            return

        sc = self._get_subcase_int()
        if sc is None:
            messagebox.showwarning("No Subcase", "Select a valid subcase.")
            return

        rt = self._rt_var.get()
        cfg = RESPONSE_TYPES.get(rt, RESPONSE_TYPES['Acceleration'])
        unit_factor = cfg['unit_factors'].get(self._units_var.get(), 1.0)
        id_attr = cfg['id_attr']

        frf_dict = getattr(self._op2, cfg['frf_attr'], None) or {}
        frf_tbl = lookup_subcase(frf_dict, sc)
        if frf_tbl is None:
            messagebox.showerror("Error", f"Subcase {sc} not found in FRF results.")
            return

        freqs = frf_tbl._times  # (nfreq,)
        arr = getattr(frf_tbl, id_attr)
        entity_ids = arr[:, 0] if id_attr == 'node_gridtype' else arr

        # Interpolate input and limit ASDs onto FRF frequency grid
        orig_interp = interp_loglog(self._input_asd_freqs, self._input_asd_vals, freqs)
        limit_interp = interp_loglog(self._limit_asd_freqs, self._limit_asd_vals, freqs)

        # Compute minimum allowed input ASD across all limited DOFs
        min_allowed = np.full(len(freqs), np.inf)
        missing_nodes = []

        for nid, dof_idx in self._limited_dofs:
            hits = np.where(entity_ids == nid)[0]
            if not len(hits):
                missing_nodes.append(f"{nid} {_DOF_NAMES[dof_idx]}")
                continue
            H_native = frf_tbl.data[:, hits[0], dof_idx]  # complex
            H_g = H_native / unit_factor
            H_mag2 = H_g.real**2 + H_g.imag**2
            with np.errstate(divide='ignore', invalid='ignore'):
                allowed = np.where(H_mag2 > 0, limit_interp / H_mag2, np.inf)
            min_allowed = np.minimum(min_allowed, allowed)

        notched = np.minimum(orig_interp, min_allowed)
        # Replace inf values (unconstrained bins) with original
        notched = np.where(np.isinf(notched), orig_interp, notched)

        # Apply notch floor if enabled
        if self._notch_enabled_var.get():
            try:
                db = float(self._notch_db_var.get())
            except ValueError:
                db = 6.0
            floor = orig_interp * 10**(-db / 10.0)
            notched = np.maximum(notched, floor)

        # Compute response curves for each limited DOF
        response_curves = {}
        for nid, dof_idx in self._limited_dofs:
            hits = np.where(entity_ids == nid)[0]
            if not len(hits):
                continue
            H_native = frf_tbl.data[:, hits[0], dof_idx]
            H_g = H_native / unit_factor
            H_mag2 = H_g.real**2 + H_g.imag**2
            resp_before = H_mag2 * orig_interp
            resp_after = H_mag2 * notched
            response_curves[(nid, dof_idx)] = (resp_before, resp_after)

        # Store results
        self._frf_freqs = freqs
        self._orig_asd_interp = orig_interp
        self._limit_asd_interp = limit_interp
        self._notched_asd = notched
        self._response_curves = response_curves

        # Update DOF picker
        dof_labels = [f"Node {nid} {_DOF_NAMES[dof_idx]}"
                      for nid, dof_idx in response_curves]
        self._dof_picker_menu.configure(values=dof_labels if dof_labels else ["—"])
        if dof_labels:
            self._dof_picker_var.set(dof_labels[0])

        # Enable exports
        self._export_excel_btn.configure(state=tk.NORMAL)
        self._export_csv_btn.configure(state=tk.NORMAL)

        # Status
        status_parts = [f"Notch computed: {len(response_curves)} DOF(s)."]
        if missing_nodes:
            status_parts.append(f"Not found: {', '.join(missing_nodes[:4])}")
        self._status_label.configure(text="  ".join(status_parts),
                                     text_color=("gray10", "gray90"))

        self._redraw()

    def _clear_results(self):
        self._frf_freqs = None
        self._orig_asd_interp = None
        self._limit_asd_interp = None
        self._notched_asd = None
        self._response_curves = {}
        self._export_excel_btn.configure(state=tk.DISABLED)
        self._export_csv_btn.configure(state=tk.DISABLED)
        self._dof_picker_menu.configure(values=["—"])
        self._dof_picker_var.set("—")

    # ── Plot helpers ──────────────────────────────────────────────────────

    def _get_theme(self):
        mode = ctk.get_appearance_mode().lower()
        return _THEMES.get(mode, _THEMES["light"])

    def _format_plot(self, ax, title="", xlabel="Frequency (Hz)", ylabel=""):
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
        ax.set_xlabel(xlabel)
        if ylabel:
            ax.set_ylabel(ylabel)
        if title:
            ax.set_title(title)

    def _draw_idle_plot(self):
        th = self._get_theme()
        self._fig.clf()
        ax = self._fig.add_subplot(111)
        self._ax = ax
        self._format_plot(ax, title="Response Limiting")
        ax.text(0.5, 0.5, "Load an OP2 and define DOFs,\nthen click Compute Notch.",
                ha='center', va='center', transform=ax.transAxes,
                color=th["text"], fontsize=12)
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
        # Original input (on its own native grid)
        if self._input_asd_freqs is not None:
            ax.loglog(self._input_asd_freqs, self._input_asd_vals,
                      color="#1f77b4", label="Original Input")
        # Notched input (on FRF grid)
        ax.loglog(freqs, self._notched_asd,
                  color="#d62728", label="Notched Input")
        # Annotate GRMS
        orig_grms = np.sqrt(max(0, grms_loglog(self._input_asd_freqs, self._input_asd_vals))) \
            if self._input_asd_freqs is not None else 0.0
        notch_grms = np.sqrt(max(0, grms_loglog(freqs, self._notched_asd)))
        ax.legend(title=f"Orig GRMS: {orig_grms:.3g} g\nNotch GRMS: {notch_grms:.3g} g")
        self._format_plot(ax, title="Input ASD: Original vs Notched",
                          ylabel="ASD (g²/Hz)")
        ax.set_xlim(left=freqs[0], right=freqs[-1])

    def _draw_response_dof_view(self):
        ax = self._ax
        freqs = self._frf_freqs
        sel = self._dof_picker_var.get()
        # Identify selected DOF
        selected_key = None
        for nid, dof_idx in self._response_curves:
            if f"Node {nid} {_DOF_NAMES[dof_idx]}" == sel:
                selected_key = (nid, dof_idx)
                break
        if selected_key is None and self._response_curves:
            selected_key = next(iter(self._response_curves))

        if selected_key is not None:
            resp_before, resp_after = self._response_curves[selected_key]
            ax.loglog(freqs, resp_before, color="#1f77b4", label="Response (original)")
            ax.loglog(freqs, resp_after, color="#d62728", label="Response (notched)")

        # Limit ASD
        ax.loglog(freqs, self._limit_asd_interp,
                  color="#ff7f0e", linestyle="--", label="Response Limit")

        ax.legend()
        title = f"Response: {sel}" if sel and sel != "—" else "Response"
        self._format_plot(ax, title=title, ylabel="ASD (g²/Hz)")
        ax.set_xlim(left=freqs[0], right=freqs[-1])

    def _draw_response_all_view(self):
        ax = self._ax
        freqs = self._frf_freqs
        colors = ("#1f77b4", "#2ca02c", "#9467bd", "#17becf", "#bcbd22",
                  "#d62728", "#ff7f0e", "#8c564b", "#e377c2", "#7f7f7f")
        for i, ((nid, dof_idx), (_, resp_after)) in enumerate(self._response_curves.items()):
            col = colors[i % len(colors)]
            label = f"Node {nid} {_DOF_NAMES[dof_idx]}"
            ax.loglog(freqs, resp_after, color=col, label=label)

        ax.loglog(freqs, self._limit_asd_interp,
                  color="#ff7f0e", linestyle="--", linewidth=2, label="Response Limit")
        ax.legend(fontsize=8)
        self._format_plot(ax, title="All Notched Responses vs Limit",
                          ylabel="ASD (g²/Hz)")
        ax.set_xlim(left=freqs[0], right=freqs[-1])

    def _draw_grms_view(self):
        freqs = self._frf_freqs
        orig_grms = np.sqrt(max(0, grms_loglog(freqs, self._orig_asd_interp)))
        notch_grms = np.sqrt(max(0, grms_loglog(freqs, self._notched_asd)))

        rows = []
        for (nid, dof_idx), (resp_before, resp_after) in self._response_curves.items():
            dof_label = f"Node {nid} {_DOF_NAMES[dof_idx]}"
            rb_grms = np.sqrt(max(0, grms_loglog(freqs, resp_before)))
            ra_grms = np.sqrt(max(0, grms_loglog(freqs, resp_after)))
            # Max notch depth: 10*log10(orig / notched) at any frequency
            with np.errstate(divide='ignore', invalid='ignore'):
                notch_ratio = np.where(self._notched_asd > 0,
                                       self._orig_asd_interp / self._notched_asd, 1.0)
            max_notch_db = 10.0 * np.log10(max(notch_ratio.max(), 1.0))
            rows.append([
                dof_label,
                f"{orig_grms:.4g}",
                f"{notch_grms:.4g}",
                f"{rb_grms:.4g}",
                f"{ra_grms:.4g}",
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

        wb = openpyxl.Workbook()
        freqs = self._frf_freqs

        # Sheet: Notched ASD
        ws = wb.active
        ws.title = "Notched ASD"
        ws.append(["Freq (Hz)", "Original ASD (g²/Hz)", "Notched ASD (g²/Hz)"])
        for f, orig, notch in zip(freqs, self._orig_asd_interp, self._notched_asd):
            ws.append([float(f), float(orig), float(notch)])

        # Sheet: Limit ASD
        ws2 = wb.create_sheet("Limit ASD")
        ws2.append(["Freq (Hz)", "Limit ASD (g²/Hz)"])
        for f, lim in zip(freqs, self._limit_asd_interp):
            ws2.append([float(f), float(lim)])

        # Per-DOF sheets
        for (nid, dof_idx), (resp_before, resp_after) in self._response_curves.items():
            sname = f"N{nid}_{_DOF_NAMES[dof_idx]}"[:31]
            ws_dof = wb.create_sheet(sname)
            ws_dof.append(["Freq (Hz)", "Resp Before (g²/Hz)", "Resp After (g²/Hz)"])
            for f, rb, ra in zip(freqs, resp_before, resp_after):
                ws_dof.append([float(f), float(rb), float(ra)])

        # GRMS Summary sheet
        ws_grms = wb.create_sheet("GRMS Summary")
        ws_grms.append(["DOF", "Orig Input GRMS (g)", "Notched Input GRMS (g)",
                         "Resp Before (g)", "Resp After (g)", "Max Notch (dB)"])
        orig_grms = np.sqrt(max(0, grms_loglog(freqs, self._orig_asd_interp)))
        notch_grms = np.sqrt(max(0, grms_loglog(freqs, self._notched_asd)))
        with np.errstate(divide='ignore', invalid='ignore'):
            notch_ratio = np.where(self._notched_asd > 0,
                                   self._orig_asd_interp / self._notched_asd, 1.0)
        max_notch_db = 10.0 * np.log10(max(notch_ratio.max(), 1.0))
        for (nid, dof_idx), (resp_before, resp_after) in self._response_curves.items():
            dof_label = f"Node {nid} {_DOF_NAMES[dof_idx]}"
            rb_grms = np.sqrt(max(0, grms_loglog(freqs, resp_before)))
            ra_grms = np.sqrt(max(0, grms_loglog(freqs, resp_after)))
            ws_grms.append([dof_label, round(orig_grms, 6), round(notch_grms, 6),
                             round(rb_grms, 6), round(ra_grms, 6), round(max_notch_db, 4)])

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
