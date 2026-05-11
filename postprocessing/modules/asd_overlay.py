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

_SLOT_COLORS = ("#1f77b4", "#d62728")   # blue-A, red-B
_SLOT_TAGS = ("A", "B")
_SLOT_LINES = ("-", "--")

_DARK_BG = "#2b2b2b"
_PLOT_BG = "#1e1e1e"
_GRID_COLOR = "#3a3a3a"
_TEXT_COLOR = "#c0c0c0"
_SPINE_COLOR = "#505050"


class AsdOverlayModule:
    name = "ASD Overlay"

    DOF_LABELS = ("T1 (X)", "T2 (Y)", "T3 (Z)", "R1", "R2", "R3")
    UNIT_OPTIONS = ("in/s²", "m/s²")
    UNIT_FACTORS = {"in/s²": 386.089, "m/s²": 9.80665}

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
   Or use "CSV" to load a file with columns: grid_id, label
   (header row is auto-detected and skipped).
4. Check/uncheck nodes to show or hide their curves on the plot.
5. Use the DOF dropdown to switch which response direction is plotted.
   All checked nodes across both OP2s update simultaneously.

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

NAVIGATION
Use the matplotlib toolbar below the plot to zoom, pan, and save images.
"""

    def __init__(self, parent):
        self.frame = ctk.CTkFrame(parent)

        # Slot state: {0: slot_A, 1: slot_B}
        self._op2_slots = {0: self._empty_slot(), 1: self._empty_slot()}

        # Node rows: [{"id": int, "label": str, "checked": BooleanVar, "row_frame": widget}]
        self._nodes = []

        self._dof_var = ctk.StringVar(value="T3 (Z)")

        # Per-slot UI widgets populated in _build_ui
        self._open_btn = [None, None]
        self._file_label = [None, None]
        self._unit_var = [ctk.StringVar(value="in/s²"), ctk.StringVar(value="in/s²")]
        self._sc_var = [tk.StringVar(value="(none)"), tk.StringVar(value="(none)")]
        self._sc_menu = [None, None]

        self._build_ui()

    @staticmethod
    def _empty_slot():
        return {"op2": None, "path": None, "subcase": None}

    # ── UI construction ──────────────────────────────────────────────────────

    def _build_ui(self):
        # ── Toolbar ──────────────────────────────────────────────────────────
        toolbar = ctk.CTkFrame(self.frame, fg_color="transparent")
        toolbar.pack(fill=tk.X, padx=5, pady=(5, 0))

        for i, tag in enumerate(_SLOT_TAGS):
            row = ctk.CTkFrame(toolbar, fg_color="transparent")
            row.pack(fill=tk.X, pady=1)

            color = _SLOT_COLORS[i]
            btn = ctk.CTkButton(
                row, text=f"Open OP2 {tag}…", width=120,
                command=lambda idx=i: self._open_op2(idx),
                fg_color=color,
            )
            btn.pack(side=tk.LEFT, padx=(0, 6))
            self._open_btn[i] = btn

            lbl = ctk.CTkLabel(row, text="(no file)", text_color="gray",
                               anchor=tk.W, width=200)
            lbl.pack(side=tk.LEFT, padx=(0, 12))
            self._file_label[i] = lbl

            ctk.CTkLabel(row, text="Units:").pack(side=tk.LEFT, padx=(0, 2))
            ctk.CTkOptionMenu(
                row, variable=self._unit_var[i],
                values=list(self.UNIT_OPTIONS),
                command=lambda _v, idx=i: self._refresh_plot(),
                width=80,
            ).pack(side=tk.LEFT, padx=(0, 12))

            ctk.CTkLabel(row, text="Subcase:").pack(side=tk.LEFT, padx=(0, 2))
            scmenu = ctk.CTkOptionMenu(
                row, variable=self._sc_var[i], values=["(none)"],
                command=lambda _v, idx=i: self._on_sc_select(idx),
                width=120,
            )
            scmenu.pack(side=tk.LEFT)
            self._sc_menu[i] = scmenu

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
        ).pack(side=tk.LEFT)

        self._status_label = ctk.CTkLabel(
            dof_row, text="Open an OP2 to begin", text_color="gray")
        self._status_label.pack(side=tk.LEFT, padx=(10, 0))

        # ── Body: node panel + plot ───────────────────────────────────────────
        body = ctk.CTkFrame(self.frame, fg_color="transparent")
        body.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Left: node management panel
        node_panel = ctk.CTkFrame(body, width=210)
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
        ctk.CTkButton(btn_row1, text="CSV", width=50,
                      command=self._load_csv).pack(side=tk.LEFT, padx=3)
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

        # Right: matplotlib plot
        plot_container = ctk.CTkFrame(body)
        plot_container.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._fig = Figure(figsize=(8, 5), dpi=100, facecolor=_DARK_BG)
        self._ax = self._fig.add_subplot(111)

        self._canvas = FigureCanvasTkAgg(self._fig, master=plot_container)
        self._canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        toolbar_tk = tk.Frame(plot_container)
        toolbar_tk.pack(side=tk.BOTTOM, fill=tk.X)
        self._mpl_toolbar = NavigationToolbar2Tk(self._canvas, toolbar_tk)
        self._mpl_toolbar.update()

        self._draw_empty_axes()

    def _draw_empty_axes(self):
        ax = self._ax
        ax.clear()
        ax.set_facecolor(_PLOT_BG)
        ax.set_xlabel("Frequency (Hz)", color=_TEXT_COLOR)
        ax.set_ylabel("ASD (g²/Hz)", color=_TEXT_COLOR)
        ax.tick_params(colors=_TEXT_COLOR, which="both")
        for spine in ax.spines.values():
            spine.set_edgecolor(_SPINE_COLOR)
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
            if not psd_dict:
                messagebox.showwarning(
                    "No PSD Data",
                    f"OP2 {tag} contains no PSD acceleration results.\n\n"
                    "Ensure the deck includes:\n"
                    "  RANDOM = <sid>\n"
                    "  ACCELERATION(PLOT) = ALL")
                self._file_label[slot_idx].configure(
                    text="(no PSD data)", text_color="orange")
                return

            self._op2_slots[slot_idx]['op2'] = op2
            self._op2_slots[slot_idx]['path'] = path

            subcases = sorted(psd_dict.keys())
            sc_strs = [str(s) for s in subcases]
            self._sc_menu[slot_idx].configure(values=sc_strs)
            self._sc_var[slot_idx].set(sc_strs[0])
            self._op2_slots[slot_idx]['subcase'] = subcases[0]

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
        self._op2_slots[slot_idx]['subcase'] = int(val) if val != "(none)" else None
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

    def _load_csv(self):
        path = filedialog.askopenfilename(
            title="Load Nodes CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if not path:
            return
        try:
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
                        else f"Node {gid}"
                    lines.append(f"{gid},{label}")
        except Exception as exc:
            messagebox.showerror("CSV Error", str(exc))
            return
        self._parse_and_add_nodes("\n".join(lines))

    def _add_node_row(self, gid, label):
        var = tk.BooleanVar(value=True)
        row_frame = ctk.CTkFrame(self._node_scroll, fg_color="transparent")
        row_frame.pack(fill=tk.X, pady=1)
        ctk.CTkCheckBox(
            row_frame,
            text=f"{gid}  {label}",
            variable=var,
            command=self._refresh_plot,
        ).pack(anchor=tk.W, padx=4)
        self._nodes.append({"id": gid, "label": label,
                             "checked": var, "row_frame": row_frame})

    def _clear_nodes(self):
        for n in self._nodes:
            n['row_frame'].destroy()
        self._nodes.clear()
        self._refresh_plot()

    def _select_all(self, state):
        for n in self._nodes:
            n['checked'].set(state)
        self._refresh_plot()

    # ── RMS helper ───────────────────────────────────────────────────────────

    @staticmethod
    def _get_rms_g(op2, subcase, nid, idof, freqs, psd_curve_g2hz, unit_factor):
        """Return RMS in g for one node/DOF curve.

        Prefers the Nastran RMS acceleration table; falls back to trapz
        integration of the supplied psd_curve_g2hz array.
        """
        try:
            rms_dict = op2.op2_results.rms.accelerations
            if subcase in rms_dict:
                rms_tbl = rms_dict[subcase]
                nids_arr = rms_tbl.node_gridtype[:, 0]
                hits = np.where(nids_arr == nid)[0]
                if len(hits):
                    # data shape: (1, nnodes, 6) — single RMS value per node/DOF
                    rms_native = float(rms_tbl.data[0, hits[0], idof])
                    return rms_native / unit_factor
        except Exception:
            pass
        return float(np.sqrt(np.trapz(psd_curve_g2hz, freqs)))

    # ── Plot ─────────────────────────────────────────────────────────────────

    def _refresh_plot(self):
        ax = self._ax
        ax.clear()
        ax.set_facecolor(_PLOT_BG)

        idof = self.DOF_LABELS.index(self._dof_var.get())
        has_curves = False

        for slot_idx, slot in self._op2_slots.items():
            op2 = slot['op2']
            subcase = slot['subcase']
            if op2 is None or subcase is None:
                continue

            psd_dict = op2.op2_results.psd.accelerations
            if subcase not in psd_dict:
                continue

            psd_tbl = psd_dict[subcase]
            freqs = psd_tbl._times
            op2_nids = psd_tbl.node_gridtype[:, 0]
            unit_factor = self.UNIT_FACTORS[self._unit_var[slot_idx].get()]
            conv = unit_factor ** 2
            tag = _SLOT_TAGS[slot_idx]
            color = _SLOT_COLORS[slot_idx]
            ls = _SLOT_LINES[slot_idx]

            for node in self._nodes:
                if not node['checked'].get():
                    continue
                nid = node['id']
                hits = np.where(op2_nids == nid)[0]
                if not len(hits):
                    continue

                inode = hits[0]
                raw_psd = psd_tbl.data[:, inode, idof]
                psd_g2hz = raw_psd / conv

                rms_g = self._get_rms_g(
                    op2, subcase, nid, idof, freqs, psd_g2hz, unit_factor)

                label = f"{tag}: {node['label']}  (RMS = {rms_g:.3g} g)"
                ax.loglog(freqs, psd_g2hz, label=label, color=color, linestyle=ls)
                has_curves = True

        ax.set_xlabel("Frequency (Hz)", color=_TEXT_COLOR)
        ax.set_ylabel("ASD (g²/Hz)", color=_TEXT_COLOR)
        ax.tick_params(colors=_TEXT_COLOR, which="both")
        for spine in ax.spines.values():
            spine.set_edgecolor(_SPINE_COLOR)
        ax.grid(True, which="both", alpha=0.25, color=_GRID_COLOR)
        self._fig.set_facecolor(_DARK_BG)

        if has_curves:
            ax.legend(loc="best", fontsize=8,
                      facecolor="#383838", labelcolor=_TEXT_COLOR,
                      edgecolor=_SPINE_COLOR)
        else:
            ax.text(0.5, 0.5,
                    "No data — check OP2 loaded, nodes added, and boxes checked",
                    transform=ax.transAxes,
                    ha="center", va="center", color="gray", fontsize=10)

        self._canvas.draw_idle()
