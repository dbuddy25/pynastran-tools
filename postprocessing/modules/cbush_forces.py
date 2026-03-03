"""CBUSH element forces module.

Reads CBUSH force results (Fx, Fy, Fz, Mx, My, Mz) from a Nastran OP2
file across all subcases/load cases.  For random analysis the values are
RMS.  Each load case can be named by the user and exported to its own
Excel sheet.

Optional BDF loading adds a Property column (comment names from PBUSH
cards) and enables sorting by property group.  Per-case scale factors
are applied on export only.
"""
import os
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk
from tksheet import Sheet

import numpy as np

_HEADERS = ['EID', 'Fx', 'Fy', 'Fz', 'Mx', 'My', 'Mz']
_HEADERS_WITH_PROP = ['Property', 'EID', 'Fx', 'Fy', 'Fz', 'Mx', 'My', 'Mz']
_FORCE_COLS = ['fx', 'fy', 'fz', 'mx', 'my', 'mz']


# --------------------------------------------------------- Excel helpers

def make_cbush_styles():
    """Return a dict of openpyxl style objects for CBUSH Forces sheets."""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    return {
        'dark_fill': PatternFill("solid", fgColor="1F4E79"),
        'mid_fill': PatternFill("solid", fgColor="2E75B6"),
        'white_bold': Font(bold=True, color="FFFFFF", size=11),
        'sub_font': Font(bold=True, color="FFFFFF", size=10),
        'center': Alignment(horizontal="center", vertical="center"),
        'left': Alignment(horizontal="left", vertical="center"),
        'cell_border': Border(bottom=Side(style='thin', color="B4C6E7")),
        'sci_fmt': '0.00E+00',
    }


def write_cbush_sheet(ws, eids, forces, styles, op2_name=None,
                       title=None, sheet_label=None, prop_names=None,
                       scale_factor=1.0, start_row=0):
    """Write one CBUSH forces block to an openpyxl worksheet.

    Parameters
    ----------
    ws : openpyxl Worksheet
    eids : array-like, shape (nelems,)
    forces : ndarray, shape (nelems, 6)
    styles : dict from ``make_cbush_styles()``
    op2_name : str or None -- OP2 filename for the header
    title : str or None -- optional user title (row 1)
    sheet_label : str or None -- load case label
    prop_names : list[str] or None -- per-element property names (same order as eids)
    scale_factor : float -- multiplier applied to force values
    start_row : int -- 0-based row offset (for stacking multiple blocks)

    Returns
    -------
    int : next available row (0-based) after this block
    """
    from openpyxl.utils import get_column_letter

    s = styles
    has_prop = prop_names is not None
    headers = _HEADERS_WITH_PROP if has_prop else _HEADERS
    total_cols = len(headers)
    cur_row = start_row

    # Row 1 -- Title (always written; blank if no title)
    cur_row += 1
    cell = ws.cell(row=cur_row, column=1, value=title or "")
    cell.font = s['white_bold']
    cell.fill = s['dark_fill']
    cell.alignment = s['center']
    ws.merge_cells(start_row=cur_row, start_column=1,
                   end_row=cur_row, end_column=total_cols)
    for ci in range(2, total_cols + 1):
        ws.cell(row=cur_row, column=ci).fill = s['dark_fill']

    # Row 2 -- OP2 filename (always written; blank if no op2_name)
    cur_row += 1
    cell = ws.cell(row=cur_row, column=1, value=op2_name or "")
    cell.font = s['white_bold']
    cell.fill = s['dark_fill']
    cell.alignment = s['center']
    ws.merge_cells(start_row=cur_row, start_column=1,
                   end_row=cur_row, end_column=total_cols)
    for ci in range(2, total_cols + 1):
        ws.cell(row=cur_row, column=ci).fill = s['dark_fill']

    # Row 3 -- Load case label (always written; blank if no label)
    cur_row += 1
    cell = ws.cell(row=cur_row, column=1, value=sheet_label or "")
    cell.font = s['white_bold']
    cell.fill = s['dark_fill']
    cell.alignment = s['center']
    ws.merge_cells(start_row=cur_row, start_column=1,
                   end_row=cur_row, end_column=total_cols)
    for ci in range(2, total_cols + 1):
        ws.cell(row=cur_row, column=ci).fill = s['dark_fill']

    # Column headers
    cur_row += 1
    for ci, h in enumerate(headers, 1):
        cell = ws.cell(row=cur_row, column=ci, value=h)
        cell.font = s['sub_font']
        cell.fill = s['mid_fill']
        cell.alignment = s['center']

    # Data rows
    data_start = cur_row + 1
    force_col_offset = 2 if has_prop else 1  # 1-based column of first force
    eid_col = 2 if has_prop else 1

    for i in range(len(eids)):
        row = data_start + i
        if has_prop:
            c = ws.cell(row=row, column=1, value=prop_names[i])
            c.alignment = s['left']
        ws.cell(row=row, column=eid_col,
                value=int(eids[i])).alignment = s['center']
        for j in range(6):
            c = ws.cell(row=row, column=force_col_offset + 1 + j,
                        value=float(forces[i, j]) * scale_factor)
            c.number_format = s['sci_fmt']
            c.alignment = s['center']
        for ci in range(1, total_cols + 1):
            ws.cell(row=row, column=ci).border = s['cell_border']

    # Column widths (only set on first block — start_row == 0)
    if start_row == 0:
        if has_prop:
            ws.column_dimensions['A'].width = 18  # Property
            ws.column_dimensions['B'].width = 10  # EID
            for ci in range(3, total_cols + 1):
                ws.column_dimensions[get_column_letter(ci)].width = 14
        else:
            ws.column_dimensions['A'].width = 10  # EID
            for ci in range(2, total_cols + 1):
                ws.column_dimensions[get_column_letter(ci)].width = 14

    ws.freeze_panes = f'A{data_start}'

    return data_start + len(eids)  # next available row (0-based not needed; caller uses as start_row)


def _write_cbush_block_combined(ws, eids, forces, styles, sheet_label=None,
                                prop_names=None, scale_factor=1.0,
                                start_row=0):
    """Write a subcase block for combined-sheet mode (no title/OP2 rows).

    Returns next available row (1-based).
    """
    from openpyxl.utils import get_column_letter

    s = styles
    has_prop = prop_names is not None
    headers = _HEADERS_WITH_PROP if has_prop else _HEADERS
    total_cols = len(headers)
    cur_row = start_row

    # Load case label row
    cur_row += 1
    cell = ws.cell(row=cur_row, column=1, value=sheet_label or "")
    cell.font = s['white_bold']
    cell.fill = s['dark_fill']
    cell.alignment = s['center']
    ws.merge_cells(start_row=cur_row, start_column=1,
                   end_row=cur_row, end_column=total_cols)
    for ci in range(2, total_cols + 1):
        ws.cell(row=cur_row, column=ci).fill = s['dark_fill']

    # Column headers
    cur_row += 1
    for ci, h in enumerate(headers, 1):
        cell = ws.cell(row=cur_row, column=ci, value=h)
        cell.font = s['sub_font']
        cell.fill = s['mid_fill']
        cell.alignment = s['center']

    # Data rows
    data_start = cur_row + 1
    force_col_offset = 2 if has_prop else 1
    eid_col = 2 if has_prop else 1

    for i in range(len(eids)):
        row = data_start + i
        if has_prop:
            c = ws.cell(row=row, column=1, value=prop_names[i])
            c.alignment = s['left']
        ws.cell(row=row, column=eid_col,
                value=int(eids[i])).alignment = s['center']
        for j in range(6):
            c = ws.cell(row=row, column=force_col_offset + 1 + j,
                        value=float(forces[i, j]) * scale_factor)
            c.number_format = s['sci_fmt']
            c.alignment = s['center']
        for ci in range(1, total_cols + 1):
            ws.cell(row=row, column=ci).border = s['cell_border']

    return data_start + len(eids)  # next available 1-based row


# ---------------------------------------------------------- post-export dialog

def _open_path(path):
    """Open a file or directory with the platform default handler."""
    if sys.platform == 'darwin':
        subprocess.Popen(['open', path])
    elif sys.platform == 'win32':
        os.startfile(path)
    else:
        subprocess.Popen(['xdg-open', path])


class _ExportDoneDialog(tk.Toplevel):
    """Modal dialog shown after a successful Excel export."""

    def __init__(self, parent, message, file_path):
        super().__init__(parent)
        self.title("Exported")
        self.resizable(False, False)
        self._file_path = file_path

        tk.Label(self, text=message, justify='left').pack(
            padx=16, pady=(16, 8))

        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=(0, 16))

        tk.Button(btn_frame, text="Open File",
                  command=self._open_file).pack(side='left', padx=4)
        tk.Button(btn_frame, text="Open Folder",
                  command=self._open_folder).pack(side='left', padx=4)
        tk.Button(btn_frame, text="Close",
                  command=self.destroy).pack(side='left', padx=4)

        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self.destroy)

    def _open_file(self):
        _open_path(self._file_path)
        self.destroy()

    def _open_folder(self):
        _open_path(os.path.dirname(self._file_path))
        self.destroy()


# ---------------------------------------------------------------- GUI module

class CbushForcesModule:
    name = "CBUSH Forces"

    def __init__(self, parent):
        self.frame = ctk.CTkFrame(parent)
        self._op2_path = None
        self._bdf_path = None
        self._title_var = tk.StringVar(value='')
        self._name_var = tk.StringVar(value='')
        self._scale_var = tk.StringVar(value='1.0')
        self._combined_var = tk.BooleanVar(value=False)

        # Per-subcase state
        self._subcase_data = {}    # {sc_id: ndarray(nelems, 6)}
        self._subcase_eids = {}    # {sc_id: ndarray}
        self._subcase_titles = {}  # {sc_id: str}  (OP2 subtitle)
        self._subcase_names = {}   # {sc_id: str}  (user-editable)
        self._subcase_scales = {}  # {sc_id: str}  (scale factor per subcase)
        self._subcase_order = []   # sorted subcase IDs
        self._active_subcase = None

        # BDF mappings
        self._pid_names = {}       # {pid: comment name}
        self._eid_to_pid = {}      # {eid: pid}

        self._build_ui()

    # ------------------------------------------------------------------ UI
    _GUIDE_TEXT = """\
CBUSH Forces Tool -- Quick Guide

PURPOSE
Extract CBUSH element force results (Fx, Fy, Fz, Mx, My, Mz) from a
Nastran OP2 file.  Displays forces per element for each subcase / load
case.  For random analysis the values represent RMS.

WORKFLOW
1. Open OP2 -- select a Nastran OP2 containing CBUSH force output.
   (Requires FORCE(PLOT) = ALL or FORCE = ALL in case control.)
2. Open BDF (optional) -- load the corresponding BDF to add a Property
   column showing comment names from PBUSH cards.  Rows are sorted by
   property group when a BDF is loaded.
3. Select Load Case -- use the dropdown to switch between subcases.
4. Name Load Cases -- type a descriptive name in the Name field.
   Names persist when switching between subcases.
5. Scale Factor -- enter a multiplier per load case (applied on export
   only; the table always shows raw OP2 values).  Noted in the Excel
   header when not 1.0.
6. Export to Excel -- choose separate sheets (one per subcase) or
   combined (all subcases stacked on a single sheet).

TITLE FIELD
Optional title text that appears as a header row in every Excel sheet.
Leave blank to omit.

EXPORT MODES
  Separate sheets (default) -- one worksheet per load case
  All cases on one sheet -- subcases stacked vertically with a blank
    row between blocks

EXCEL OUTPUT
  - Scientific notation for force values (0.000E+00)
  - Frozen header rows for easy scrolling
  - Color-coded header bands
  - Property column included when BDF is loaded

REQUIREMENTS
  - pyNastran (for OP2/BDF reading)
  - numpy
  - openpyxl (for Excel export only)\
"""

    def _build_ui(self):
        # Row 1: Open OP2 | Open BDF | Load Case | Name | ? | Export
        row1 = ctk.CTkFrame(self.frame, fg_color="transparent")
        row1.pack(fill=tk.X, padx=5, pady=(5, 0))

        self._op2_btn = ctk.CTkButton(row1, text="Open OP2\u2026", width=100,
                                      command=self._open_op2)
        self._op2_btn.pack(side=tk.LEFT)

        self._bdf_btn = ctk.CTkButton(row1, text="Open BDF\u2026", width=100,
                                      command=self._open_bdf)
        self._bdf_btn.pack(side=tk.LEFT, padx=(5, 0))

        # Load case dropdown
        ctk.CTkLabel(row1, text="Load Case:").pack(
            side=tk.LEFT, padx=(10, 2))
        self._lc_var = tk.StringVar(value="(none)")
        self._lc_menu = ctk.CTkOptionMenu(
            row1, variable=self._lc_var, values=["(none)"],
            command=self._on_lc_select, width=220)
        self._lc_menu.pack(side=tk.LEFT, padx=(0, 4))

        # Per-subcase name entry
        ctk.CTkLabel(row1, text="Name:").pack(side=tk.LEFT, padx=(10, 2))
        name_entry = ctk.CTkEntry(row1, textvariable=self._name_var,
                                  width=180)
        name_entry.pack(side=tk.LEFT, padx=(0, 4))
        self._name_var.trace_add('write', lambda *_: self._on_lc_name_change())

        # Right-side buttons (row 1)
        ctk.CTkButton(
            row1, text="?", width=30, font=ctk.CTkFont(weight="bold"),
            command=self._show_guide,
        ).pack(side=tk.RIGHT, padx=(5, 0))

        ctk.CTkButton(row1, text="Export to Excel\u2026", width=130,
                       command=self._export_excel).pack(side=tk.RIGHT)

        # Row 2: Title | Scale Factor | Combined checkbox
        row2 = ctk.CTkFrame(self.frame, fg_color="transparent")
        row2.pack(fill=tk.X, padx=5, pady=(2, 0))

        ctk.CTkLabel(row2, text="Title:").pack(side=tk.LEFT, padx=(0, 2))
        ctk.CTkEntry(row2, textvariable=self._title_var, width=160).pack(
            side=tk.LEFT, padx=(0, 4))

        ctk.CTkLabel(row2, text="|", text_color="gray").pack(
            side=tk.LEFT, padx=6)

        ctk.CTkLabel(row2, text="Scale Factor:").pack(
            side=tk.LEFT, padx=(0, 2))
        self._scale_entry = ctk.CTkEntry(
            row2, textvariable=self._scale_var, width=60)
        self._scale_entry.pack(side=tk.LEFT, padx=(0, 4))
        self._scale_var.trace_add('write', lambda *_: self._on_scale_change())

        ctk.CTkLabel(row2, text="|", text_color="gray").pack(
            side=tk.LEFT, padx=6)

        ctk.CTkCheckBox(row2, text="All cases on one sheet",
                        variable=self._combined_var).pack(
            side=tk.LEFT, padx=(0, 4))

        # Status label
        self._status_label = ctk.CTkLabel(
            self.frame, text="No OP2 loaded", text_color="gray")
        self._status_label.pack(anchor=tk.W, padx=10, pady=(2, 0))

        # Table
        self._sheet = Sheet(
            self.frame,
            headers=list(_HEADERS),
            show_top_left=False,
            show_row_index=False,
        )
        self._sheet.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self._sheet.disable_bindings()
        self._sheet.enable_bindings(
            "single_select", "drag_select", "row_select",
            "column_select", "copy", "arrowkeys",
            "column_width_resize", "row_height_resize",
        )
        self._sheet.readonly_columns(columns=list(range(len(_HEADERS))))

    # ---------------------------------------------------------- Guide
    def _show_guide(self):
        try:
            from nastran_tools import show_guide
        except ImportError:
            return
        show_guide(self.frame.winfo_toplevel(), "CBUSH Forces Guide",
                   self._GUIDE_TEXT)

    # ---------------------------------------------------------- background work
    def _run_in_background(self, label, work_fn, done_fn):
        """Run *work_fn* in a background thread, keeping the UI responsive."""
        self._status_label.configure(text=label, text_color="gray")
        self._op2_btn.configure(state=tk.DISABLED)
        self._bdf_btn.configure(state=tk.DISABLED)

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
                self._op2_btn.configure(state=tk.NORMAL)
                self._bdf_btn.configure(state=tk.NORMAL)
                done_fn(container.get('result'), container.get('error'))

        thread = threading.Thread(target=_worker, daemon=True)
        thread.start()
        self.frame.after(50, _poll)

    # ---------------------------------------------------------- OP2 loading
    def _open_op2(self):
        path = filedialog.askopenfilename(
            title="Open OP2 File",
            filetypes=[("OP2 files", "*.op2"), ("All files", "*.*")])
        if not path:
            return

        def _work():
            from pyNastran.op2.op2 import OP2
            op2 = OP2(mode='nx')
            op2.read_op2(path)
            return op2

        def _done(op2, error):
            if error is not None:
                messagebox.showerror("Error",
                                     f"Could not read OP2:\n{error}")
                self._status_label.configure(text="Load failed",
                                             text_color="red")
                return

            self._op2_path = path
            self._status_label.configure(
                text=os.path.basename(path),
                text_color=("gray10", "gray90"))
            self.load(op2)

        self._run_in_background("Loading\u2026", _work, _done)

    # ---------------------------------------------------------- BDF loading
    def _open_bdf(self):
        path = filedialog.askopenfilename(
            title="Open BDF File",
            filetypes=[("BDF files", "*.bdf *.dat *.nas *.bulk"),
                       ("All files", "*.*")])
        if not path:
            return

        def _work():
            self._build_bdf_mappings(path)

        def _done(_result, error):
            if error is not None:
                messagebox.showerror("Error",
                                     f"Could not read BDF:\n{error}")
                self._status_label.configure(text="BDF load failed",
                                             text_color="red")
                return

            self._bdf_path = path

            status = ""
            if self._op2_path:
                status += f"OP2: {os.path.basename(self._op2_path)}  |  "
            status += (f"BDF: {os.path.basename(path)} "
                       f"({len(self._pid_names)} properties, "
                       f"{len(self._eid_to_pid)} elements)")
            self._status_label.configure(text=status,
                                         text_color=("gray10", "gray90"))

            if self._subcase_data:
                self._show_subcase()

        self._run_in_background("Loading BDF\u2026", _work, _done)

    @staticmethod
    def _extract_comment_name(comment):
        """Extract a descriptive name from a BDF card comment string.

        Takes the last non-empty comment line (directly above the card).
        Strips any prefix before a colon (e.g. ``$ Skin: Wing Upper``
        becomes ``Wing Upper``).
        """
        if not comment:
            return None
        result = None
        for line in comment.splitlines():
            line = line.strip().lstrip('$').strip()
            if line:
                if ':' in line:
                    line = line.split(':', 1)[1].strip()
                if line:
                    result = line
        return result

    def _build_bdf_mappings(self, bdf_path):
        """Build pid_names and eid_to_pid mappings from BDF."""
        from bdf_utils import make_model

        model = make_model()
        model.read_bdf(bdf_path)

        # Extract comment names from property cards
        self._pid_names = {}
        for pid, prop in model.properties.items():
            name = self._extract_comment_name(getattr(prop, 'comment', ''))
            if name:
                self._pid_names[pid] = name

        self._eid_to_pid = {}
        for eid, elem in model.elements.items():
            pid = getattr(elem, 'pid', None)
            if pid is not None:
                try:
                    pid_int = int(pid)
                except (ValueError, TypeError):
                    continue
                if pid_int != 0:
                    self._eid_to_pid[eid] = pid_int

    def _get_prop_name(self, eid):
        """Return property display name for an element ID."""
        pid = self._eid_to_pid.get(eid)
        if pid is None:
            return ""
        return self._pid_names.get(pid, f"PID {pid}")

    def _has_bdf(self):
        """True if BDF mappings are loaded."""
        return bool(self._eid_to_pid)

    # -------------------------------------------------------------- load
    def load(self, op2):
        """Extract CBUSH force data from all subcases."""
        self._subcase_data.clear()
        self._subcase_eids.clear()
        self._subcase_titles.clear()
        self._subcase_names.clear()
        self._subcase_scales.clear()
        self._subcase_order.clear()
        self._active_subcase = None
        self._sheet.set_sheet_data([])

        if not hasattr(op2, 'cbush_force') or not op2.cbush_force:
            messagebox.showwarning(
                "No CBUSH Forces",
                "No CBUSH element force results found in this OP2.\n\n"
                "Ensure your case control includes:\n"
                "  FORCE(PLOT) = ALL  (or FORCE = ALL)")
            self._status_label.configure(text="No CBUSH forces",
                                         text_color="red")
            return

        for sc_id, result in op2.cbush_force.items():
            # Element IDs -- may be 2D (ntimes, nelems)
            eids = result.element
            if eids.ndim == 2:
                eids = eids[0]
            self._subcase_eids[sc_id] = eids

            # Force data -- first time step: result.data shape (ntimes, nelems, 6)
            data = result.data
            if data.ndim == 3:
                forces = data[0]  # (nelems, 6)
            else:
                forces = data
            self._subcase_data[sc_id] = forces

            subtitle = getattr(result, 'subtitle', '').strip()
            self._subcase_titles[sc_id] = subtitle
            self._subcase_names[sc_id] = ''
            self._subcase_scales[sc_id] = '1.0'

        self._subcase_order = sorted(self._subcase_data.keys())

        if not self._subcase_order:
            return

        # Build dropdown labels
        labels = []
        for sc_id in self._subcase_order:
            sub = self._subcase_titles[sc_id]
            lbl = f"Subcase {sc_id}"
            if sub:
                lbl += f" - {sub}"
            labels.append(lbl)

        self._lc_menu.configure(values=labels)
        self._lc_var.set(labels[0])
        self._active_subcase = self._subcase_order[0]
        self._name_var.set('')
        self._scale_var.set('1.0')
        self._show_subcase()

    # ---------------------------------------------------------- subcase switching
    def _on_lc_select(self, choice):
        """Handle load case dropdown selection."""
        # Save current name and scale
        if self._active_subcase is not None:
            self._subcase_names[self._active_subcase] = self._name_var.get()
            self._subcase_scales[self._active_subcase] = self._scale_var.get()

        # Parse subcase ID from label
        idx = self._lc_menu.cget("values").index(choice)
        self._active_subcase = self._subcase_order[idx]

        # Restore name and scale for new subcase
        self._name_var.set(self._subcase_names.get(self._active_subcase, ''))
        self._scale_var.set(self._subcase_scales.get(self._active_subcase, '1.0'))
        self._show_subcase()

    def _on_lc_name_change(self):
        """Store name entry value for the active subcase."""
        if self._active_subcase is not None:
            self._subcase_names[self._active_subcase] = self._name_var.get()

    def _on_scale_change(self):
        """Store scale entry value for the active subcase."""
        if self._active_subcase is not None:
            self._subcase_scales[self._active_subcase] = self._scale_var.get()

    def _show_subcase(self):
        """Populate sheet with the active subcase data."""
        sc = self._active_subcase
        if sc is None or sc not in self._subcase_data:
            self._sheet.set_sheet_data([])
            return

        eids = self._subcase_eids[sc]
        forces = self._subcase_data[sc]
        has_bdf = self._has_bdf()

        if has_bdf:
            headers = list(_HEADERS_WITH_PROP)
            # Build rows with property names, then sort by (prop_name, eid)
            raw_rows = []
            for i in range(len(eids)):
                eid = int(eids[i])
                prop = self._get_prop_name(eid)
                row = [prop, eid]
                for j in range(6):
                    row.append(f"{forces[i, j]:.2E}")
                raw_rows.append(row)
            raw_rows.sort(key=lambda r: r[1])
            rows = raw_rows
        else:
            headers = list(_HEADERS)
            rows = []
            for i in range(len(eids)):
                row = [int(eids[i])]
                for j in range(6):
                    row.append(f"{forces[i, j]:.2E}")
                rows.append(row)

        self._sheet.headers(headers)
        self._sheet.set_sheet_data(rows)
        ncols = len(headers)
        self._sheet.set_all_column_widths(90)
        if has_bdf:
            self._sheet.column_width(column=0, width=140)  # Property wider
            self._sheet.column_width(column=1, width=80)   # EID narrower
        else:
            self._sheet.column_width(column=0, width=80)   # EID narrower
        self._sheet.align_columns(
            list(range(ncols)), align="center", align_header=True)
        self._sheet.readonly_columns(columns=list(range(ncols)))

    # ------------------------------------------------------------ export helpers

    def _get_scale_factor(self, sc_id):
        """Parse and return scale factor for a subcase, defaulting to 1.0."""
        try:
            return float(self._subcase_scales.get(sc_id, '1.0'))
        except (ValueError, TypeError):
            return 1.0

    def _prepare_export_data(self, sc_id):
        """Build sorted eids, forces, prop_names, scale, label for one subcase.

        Returns (eids, forces, prop_names_or_None, scale_factor, lc_label).
        """
        eids = self._subcase_eids[sc_id]
        forces = self._subcase_data[sc_id]
        has_bdf = self._has_bdf()
        scale = self._get_scale_factor(sc_id)

        if has_bdf:
            # Sort by (property_name, eid)
            indices = list(range(len(eids)))
            prop_list = [self._get_prop_name(int(eids[i])) for i in indices]
            order = sorted(indices, key=lambda i: int(eids[i]))
            eids = eids[order]
            forces = forces[order]
            prop_names = [prop_list[i] for i in order]
        else:
            prop_names = None

        # Build label
        user_name = self._subcase_names.get(sc_id, '').strip()
        effective_name = user_name or self._subcase_titles.get(sc_id, '')
        lc_label = f"Subcase {sc_id}"
        if effective_name:
            lc_label += f" \u2014 {effective_name}"
        if scale != 1.0:
            lc_label += f" (SF: {scale:g})"

        return eids, forces, prop_names, scale, lc_label

    def _make_sheet_name(self, sc_id, used_names):
        """Build a unique 31-char Excel sheet name."""
        user_name = self._subcase_names.get(sc_id, '').strip()
        effective_name = user_name or self._subcase_titles.get(sc_id, '')
        if effective_name:
            base = effective_name[:31]
        else:
            base = f"Subcase {sc_id}"[:31]

        sheet_name = base
        counter = 2
        while sheet_name.lower() in used_names:
            suffix = f" ({counter})"
            sheet_name = base[:31 - len(suffix)] + suffix
            counter += 1
        used_names.add(sheet_name.lower())
        return sheet_name

    # ------------------------------------------------------------ export
    def _export_excel(self):
        if not self._subcase_data:
            messagebox.showinfo("Nothing to export",
                                "Open an OP2 file first.")
            return

        # Save current name and scale before export
        if self._active_subcase is not None:
            self._subcase_names[self._active_subcase] = self._name_var.get()
            self._subcase_scales[self._active_subcase] = self._scale_var.get()

        path = filedialog.asksaveasfilename(
            defaultextension='.xlsx',
            filetypes=[("Excel workbook", "*.xlsx")])
        if not path:
            return

        try:
            from openpyxl import Workbook
        except ImportError:
            messagebox.showerror(
                "Missing dependency",
                "openpyxl is required for Excel export.\n\n"
                "pip install openpyxl")
            return

        op2_name = os.path.basename(self._op2_path) if self._op2_path else None
        title = self._title_var.get().strip() or None
        styles = make_cbush_styles()
        combined = self._combined_var.get()

        wb = Workbook()

        if combined:
            # All subcases on one sheet
            ws = wb.active
            ws.title = "CBUSH Forces"
            cur_row = 0

            for i, sc_id in enumerate(self._subcase_order):
                eids, forces, prop_names, scale, lc_label = \
                    self._prepare_export_data(sc_id)

                if i == 0:
                    # First block gets full header (title + OP2 name + case)
                    cur_row = write_cbush_sheet(
                        ws, eids, forces, styles,
                        op2_name=op2_name, title=title,
                        sheet_label=lc_label, prop_names=prop_names,
                        scale_factor=scale, start_row=cur_row)
                else:
                    # Subsequent blocks: blank row + case label + headers + data
                    cur_row += 1  # blank row
                    cur_row = _write_cbush_block_combined(
                        ws, eids, forces, styles,
                        sheet_label=lc_label, prop_names=prop_names,
                        scale_factor=scale, start_row=cur_row)

            # Set column widths once (if not already set by first block for
            # start_row==0 case in write_cbush_sheet)
            n = len(self._subcase_order)
            msg = f"Saved {n} case(s) to 1 sheet:\n{path}"
        else:
            # Separate sheets (original behavior)
            used_names = set()

            for i, sc_id in enumerate(self._subcase_order):
                eids, forces, prop_names, scale, lc_label = \
                    self._prepare_export_data(sc_id)
                sheet_name = self._make_sheet_name(sc_id, used_names)

                if i == 0:
                    ws = wb.active
                    ws.title = sheet_name
                else:
                    ws = wb.create_sheet(title=sheet_name)

                write_cbush_sheet(
                    ws, eids, forces, styles,
                    op2_name=op2_name, title=title,
                    sheet_label=lc_label, prop_names=prop_names,
                    scale_factor=scale,
                )

            n = len(self._subcase_order)
            msg = f"Saved {n} sheet(s) to:\n{path}"

        try:
            wb.save(path)
            _ExportDoneDialog(
                self.frame.winfo_toplevel(), msg, path)
        except Exception as exc:
            messagebox.showerror("Export failed", str(exc))


def main():
    import logging
    logging.getLogger("customtkinter").setLevel(logging.ERROR)
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")
    root = ctk.CTk()
    root.title("CBUSH Forces")
    root.geometry("1000x600")
    mod = CbushForcesModule(root)
    mod.frame.pack(fill=tk.BOTH, expand=True)
    root.mainloop()


if __name__ == '__main__':
    main()
