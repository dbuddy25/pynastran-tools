"""CBUSH element forces module.

Reads CBUSH force results (Fx, Fy, Fz, Mx, My, Mz) from a Nastran OP2
file across all subcases/load cases.  For random analysis the values are
RMS.  Each load case can be named by the user and exported to its own
Excel sheet.
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
        'cell_border': Border(bottom=Side(style='thin', color="B4C6E7")),
        'sci_fmt': '0.00E+00',
    }


def write_cbush_sheet(ws, eids, forces, styles, op2_name=None,
                       title=None, sheet_label=None):
    """Write one CBUSH forces worksheet.

    Parameters
    ----------
    ws : openpyxl Worksheet
    eids : array-like, shape (nelems,)
    forces : ndarray, shape (nelems, 6)
    styles : dict from ``make_cbush_styles()``
    op2_name : str or None — OP2 filename for the header
    title : str or None — optional user title (row 1)
    sheet_label : str or None — load case label (e.g. "Subcase 3 — My Name")
    """
    from openpyxl.utils import get_column_letter

    s = styles
    total_cols = len(_HEADERS)
    cur_row = 0

    # Row 1 — Title (always written; blank if no title)
    cur_row += 1
    cell = ws.cell(row=cur_row, column=1, value=title or "")
    cell.font = s['white_bold']
    cell.fill = s['dark_fill']
    cell.alignment = s['center']
    ws.merge_cells(start_row=cur_row, start_column=1,
                   end_row=cur_row, end_column=total_cols)
    for ci in range(2, total_cols + 1):
        ws.cell(row=cur_row, column=ci).fill = s['dark_fill']

    # Row 2 — OP2 filename (always written; blank if no op2_name)
    cur_row += 1
    cell = ws.cell(row=cur_row, column=1, value=op2_name or "")
    cell.font = s['white_bold']
    cell.fill = s['dark_fill']
    cell.alignment = s['center']
    ws.merge_cells(start_row=cur_row, start_column=1,
                   end_row=cur_row, end_column=total_cols)
    for ci in range(2, total_cols + 1):
        ws.cell(row=cur_row, column=ci).fill = s['dark_fill']

    # Row 3 — Load case label (always written; blank if no label)
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
    for ci, h in enumerate(_HEADERS, 1):
        cell = ws.cell(row=cur_row, column=ci, value=h)
        cell.font = s['sub_font']
        cell.fill = s['mid_fill']
        cell.alignment = s['center']

    # Data rows
    data_start = cur_row + 1
    for i in range(len(eids)):
        row = data_start + i
        ws.cell(row=row, column=1, value=int(eids[i])).alignment = s['center']
        for j in range(6):
            c = ws.cell(row=row, column=j + 2, value=float(forces[i, j]))
            c.number_format = s['sci_fmt']
            c.alignment = s['center']
        for ci in range(1, total_cols + 1):
            ws.cell(row=row, column=ci).border = s['cell_border']

    # Column widths
    ws.column_dimensions['A'].width = 10  # EID
    for ci in range(2, total_cols + 1):
        ws.column_dimensions[get_column_letter(ci)].width = 14

    ws.freeze_panes = f'A{data_start}'


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

        tk.Label(self, text=message, justify='left',
                 padx=16, pady=(16, 8)).pack()

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
        self._title_var = tk.StringVar(value='')
        self._name_var = tk.StringVar(value='')

        # Per-subcase state
        self._subcase_data = {}    # {sc_id: ndarray(nelems, 6)}
        self._subcase_eids = {}    # {sc_id: ndarray}
        self._subcase_titles = {}  # {sc_id: str}  (OP2 subtitle)
        self._subcase_names = {}   # {sc_id: str}  (user-editable)
        self._subcase_order = []   # sorted subcase IDs
        self._active_subcase = None

        self._build_ui()

    # ------------------------------------------------------------------ UI
    _GUIDE_TEXT = """\
CBUSH Forces Tool — Quick Guide

PURPOSE
Extract CBUSH element force results (Fx, Fy, Fz, Mx, My, Mz) from a
Nastran OP2 file.  Displays forces per element for each subcase / load
case.  For random analysis the values represent RMS.

WORKFLOW
1. Open OP2 — select a Nastran OP2 containing CBUSH force output.
   (Requires FORCE(PLOT) = ALL or FORCE = ALL in case control.)
2. Select Load Case — use the dropdown to switch between subcases.
3. Name Load Cases — type a descriptive name in the Name field.
   Names persist when switching between subcases.
4. Export to Excel — each load case gets its own worksheet, named
   from your load case names (or "Subcase N" if unnamed).

TITLE FIELD
Optional title text that appears as a header row in every Excel sheet.
Leave blank to omit.

EXCEL EXPORT
Produces a styled workbook with:
  - One sheet per subcase / load case
  - Scientific notation for force values (0.000E+00)
  - Frozen header rows for easy scrolling
  - Color-coded header bands

REQUIREMENTS
  - pyNastran (for OP2 reading)
  - numpy
  - openpyxl (for Excel export only)\
"""

    def _build_ui(self):
        toolbar = ctk.CTkFrame(self.frame, fg_color="transparent")
        toolbar.pack(fill=tk.X, padx=5, pady=(5, 0))

        self._op2_btn = ctk.CTkButton(toolbar, text="Open OP2\u2026", width=100,
                                      command=self._open_op2)
        self._op2_btn.pack(side=tk.LEFT)

        # Load case dropdown
        ctk.CTkLabel(toolbar, text="Load Case:").pack(
            side=tk.LEFT, padx=(10, 2))
        self._lc_var = tk.StringVar(value="(none)")
        self._lc_menu = ctk.CTkOptionMenu(
            toolbar, variable=self._lc_var, values=["(none)"],
            command=self._on_lc_select, width=220)
        self._lc_menu.pack(side=tk.LEFT, padx=(0, 4))

        # Per-subcase name entry
        ctk.CTkLabel(toolbar, text="Name:").pack(side=tk.LEFT, padx=(10, 2))
        name_entry = ctk.CTkEntry(toolbar, textvariable=self._name_var,
                                  width=180)
        name_entry.pack(side=tk.LEFT, padx=(0, 4))
        self._name_var.trace_add('write', lambda *_: self._on_lc_name_change())

        # Title field
        ctk.CTkLabel(toolbar, text="Title:").pack(side=tk.LEFT, padx=(10, 2))
        ctk.CTkEntry(toolbar, textvariable=self._title_var, width=160).pack(
            side=tk.LEFT, padx=(0, 4))

        # Right-side buttons
        ctk.CTkButton(
            toolbar, text="?", width=30, font=ctk.CTkFont(weight="bold"),
            command=self._show_guide,
        ).pack(side=tk.RIGHT, padx=(5, 0))

        ctk.CTkButton(toolbar, text="Export to Excel\u2026", width=130,
                       command=self._export_excel).pack(side=tk.RIGHT)

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
            "single_select", "copy", "arrowkeys",
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

    # -------------------------------------------------------------- load
    def load(self, op2):
        """Extract CBUSH force data from all subcases."""
        self._subcase_data.clear()
        self._subcase_eids.clear()
        self._subcase_titles.clear()
        self._subcase_names.clear()
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
            # Element IDs — may be 2D (ntimes, nelems)
            eids = result.element
            if eids.ndim == 2:
                eids = eids[0]
            self._subcase_eids[sc_id] = eids

            # Force data — first time step: result.data shape (ntimes, nelems, 6)
            data = result.data
            if data.ndim == 3:
                forces = data[0]  # (nelems, 6)
            else:
                forces = data
            self._subcase_data[sc_id] = forces

            subtitle = getattr(result, 'subtitle', '').strip()
            self._subcase_titles[sc_id] = subtitle
            self._subcase_names[sc_id] = ''

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
        self._show_subcase()

    # ---------------------------------------------------------- subcase switching
    def _on_lc_select(self, choice):
        """Handle load case dropdown selection."""
        # Save current name
        if self._active_subcase is not None:
            self._subcase_names[self._active_subcase] = self._name_var.get()

        # Parse subcase ID from label ("Subcase 3" or "Subcase 3 - title")
        idx = self._lc_menu.cget("values").index(choice)
        self._active_subcase = self._subcase_order[idx]

        # Restore name for new subcase
        self._name_var.set(self._subcase_names.get(self._active_subcase, ''))
        self._show_subcase()

    def _on_lc_name_change(self):
        """Store name entry value for the active subcase."""
        if self._active_subcase is not None:
            self._subcase_names[self._active_subcase] = self._name_var.get()

    def _show_subcase(self):
        """Populate sheet with the active subcase data."""
        sc = self._active_subcase
        if sc is None or sc not in self._subcase_data:
            self._sheet.set_sheet_data([])
            return

        eids = self._subcase_eids[sc]
        forces = self._subcase_data[sc]

        rows = []
        for i in range(len(eids)):
            row = [int(eids[i])]
            for j in range(6):
                row.append(f"{forces[i, j]:.2E}")
            rows.append(row)

        self._sheet.set_sheet_data(rows)
        ncols = len(_HEADERS)
        self._sheet.set_all_column_widths(90)
        self._sheet.column_width(column=0, width=80)  # EID narrower
        self._sheet.align_columns(
            list(range(ncols)), align="center", align_header=True)

    # ------------------------------------------------------------ export
    def _export_excel(self):
        if not self._subcase_data:
            messagebox.showinfo("Nothing to export",
                                "Open an OP2 file first.")
            return

        # Save current name before export
        if self._active_subcase is not None:
            self._subcase_names[self._active_subcase] = self._name_var.get()

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

        wb = Workbook()
        used_names = set()

        for i, sc_id in enumerate(self._subcase_order):
            user_name = self._subcase_names.get(sc_id, '').strip()

            # Build sheet name (31-char Excel limit, deduplicated)
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

            # Build load case label for header
            lc_label = f"Subcase {sc_id}"
            if effective_name:
                lc_label += f" \u2014 {effective_name}"

            if i == 0:
                ws = wb.active
                ws.title = sheet_name
            else:
                ws = wb.create_sheet(title=sheet_name)

            write_cbush_sheet(
                ws,
                self._subcase_eids[sc_id],
                self._subcase_data[sc_id],
                styles,
                op2_name=op2_name,
                title=title,
                sheet_label=lc_label,
            )

        try:
            wb.save(path)
            n = len(self._subcase_order)
            _ExportDoneDialog(
                self.frame.winfo_toplevel(),
                f"Saved {n} sheet(s) to:\n{path}",
                path,
            )
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
