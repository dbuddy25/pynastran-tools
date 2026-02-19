"""Modal effective mass fractions module.

Reads EFMFACS from an OP2 and displays per-mode fractions with
cumulative sums for each direction (Tx-Rz).
"""
import os
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk
from tksheet import Sheet

import numpy as np
import scipy.sparse


DIRECTIONS = ['Tx', 'Ty', 'Tz', 'Rx', 'Ry', 'Rz']

# Header definitions for tksheet
_SINGLE_HEADERS = ['Mode', 'Freq (Hz)']
for _d in DIRECTIONS:
    _SINGLE_HEADERS.extend([f'{_d} Frac', f'{_d} Sum'])


# --------------------------------------------------------- Excel helpers

def make_meff_styles():
    """Return a dict of openpyxl style objects for MEFFMASS Excel sheets."""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    return {
        'dark_fill': PatternFill("solid", fgColor="1F4E79"),
        'mid_fill': PatternFill("solid", fgColor="2E75B6"),
        'white_bold': Font(bold=True, color="FFFFFF", size=11),
        'sub_font': Font(bold=True, color="FFFFFF", size=10),
        'center': Alignment(horizontal="center", vertical="center"),
        'right': Alignment(horizontal="right", vertical="center"),
        'cell_border': Border(bottom=Side(style='thin', color="B4C6E7")),
        'weak_font': Font(color="FF0000"),
        'bold_font': Font(bold=True),
        'bold_weak_font': Font(bold=True, color="FF0000"),
        'num4': '0.0000',
        'num2': '0.00',
        'num1': '0.0',
    }


def write_meff_single_sheet(ws, data, styles, op2_name=None, threshold=0.1):
    """Write a single-file MEFFMASS fraction sheet (row-1 title,
    row-2 merged direction headers, row-3 sub-headers, data from row 4)."""
    from openpyxl.utils import get_column_letter

    s = styles
    modes, freqs = data['modes'], data['freqs']
    frac, cumsum = data['frac'], data['cumsum']

    sub = ['Mode', 'Freq (Hz)']
    for _ in DIRECTIONS:
        sub.extend(['Frac', 'Sum'])
    total_cols = len(sub)

    # Row 1: title with OP2 filename
    title = op2_name if op2_name else ""
    cell = ws.cell(row=1, column=1, value=title)
    cell.font = s['white_bold']
    cell.fill = s['dark_fill']
    cell.alignment = s['center']
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
    for ci in range(2, total_cols + 1):
        ws.cell(row=1, column=ci).fill = s['dark_fill']

    # Row 2: direction group headers
    ws.cell(row=2, column=1, value="").fill = s['dark_fill']
    ws.cell(row=2, column=2, value="").fill = s['dark_fill']
    for idx, d in enumerate(DIRECTIONS):
        c1, c2 = 3 + idx * 2, 4 + idx * 2
        cell = ws.cell(row=2, column=c1, value=d)
        cell.font = s['white_bold']
        cell.fill = s['dark_fill']
        cell.alignment = s['center']
        ws.merge_cells(start_row=2, start_column=c1, end_row=2, end_column=c2)
        ws.cell(row=2, column=c2).fill = s['dark_fill']

    # Row 3: sub-headers
    for ci, h in enumerate(sub, 1):
        cell = ws.cell(row=3, column=ci, value=h)
        cell.font = s['sub_font']
        cell.fill = s['mid_fill']
        cell.alignment = s['center']

    # Data rows
    for i in range(len(modes)):
        row = i + 4
        ws.cell(row=row, column=1, value=int(modes[i])).alignment = s['center']
        c = ws.cell(row=row, column=2, value=float(freqs[i]))
        c.number_format = s['num1']
        c.alignment = s['center']
        for j in range(6):
            fc = ws.cell(row=row, column=3 + j * 2, value=float(frac[i, j]))
            fc.number_format = s['num2']
            fc.alignment = s['center']
            if frac[i, j] >= threshold:
                fc.font = s['bold_font']
            sc = ws.cell(row=row, column=4 + j * 2, value=float(cumsum[i, j]))
            sc.number_format = s['num2']
            sc.alignment = s['center']
        for ci in range(1, total_cols + 1):
            ws.cell(row=row, column=ci).border = s['cell_border']

    # Column widths
    ws.column_dimensions['A'].width = 7
    ws.column_dimensions['B'].width = 12
    for ci in range(3, total_cols + 1):
        ws.column_dimensions[get_column_letter(ci)].width = 9
    ws.freeze_panes = 'A4'


# ---------------------------------------------------------------- GUI module

class MeffModule:
    name = "MEFFMASS"

    def __init__(self, parent):
        self.frame = ctk.CTkFrame(parent)
        self.data = None
        self._op2_path = None
        self._threshold_var = tk.StringVar(value='0.1')
        self._build_ui()

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        toolbar = ctk.CTkFrame(self.frame, fg_color="transparent")
        toolbar.pack(fill=tk.X, padx=5, pady=(5, 0))

        ctk.CTkButton(toolbar, text="Open OP2\u2026", width=100,
                       command=self._open_op2).pack(side=tk.LEFT)

        ctk.CTkButton(toolbar, text="Export to Excel\u2026", width=130,
                       command=self._export_excel).pack(side=tk.RIGHT)

        # Threshold entry (packed RIGHT so it sits left of Export button)
        ctk.CTkEntry(toolbar, width=50,
                      textvariable=self._threshold_var).pack(
            side=tk.RIGHT, padx=(0, 4))
        ctk.CTkLabel(toolbar, text="Threshold:").pack(
            side=tk.RIGHT, padx=(10, 2))

        self._threshold_var.trace_add('write', self._on_threshold_change)

        # Status label
        self._status_label = ctk.CTkLabel(
            toolbar, text="No OP2 loaded", text_color="gray")
        self._status_label.pack(side=tk.LEFT, padx=(10, 0))

        # Table (tksheet)
        self._sheet = Sheet(
            self.frame,
            headers=list(_SINGLE_HEADERS),
            show_top_left=False,
            show_row_index=False,
        )
        self._sheet.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self._sheet.disable_bindings()
        self._sheet.enable_bindings(
            "single_select", "copy", "arrowkeys",
            "column_width_resize", "row_height_resize",
        )
        self._sheet.readonly_columns(
            columns=list(range(len(_SINGLE_HEADERS))))

    def _configure_sheet(self, headers):
        """Reconfigure sheet with new headers and clear data."""
        self._sheet.headers(headers)
        self._sheet.set_sheet_data([])
        self._sheet.readonly_columns(columns=list(range(len(headers))))

    # ---------------------------------------------------------- threshold helpers
    def _get_threshold(self):
        """Parse threshold from entry, default 0.1 on invalid input."""
        try:
            return float(self._threshold_var.get())
        except (ValueError, tk.TclError):
            return 0.1

    def _on_threshold_change(self, *args):
        """Re-display current view when threshold changes."""
        self._show_single_view()

    def _apply_highlights(self):
        """Apply threshold-based cell highlighting to the current view."""
        threshold = self._get_threshold()

        if self.data is not None:
            frac = self.data['frac']
            for i in range(len(self.data['modes'])):
                for j in range(6):
                    col = 2 + j * 2  # Frac columns: indices 2, 4, 6, 8, 10, 12
                    if frac[i, j] >= threshold:
                        self._sheet.highlight_cells(row=i, column=col, fg="blue")

    # ---------------------------------------------------------- OP2 loading
    def _open_op2(self):
        """Open a primary OP2 file."""
        path = filedialog.askopenfilename(
            title="Open OP2 File",
            filetypes=[("OP2 files", "*.op2"), ("All files", "*.*")])
        if not path:
            return

        self._status_label.configure(text=f"Loading\u2026", text_color="gray")
        self.frame.update_idletasks()

        try:
            from pyNastran.op2.op2 import OP2
            op2 = OP2(mode='nx')
            op2.read_op2(path)
        except Exception as exc:
            messagebox.showerror("Error", f"Could not read OP2:\n{exc}")
            self._status_label.configure(text="Load failed", text_color="red")
            return

        self._op2_path = path
        self._status_label.configure(
            text=path, text_color=("gray10", "gray90"))

        self.load(op2)

    # -------------------------------------------------------------- load
    def load(self, op2):
        """Populate table from OP2 data."""
        self._configure_sheet(list(_SINGLE_HEADERS))
        self.data = None

        if not op2.eigenvalues:
            return

        eigval_table = next(iter(op2.eigenvalues.values()))
        modes = np.array(eigval_table.mode)
        freqs = np.array(eigval_table.cycles)

        meff_frac = op2.modal_effective_mass_fraction
        if meff_frac is None:
            messagebox.showwarning(
                "No MEFFMASS Data",
                "No MEFFMASS matrices found in this OP2.\n\n"
                "Add to your Nastran case control:\n"
                "  MEFFMASS(PLOT) = ALL")
            return

        raw = meff_frac.data
        if scipy.sparse.issparse(raw):
            raw = raw.toarray()
        raw = np.asarray(raw)

        frac = raw.T  # (nmodes, 6)
        cumsum = np.cumsum(frac, axis=0)
        nmodes = min(frac.shape[0], len(modes))

        self.data = {
            'modes': modes[:nmodes],
            'freqs': freqs[:nmodes],
            'frac': frac[:nmodes],
            'cumsum': cumsum[:nmodes],
        }

        self._show_single_view()

    # ---------------------------------------------------------- view helpers
    def _show_single_view(self):
        self._configure_sheet(list(_SINGLE_HEADERS))
        if self.data is None:
            return
        modes, freqs = self.data['modes'], self.data['freqs']
        frac, cumsum = self.data['frac'], self.data['cumsum']
        data = []
        for i in range(len(modes)):
            row = [int(modes[i]), f"{freqs[i]:.1f}"]
            for j in range(6):
                row.extend([f"{frac[i, j]:.2f}", f"{cumsum[i, j]:.2f}"])
            data.append(row)
        self._sheet.set_sheet_data(data)
        self._apply_highlights()

    # ------------------------------------------------------------ export
    def _export_excel(self):
        if self.data is None:
            messagebox.showinfo("Nothing to export",
                                "Open an OP2 file first.")
            return

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

        name_a = os.path.basename(self._op2_path) if self._op2_path else None
        threshold = self._get_threshold()

        wb = Workbook()
        styles = make_meff_styles()
        ws = wb.active
        ws.title = "Effective Mass Fractions"
        write_meff_single_sheet(ws, self.data, styles,
                                op2_name=name_a, threshold=threshold)

        try:
            wb.save(path)
            messagebox.showinfo("Exported", f"Saved to:\n{path}")
        except Exception as exc:
            messagebox.showerror("Export failed", str(exc))


def main():
    import logging
    logging.getLogger("customtkinter").setLevel(logging.ERROR)
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")
    root = ctk.CTk()
    root.title("MEFFMASS")
    root.geometry("1400x600")
    meff = MeffModule(root)
    meff.frame.pack(fill=tk.BOTH, expand=True)
    root.mainloop()


if __name__ == '__main__':
    main()
