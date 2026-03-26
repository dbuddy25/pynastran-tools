#!/usr/bin/env python3
"""Thermal CTE Check Tool.

Reads a Nastran BDF, displays all material CTEs, and writes a modified
copy with uniform CTE for the classic same-CTE thermal excursion check.

A model with uniform CTE under uniform temperature should expand freely
with zero stress — any stress indicates thermal-elastic coupling errors.

Usage:
    python thermal_cte.py
"""
import os
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk
import numpy as np
if not hasattr(np, 'in1d'):
    np.in1d = np.isin
from tksheet import Sheet

try:
    from bdf_utils import make_model
except ImportError:
    from preprocessing.bdf_utils import make_model


# Material types that carry CTE fields
_CTE_MAT_TYPES = frozenset(('MAT1', 'MAT2', 'MAT8', 'MAT9', 'MAT11'))

# Cards safe to skip (contact, etc.)
_CARDS_TO_SKIP = [
    'BCPROPS', 'BCTPARM', 'BSURF', 'BSURFS', 'BCPARA', 'BCTSET',
    'BCONP', 'BFRIC', 'BLSEG', 'BOUTPUT', 'BGPARM', 'BGSET',
    'BEDGE', 'BCRPARA', 'BCHANGE', 'BCBODY', 'BCAUTOP',
]


# ---------------------------------------------------------------------------
# CTE helpers
# ---------------------------------------------------------------------------

def _get_cte_values(mat):
    """Extract CTE values from a material card as an ordered dict."""
    mtype = mat.type
    if mtype == 'MAT1':
        return {'a': mat.a}
    elif mtype == 'MAT2':
        return {'a1': mat.a1, 'a2': mat.a2, 'a3': mat.a3}
    elif mtype == 'MAT8':
        return {'a1': mat.a1, 'a2': mat.a2}
    elif mtype == 'MAT9':
        A = mat.A if mat.A is not None else [0.0] * 6
        return {'a1': A[0], 'a2': A[1], 'a3': A[2],
                'a12': A[3], 'a23': A[4], 'a13': A[5]}
    elif mtype == 'MAT11':
        return {'a1': mat.a1, 'a2': mat.a2, 'a3': mat.a3}
    return {}


def _get_tref(mat):
    """Extract reference temperature from a material card."""
    return getattr(mat, 'tref', 0.0) or 0.0


def _set_uniform_cte(mat, target):
    """Set all normal CTE components to *target*; coupling terms to 0."""
    mtype = mat.type
    if mtype == 'MAT1':
        mat.a = target
    elif mtype == 'MAT2':
        mat.a1 = target
        mat.a2 = target
        mat.a3 = 0.0          # coupling
    elif mtype == 'MAT8':
        mat.a1 = target
        mat.a2 = target
    elif mtype == 'MAT9':
        mat.A = [target, target, target, 0.0, 0.0, 0.0]
    elif mtype == 'MAT11':
        mat.a1 = target
        mat.a2 = target
        mat.a3 = target


def _restore_cte(mat, cte):
    """Restore original CTE values from a saved dict."""
    mtype = mat.type
    if mtype == 'MAT1':
        mat.a = cte.get('a', 0.0)
    elif mtype == 'MAT2':
        mat.a1 = cte.get('a1', 0.0)
        mat.a2 = cte.get('a2', 0.0)
        mat.a3 = cte.get('a3', 0.0)
    elif mtype == 'MAT8':
        mat.a1 = cte.get('a1', 0.0)
        mat.a2 = cte.get('a2', 0.0)
    elif mtype == 'MAT9':
        mat.A = [cte.get('a1', 0.0), cte.get('a2', 0.0),
                 cte.get('a3', 0.0), cte.get('a12', 0.0),
                 cte.get('a23', 0.0), cte.get('a13', 0.0)]
    elif mtype == 'MAT11':
        mat.a1 = cte.get('a1', 0.0)
        mat.a2 = cte.get('a2', 0.0)
        mat.a3 = cte.get('a3', 0.0)


def _set_tref(mat, val):
    """Set reference temperature if the material supports it."""
    if hasattr(mat, 'tref'):
        mat.tref = val


# ---------------------------------------------------------------------------
# Formatting
# ---------------------------------------------------------------------------

def _fmt(val):
    if val is None or val == 0.0:
        return "0.0"
    return f"{val:.4e}"


def _fmt_tref(val):
    if val is None or val == 0.0:
        return "0.0"
    return f"{val:.1f}"


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

class ThermalCteTool(ctk.CTkFrame):
    """GUI tool for uniform CTE thermal excursion check."""

    _GUIDE_TEXT = """\
Thermal CTE Check Tool
=======================

PURPOSE
This tool prepares a "same CTE check" model -- a standard
verification that the thermal-elastic coupling in a Nastran
model is set up correctly.

HOW IT WORKS
A model with uniform CTE under a uniform temperature delta
should expand freely with ZERO stress. Any non-zero stress
indicates an error in the thermal setup (wrong material
assignments, inconsistent TREFs, bad coordinate systems, etc.).

WORKFLOW
1. Open your BDF model
2. Review the current CTE values for all materials
3. Enter your desired uniform CTE value
4. Optionally set a uniform TREF
5. Click "Write BDF..." to save the modified model
6. Run the modified model with a uniform temperature load
7. Check for zero stress in the results

SUPPORTED MATERIALS
  MAT1  -- isotropic (a)
  MAT2  -- anisotropic 2D (a1, a2; a3 set to 0)
  MAT8  -- orthotropic shell (a1, a2)
  MAT9  -- anisotropic 3D (a1, a2, a3; coupling -> 0)
  MAT11 -- orthotropic 3D (a1, a2, a3)

TABLE HIGHLIGHTING
  Yellow rows = material has all CTEs at 0.0 (no thermal
  expansion defined -- may be intentional or an oversight)

NOTES
  - Coupling/shear CTE terms are set to 0 for a proper
    isotropic expansion check.
  - Output is a flat BDF (includes merged) -- this is a
    check model, not a replacement for your production files.
  - Remember to add TEMP/TEMPD load cards for the thermal
    excursion (e.g. delta-T = 1 degree).
  - Temperature-dependent materials (MATT1, etc.) are NOT
    modified -- remove those references manually if needed.
"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.model = None
        self._bdf_path = None
        self._mat_info = []          # list of dicts per material
        self._original_ctes = {}     # {mid: {'cte': {...}, 'tref': float}}

        self._target_cte_var = tk.StringVar(value='1.0e-5')
        self._target_tref_var = tk.StringVar(value='0.0')
        self._set_tref_var = ctk.BooleanVar(master=self, value=False)

        self._build_ui()

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        # --- Toolbar ---
        toolbar = ctk.CTkFrame(self, fg_color="transparent")
        toolbar.pack(fill=tk.X, padx=5, pady=(5, 0))

        self._open_btn = ctk.CTkButton(
            toolbar, text="Open BDF\u2026", width=100, command=self._open_bdf)
        self._open_btn.pack(side=tk.LEFT)

        self._path_label = ctk.CTkLabel(
            toolbar, text="No BDF loaded", text_color="gray")
        self._path_label.pack(side=tk.LEFT, padx=(10, 0))

        ctk.CTkButton(
            toolbar, text="?", width=30,
            font=ctk.CTkFont(weight="bold"),
            command=self._show_guide,
        ).pack(side=tk.RIGHT, padx=(5, 0))

        self._write_btn = ctk.CTkButton(
            toolbar, text="Write BDF\u2026", width=120,
            command=self._write_bdf, state=tk.DISABLED)
        self._write_btn.pack(side=tk.RIGHT, padx=(5, 0))

        # --- Options bar ---
        opts = ctk.CTkFrame(self, fg_color="transparent")
        opts.pack(fill=tk.X, padx=10, pady=(8, 0))

        ctk.CTkLabel(opts, text="Target CTE:").pack(side=tk.LEFT)
        self._cte_entry = ctk.CTkEntry(
            opts, textvariable=self._target_cte_var, width=120)
        self._cte_entry.pack(side=tk.LEFT, padx=(5, 20))

        self._tref_check = ctk.CTkCheckBox(
            opts, text="Set TREF:", variable=self._set_tref_var,
            command=self._toggle_tref)
        self._tref_check.pack(side=tk.LEFT)
        self._tref_entry = ctk.CTkEntry(
            opts, textvariable=self._target_tref_var, width=100,
            state=tk.DISABLED, placeholder_text="0.0")
        self._tref_entry.pack(side=tk.LEFT, padx=(5, 0))

        # --- Table ---
        headers = ["MID", "Type", "CTE-1", "CTE-2", "CTE-3", "TREF"]
        self._sheet = Sheet(
            self, headers=headers,
            show_top_left=False, show_row_index=False, height=400,
        )
        self._sheet.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self._sheet.disable_bindings()
        self._sheet.enable_bindings(
            "single_select", "copy", "arrowkeys",
            "column_width_resize", "row_height_resize",
        )
        self._sheet.readonly_columns(columns=list(range(len(headers))))

        # --- Detail + Summary ---
        self._detail_label = ctk.CTkLabel(
            self, text="", text_color="gray", anchor=tk.W)
        self._detail_label.pack(fill=tk.X, padx=10, pady=(0, 0))

        self._summary_label = ctk.CTkLabel(
            self, text="", font=ctk.CTkFont(size=12))
        self._summary_label.pack(fill=tk.X, padx=10, pady=(0, 5))

        # Row-select binding for detail
        self._sheet.extra_bindings("cell_select", self._on_row_select)

    def _toggle_tref(self):
        state = tk.NORMAL if self._set_tref_var.get() else tk.DISABLED
        self._tref_entry.configure(state=state)

    def _show_guide(self):
        try:
            from nastran_tools import show_guide
        except ImportError:
            return
        show_guide(self.winfo_toplevel(), "Thermal CTE Guide", self._GUIDE_TEXT)

    # -------------------------------------------------------------- Open BDF
    def _open_bdf(self):
        path = filedialog.askopenfilename(
            title="Open BDF File",
            filetypes=[("BDF files", "*.bdf *.dat *.nas *.pch"),
                       ("All files", "*.*")])
        if not path:
            return

        self._path_label.configure(text=f"Loading {os.path.basename(path)}\u2026",
                                   text_color="gray")
        self._open_btn.configure(state=tk.DISABLED)
        self.update_idletasks()

        try:
            model = make_model(_CARDS_TO_SKIP)
            model.read_bdf(path)
        except Exception:
            import traceback
            messagebox.showerror(
                "Error", f"Could not read BDF:\n{traceback.format_exc()}")
            self._path_label.configure(text="Load failed", text_color="red")
            self._open_btn.configure(state=tk.NORMAL)
            return

        self.model = model
        self._bdf_path = path
        self._extract_materials()
        self._populate_sheet()
        self._path_label.configure(
            text=os.path.basename(path), text_color=("gray10", "gray90"))
        self._open_btn.configure(state=tk.NORMAL)
        self._write_btn.configure(state=tk.NORMAL)

    # --------------------------------------------------------- Extract data
    def _extract_materials(self):
        """Scan model.materials for cards with CTE fields."""
        self._mat_info = []
        self._original_ctes = {}

        for mid in sorted(self.model.materials.keys()):
            mat = self.model.materials[mid]
            if mat.type not in _CTE_MAT_TYPES:
                continue

            cte = _get_cte_values(mat)
            tref = _get_tref(mat)

            self._original_ctes[mid] = {'cte': cte.copy(), 'tref': tref}
            self._mat_info.append({
                'mid': mid,
                'type': mat.type,
                'cte': cte,
                'tref': tref,
            })

    # --------------------------------------------------------- Populate table
    def _populate_sheet(self):
        self._sheet.dehighlight_all(redraw=False)
        data = []

        for info in self._mat_info:
            mid = info['mid']
            mtype = info['type']
            cte = info['cte']
            tref = info['tref']

            # Build row: MID, Type, CTE-1, CTE-2, CTE-3, TREF
            c1 = c2 = c3 = ""
            if mtype == 'MAT1':
                c1 = _fmt(cte.get('a', 0.0))
            elif mtype == 'MAT2':
                c1 = _fmt(cte.get('a1', 0.0))
                c2 = _fmt(cte.get('a2', 0.0))
                c3 = _fmt(cte.get('a3', 0.0))
            elif mtype == 'MAT8':
                c1 = _fmt(cte.get('a1', 0.0))
                c2 = _fmt(cte.get('a2', 0.0))
            elif mtype == 'MAT9':
                c1 = _fmt(cte.get('a1', 0.0))
                c2 = _fmt(cte.get('a2', 0.0))
                c3 = _fmt(cte.get('a3', 0.0))
            elif mtype == 'MAT11':
                c1 = _fmt(cte.get('a1', 0.0))
                c2 = _fmt(cte.get('a2', 0.0))
                c3 = _fmt(cte.get('a3', 0.0))

            data.append([str(mid), mtype, c1, c2, c3, _fmt_tref(tref)])

        self._sheet.set_sheet_data(data)
        self._sheet.readonly_columns(columns=list(range(6)))
        self._sheet.align_columns(list(range(6)),
                                  align="center", align_header=True)

        # Highlight materials with all-zero CTE (possible oversight)
        for i, info in enumerate(self._mat_info):
            vals = info['cte'].values()
            if all((v or 0.0) == 0.0 for v in vals):
                self._sheet.highlight_rows(
                    rows=[i], bg="#4a3000", fg="#ffcc00")

        # Summary
        n = len(self._mat_info)
        types = sorted(set(info['type'] for info in self._mat_info))
        self._summary_label.configure(
            text=f"{n} material(s) with CTE fields  ({', '.join(types)})")
        self._detail_label.configure(text="")

    # ---------------------------------------------------------- Row select
    def _on_row_select(self, event=None):
        try:
            row = self._sheet.get_currently_selected().row
        except Exception:
            return
        if row is None or row >= len(self._mat_info):
            return

        info = self._mat_info[row]
        mtype = info['type']
        cte = info['cte']

        # Show full CTE breakdown for MAT9 (has coupling terms)
        if mtype == 'MAT9':
            detail = (f"MAT9 MID={info['mid']}:  "
                      f"a1={_fmt(cte.get('a1'))}  a2={_fmt(cte.get('a2'))}  "
                      f"a3={_fmt(cte.get('a3'))}  "
                      f"a12={_fmt(cte.get('a12'))}  a23={_fmt(cte.get('a23'))}  "
                      f"a13={_fmt(cte.get('a13'))}")
        else:
            parts = [f"{k}={_fmt(v)}" for k, v in cte.items()]
            detail = f"{mtype} MID={info['mid']}:  {'  '.join(parts)}"
        self._detail_label.configure(text=detail)

    # --------------------------------------------------------------- Write
    def _write_bdf(self):
        if self.model is None:
            return

        # Validate target CTE
        try:
            target_cte = float(self._target_cte_var.get())
        except ValueError:
            messagebox.showerror("Invalid CTE",
                                 "Enter a valid number for Target CTE.")
            return

        # Validate target TREF if enabled
        set_tref = self._set_tref_var.get()
        target_tref = None
        if set_tref:
            try:
                target_tref = float(self._target_tref_var.get())
            except ValueError:
                messagebox.showerror("Invalid TREF",
                                     "Enter a valid number for TREF.")
                return

        # Output path
        base, ext = os.path.splitext(self._bdf_path)
        default_name = f"{os.path.basename(base)}_uniform_cte{ext}"
        out_path = filedialog.asksaveasfilename(
            title="Save Uniform-CTE BDF",
            initialfile=default_name,
            initialdir=os.path.dirname(self._bdf_path),
            filetypes=[("BDF files", "*.bdf *.dat *.nas"),
                       ("All files", "*.*")])
        if not out_path:
            return

        # Apply uniform CTE (and optionally TREF) to all CTE materials
        for mid, mat in self.model.materials.items():
            if mat.type not in _CTE_MAT_TYPES:
                continue
            _set_uniform_cte(mat, target_cte)
            if set_tref and target_tref is not None:
                _set_tref(mat, target_tref)

        # Write the modified model
        try:
            self.model.write_bdf(out_path, size=8, is_double=False)
        except Exception:
            import traceback
            messagebox.showerror(
                "Write Error",
                f"Could not write BDF:\n{traceback.format_exc()}")
            self._restore_originals()
            return

        # Restore original values so display/re-export stays correct
        self._restore_originals()

        tref_msg = f"\nTREF = {target_tref}" if set_tref else ""
        messagebox.showinfo(
            "Success",
            f"Wrote uniform-CTE model to:\n{os.path.basename(out_path)}\n\n"
            f"CTE = {target_cte:.6e}{tref_msg}")

    def _restore_originals(self):
        """Put original CTE/TREF values back on the model."""
        for mid, orig in self._original_ctes.items():
            mat = self.model.materials.get(mid)
            if mat is None:
                continue
            _restore_cte(mat, orig['cte'])
            _set_tref(mat, orig['tref'])


# ------------------------------------------------------------------- main
if __name__ == '__main__':
    import logging
    logging.getLogger("customtkinter").setLevel(logging.ERROR)

    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")

    app = ctk.CTk()
    app.title("Thermal CTE Check")
    app.geometry("1000x600")

    tool = ThermalCteTool(app)
    tool.pack(fill=tk.BOTH, expand=True)
    app.mainloop()
