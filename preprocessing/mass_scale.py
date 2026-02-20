#!/usr/bin/env python3
"""BDF Mass Scaling Tool.

Standalone GUI that reads a Nastran BDF/DAT file (with includes), shows
mass breakdown by include file, lets the user apply per-group scale factors
(scaling material density, NSM, CONM2 mass & inertia), previews the effect
live, and writes the scaled model back preserving the include file structure.

Usage:
    python mass_scale.py
"""
import copy
import os
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox
from collections import namedtuple, defaultdict

import customtkinter as ctk
import numpy as np
if not hasattr(np, 'in1d'):
    np.in1d = np.isin
from tksheet import Sheet

try:
    from bdf_utils import IncludeFileParser, make_model
except ImportError:
    from preprocessing.bdf_utils import IncludeFileParser, make_model

GroupInfo = namedtuple('GroupInfo', [
    'ifile', 'filename', 'filepath', 'original_mass',
    'material_ids', 'property_ids', 'mass_elem_ids', 'conrod_ids',
])

_NSM_PROP_TYPES = frozenset((
    'PSHELL', 'PCOMP', 'PCOMPG', 'PBAR', 'PBARL',
    'PBEAM', 'PBEAML', 'PROD',
))

_RHO_MAT_TYPES = frozenset(('MAT1', 'MAT8', 'MAT9'))

# Cards irrelevant for mass calculations — safe to skip.
# They are stored as rejected card text and written back out unchanged.
_CARDS_TO_SKIP = [
    'BCPROPS', 'BCTPARM', 'BSURF', 'BSURFS', 'BCPARA', 'BCTSET',
    'BCONP', 'BFRIC', 'BLSEG', 'BOUTPUT', 'BGPARM', 'BGSET',
    'BEDGE', 'BCRPARA', 'BCHANGE', 'BCBODY', 'BCAUTOP',
]


def _extract_card_info(line):
    """Extract card name and primary ID from a raw BDF line.

    Handles fixed-field (8-char or 16-char) and free-field (comma-delimited).
    Returns (name, id) or (None, None) for comments, continuations, blanks.
    """
    stripped = line.strip()
    if not stripped or stripped.startswith('$'):
        return None, None

    first_char = stripped[0]
    if first_char in ('+', '*') or not first_char.isalpha():
        return None, None

    if ',' in stripped:
        fields = stripped.split(',')
        card_name = fields[0].strip().upper()
        id_str = fields[1].strip() if len(fields) > 1 else ''
    else:
        card_name = stripped[:8].strip().upper()
        if card_name.endswith('*'):
            id_str = stripped[8:24].strip() if len(stripped) > 8 else ''
        else:
            id_str = stripped[8:16].strip() if len(stripped) > 8 else ''

    card_name = card_name.rstrip('*')

    try:
        card_id = int(id_str)
    except (ValueError, TypeError):
        return card_name, None

    return card_name, card_id


def _build_scaled_lookup(model, group):
    """Build {(card_type, card_id): card_object} for scaled cards in a group.

    Collects cards from material_ids, property_ids, mass_elem_ids, conrod_ids
    and maps them to their model objects.
    """
    lookup = {}

    for mid in group.material_ids:
        mat = model.materials.get(mid)
        if mat is not None:
            lookup[(mat.type, mid)] = mat

    for pid in group.property_ids:
        prop = model.properties.get(pid)
        if prop is not None:
            lookup[(prop.type, pid)] = prop

    for eid in group.mass_elem_ids:
        mass_elem = model.masses.get(eid)
        if mass_elem is not None:
            lookup[(mass_elem.type, eid)] = mass_elem

    for eid in group.conrod_ids:
        elem = model.elements.get(eid)
        if elem is not None:
            lookup[('CONROD', eid)] = elem

    return lookup


def _rewrite_file_with_scaled_cards(input_path, output_path, scaled_card_lookup,
                                     is_main_file):
    """Rewrite a BDF file, replacing only scaled cards and preserving everything else.

    State machine reads the original file line by line:
    - Before BEGIN BULK (main file only): pass through verbatim
    - In bulk data: when a new card line matches a lookup key, output the
      scaled card via write_card() and skip the original card's lines
    - Comments, blanks, INCLUDE, ENDDATA: pass through verbatim
    """
    # Read entire input (allows overwrite mode where input_path == output_path)
    with open(input_path, 'r', errors='replace') as f:
        lines = f.readlines()

    out = []
    in_bulk = not is_main_file  # include files start in bulk data
    replacing = False  # True while swallowing lines of a replaced card

    for line in lines:
        upper = line.strip().upper()

        # Before BEGIN BULK: pass through verbatim (main file only)
        if not in_bulk:
            out.append(line)
            if upper.startswith('BEGIN') and 'BULK' in upper:
                in_bulk = True
            continue

        # ENDDATA: stop replacing, pass through
        if upper.startswith('ENDDATA'):
            replacing = False
            out.append(line)
            continue

        # INCLUDE: pass through
        if upper.startswith('INCLUDE'):
            replacing = False
            out.append(line)
            continue

        # Comment or blank line
        if not line.strip() or line.strip().startswith('$'):
            if not replacing:
                out.append(line)
            continue

        # Check if this is a new card (first char is alphabetic)
        first_char = line.strip()[0]
        if first_char.isalpha():
            replacing = False  # end any previous replacement
            card_name, card_id = _extract_card_info(line)
            if card_name and card_id is not None:
                key = (card_name, card_id)
                if key in scaled_card_lookup:
                    card = scaled_card_lookup[key]
                    text = card.write_card(size=8)
                    # Strip leading comment line if present (avoids duplication)
                    card_lines = text.split('\n')
                    filtered = []
                    for cl in card_lines:
                        if cl.strip().startswith('$') and not filtered:
                            continue  # skip leading comment
                        filtered.append(cl)
                    text = '\n'.join(filtered)
                    if text and not text.endswith('\n'):
                        text += '\n'
                    out.append(text)
                    replacing = True
                    continue
            # Not a replaced card — pass through
            out.append(line)
        else:
            # Continuation line (starts with +, *, digit, space-then-nonalpha)
            if not replacing:
                out.append(line)
            # else: swallow continuation of a replaced card

    with open(output_path, 'w') as f:
        f.writelines(out)


def _read_wtmass(model):
    """Read WTMASS parameter from model; default 1.0."""
    if 'WTMASS' not in model.params:
        return 1.0
    param = model.params['WTMASS']
    try:
        if hasattr(param, 'values') and param.values:
            return float(param.values[0])
    except (ValueError, TypeError, IndexError):
        pass
    return 1.0


# ---------------------------------------------------------------- Save dialog

class SaveModeDialog(ctk.CTkToplevel):
    """Modal dialog for choosing how to save the scaled BDF."""

    def __init__(self, parent, original_path):
        super().__init__(parent)
        self.title("Write Scaled BDF")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self.result = None
        self._original_path = original_path
        self._mode = tk.StringVar(value='suffix')
        self._suffix = tk.StringVar(value='_scaled')
        self._outdir = tk.StringVar()

        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.wait_window()

    def _build_ui(self):
        pad = {'padx': 10, 'pady': 4}

        f1 = ctk.CTkFrame(self, fg_color="transparent")
        f1.pack(fill=tk.X, **pad)
        ctk.CTkRadioButton(
            f1, text="Add suffix to filenames:",
            variable=self._mode, value='suffix').pack(side=tk.LEFT)
        ctk.CTkEntry(f1, textvariable=self._suffix, width=120).pack(
            side=tk.LEFT, padx=5)

        f2 = ctk.CTkFrame(self, fg_color="transparent")
        f2.pack(fill=tk.X, **pad)
        ctk.CTkRadioButton(
            f2, text="Choose output directory:",
            variable=self._mode, value='directory').pack(side=tk.LEFT)
        ctk.CTkEntry(f2, textvariable=self._outdir, width=220).pack(
            side=tk.LEFT, padx=5)
        ctk.CTkButton(f2, text="Browse\u2026", width=80,
                      command=self._browse_dir).pack(side=tk.LEFT)

        f3 = ctk.CTkFrame(self, fg_color="transparent")
        f3.pack(fill=tk.X, **pad)
        ctk.CTkRadioButton(
            f3, text="Overwrite original files",
            variable=self._mode, value='overwrite').pack(side=tk.LEFT)

        bf = ctk.CTkFrame(self, fg_color="transparent")
        bf.pack(fill=tk.X, padx=10, pady=10)
        ctk.CTkButton(bf, text="Write", command=self._ok).pack(
            side=tk.RIGHT, padx=5)
        ctk.CTkButton(bf, text="Cancel", command=self._cancel).pack(
            side=tk.RIGHT)

    def _browse_dir(self):
        d = filedialog.askdirectory(title="Select output directory")
        if d:
            self._outdir.set(d)
            self._mode.set('directory')

    def _ok(self):
        mode = self._mode.get()

        if mode == 'suffix':
            suffix = self._suffix.get().strip()
            if not suffix:
                messagebox.showwarning(
                    "No suffix", "Please enter a suffix.", parent=self)
                return
            self.result = ('suffix', suffix)
        elif mode == 'directory':
            outdir = self._outdir.get().strip()
            if not outdir:
                messagebox.showwarning(
                    "No directory",
                    "Please choose an output directory.", parent=self)
                return
            self.result = ('directory', outdir)
        elif mode == 'overwrite':
            if not messagebox.askyesno(
                    "Confirm overwrite",
                    "This will overwrite the original BDF files.\n\n"
                    "Are you sure?",
                    parent=self):
                return
            self.result = ('overwrite', None)

        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()


# ---------------------------------------------------------------- Main app

class MassScaleTool(ctk.CTkFrame):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.model = None
        self._bdf_path = None
        self._groups = []
        self._wtmass = 1.0
        self._divide_386 = ctk.BooleanVar(master=self, value=False)
        self._hide_zero = ctk.BooleanVar(master=self, value=False)
        self._scale_overrides = {}
        self._visible_indices = []
        self._sheet = None

        self._build_ui()

    # ------------------------------------------------------------------ UI

    _GUIDE_TEXT = """\
Mass Scale Tool — Quick Guide

PURPOSE
Scale material densities, NSM values, CONM2 masses and inertias per
include file so the total model mass matches target values.

WORKFLOW
1. Open BDF — loads the main BDF and all INCLUDE files.
2. Review the table — each row is an include file with its original mass.
3. Edit Scale Factor (column 3) for any file. The Scaled Mass and Delta
   columns update live.
4. Write Scaled BDF — choose suffix, output directory, or overwrite mode.

OUTPUT MODES
  - Add suffix: appends e.g. "_scaled" to each filename.
  - Output directory: writes the full include tree to a new folder.
  - Overwrite: replaces the original files in place (confirm first).

Only files with scale != 1.0 are rewritten. Unscaled files are skipped.

OPTIONS
  - WTMASS: displayed from the model's PARAM,WTMASS card.
  - Multiply by 386.1: converts mass units (slug -> lbm) for display.
  - Hide zero-mass files: filters out include files with no mass.

DETAIL LABEL
Click a row to see which entity types (MATs, PROPs, Mass Elems, CONRODs)
are present in that file.

SUMMARY FILE
After writing, a scale_summary.md file is saved next to the original BDF
with a table of all files, scales, and entity counts.\
"""

    def _build_ui(self):
        # Toolbar
        toolbar = ctk.CTkFrame(self, fg_color="transparent")
        toolbar.pack(fill=tk.X, padx=5, pady=(5, 0))

        ctk.CTkButton(
            toolbar, text="Open BDF\u2026", width=100,
            command=self._open_bdf).pack(side=tk.LEFT)

        self._path_label = ctk.CTkLabel(
            toolbar, text="No BDF loaded", text_color="gray")
        self._path_label.pack(side=tk.LEFT, padx=(10, 0))

        ctk.CTkButton(
            toolbar, text="?", width=30, font=ctk.CTkFont(weight="bold"),
            command=self._show_guide,
        ).pack(side=tk.RIGHT, padx=(5, 0))

        self._write_btn = ctk.CTkButton(
            toolbar, text="Write Scaled BDF\u2026", width=140,
            command=self._write_scaled, state=tk.DISABLED)
        self._write_btn.pack(side=tk.RIGHT, padx=(0, 5))

        self._reset_btn = ctk.CTkButton(
            toolbar, text="Reset All to 1.0", width=120,
            command=self._reset_all, state=tk.DISABLED)
        self._reset_btn.pack(side=tk.RIGHT)

        # Info bar (WTMASS + *386.1 toggle)
        info_bar = ctk.CTkFrame(self, fg_color="transparent")
        info_bar.pack(fill=tk.X, padx=10, pady=(4, 0))

        self._wtmass_label = ctk.CTkLabel(
            info_bar, text="WTMASS = 1.0000e+00")
        self._wtmass_label.pack(side=tk.LEFT)

        ctk.CTkCheckBox(
            info_bar, text="Multiply displayed masses by 386.1",
            variable=self._divide_386,
            command=self._refresh_display).pack(side=tk.LEFT, padx=(20, 0))

        ctk.CTkCheckBox(
            info_bar, text="Hide zero-mass files",
            variable=self._hide_zero,
            command=self._refresh_display).pack(side=tk.LEFT, padx=(20, 0))

        # Table (tksheet)
        self._sheet = Sheet(
            self,
            headers=["File Name", "Original Mass", "Scale Factor",
                     "Scaled Mass", "Delta"],
            show_top_left=False,
            show_row_index=False,
            height=350,
        )
        self._sheet.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self._sheet.enable_bindings(
            "single_select", "edit_cell", "copy", "paste",
            "arrowkeys", "column_width_resize",
        )
        # Only column 2 (Scale Factor) is editable
        self._sheet.readonly_columns(columns=[0, 1, 3, 4])
        self._sheet.align_columns(columns=[1, 2, 3, 4], align="center",
                                  align_header=True)
        self._sheet.bind("<<SheetModified>>", self._on_sheet_modified)
        self._sheet.extra_bindings("cell_select", self._on_row_select)

        # Detail bar (entity breakdown for selected row)
        self._detail_var = tk.StringVar(value="")
        self._detail_label = ctk.CTkLabel(
            self, textvariable=self._detail_var, text_color="gray",
            anchor="w")
        self._detail_label.pack(fill=tk.X, padx=10, pady=(0, 0))

        # Summary bar
        self._summary_label = ctk.CTkLabel(
            self, text="", font=ctk.CTkFont(weight="bold"))
        self._summary_label.pack(fill=tk.X, padx=10, pady=(0, 5))

    # ---------------------------------------------------------- Guide

    def _show_guide(self):
        """Open the guide dialog (lazy import to avoid circular dependency)."""
        try:
            from nastran_tools import show_guide
        except ImportError:
            return
        show_guide(self.winfo_toplevel(), "Mass Scale Guide", self._GUIDE_TEXT)

    # ---------------------------------------------------------- BDF loading

    def _open_bdf(self):
        path = filedialog.askopenfilename(
            title="Open BDF File",
            filetypes=[("BDF files", "*.bdf *.dat *.nas *.pch"),
                       ("All files", "*.*")])
        if not path:
            return

        self._path_label.configure(
            text=f"Loading {os.path.basename(path)}\u2026",
            text_color=("gray10", "gray90"))
        self.update_idletasks()

        try:
            model = make_model(_CARDS_TO_SKIP)
            model.read_bdf(path)
            model.cross_reference()
        except Exception:
            import traceback
            messagebox.showerror(
                "Error",
                f"Could not read BDF:\n{traceback.format_exc()}")
            self._path_label.configure(
                text="Load failed", text_color="red")
            return

        self.model = model
        self._bdf_path = path
        self._wtmass = _read_wtmass(model)

        self._wtmass_label.configure(
            text=f"WTMASS = {self._wtmass:.4e}")

        self._scale_overrides.clear()
        self._compute_groups()
        self._populate_sheet()

        self._path_label.configure(
            text=os.path.basename(path),
            text_color=("gray10", "gray90"))
        self._write_btn.configure(state=tk.NORMAL)
        self._reset_btn.configure(state=tk.NORMAL)

    # ------------------------------------------------- Mass grouping

    def _build_ifile_lookup(self):
        """Use IncludeFileParser to map card IDs to file indices."""
        parser = IncludeFileParser()
        parser.parse(self._bdf_path)
        filenames = parser.all_files

        eid_to_ifile = {}
        mid_to_ifile = {}
        pid_to_ifile = {}

        for idx, filepath in enumerate(filenames):
            ids_by_type = parser.file_ids.get(filepath, {})
            for eid in ids_by_type.get('eid', set()):
                eid_to_ifile[eid] = idx
            for mid in ids_by_type.get('mid', set()):
                mid_to_ifile[mid] = idx
            for pid in ids_by_type.get('pid', set()):
                pid_to_ifile[pid] = idx

        return filenames, eid_to_ifile, mid_to_ifile, pid_to_ifile

    def _compute_groups(self):
        """Build mass breakdown grouped by include file."""
        model = self.model

        filenames, eid_to_ifile, mid_to_ifile, pid_to_ifile = \
            self._build_ifile_lookup()
        self._include_filenames = filenames

        mass_by_ifile = defaultdict(float)
        mats_by_ifile = defaultdict(set)
        props_by_ifile = defaultdict(set)
        mass_elems_by_ifile = defaultdict(set)
        conrods_by_ifile = defaultdict(set)

        for eid, elem in model.elements.items():
            ifile = eid_to_ifile.get(eid, 0)
            try:
                mass_by_ifile[ifile] += elem.Mass()
            except Exception:
                pass
            if elem.type == 'CONROD':
                conrods_by_ifile[ifile].add(eid)

        for eid, mass_elem in model.masses.items():
            ifile = eid_to_ifile.get(eid, 0)
            try:
                if mass_elem.type == 'CONM2':
                    mass_by_ifile[ifile] += mass_elem.mass
                else:
                    mass_by_ifile[ifile] += mass_elem.Mass()
            except Exception:
                pass
            mass_elems_by_ifile[ifile].add(eid)

        for mid, mat in model.materials.items():
            if mat.type in _RHO_MAT_TYPES:
                rho = getattr(mat, 'rho', None)
                if rho is not None and rho != 0.0:
                    ifile = mid_to_ifile.get(mid, 0)
                    mats_by_ifile[ifile].add(mid)

        for pid, prop in model.properties.items():
            if prop.type in _NSM_PROP_TYPES:
                nsm = getattr(prop, 'nsm', None)
                if nsm is not None and nsm != 0.0:
                    ifile = pid_to_ifile.get(pid, 0)
                    props_by_ifile[ifile].add(pid)

        all_ifiles = set(range(len(filenames)))
        for d in (mass_by_ifile, mats_by_ifile, props_by_ifile,
                  mass_elems_by_ifile, conrods_by_ifile):
            all_ifiles.update(d.keys())

        self._groups = []
        for ifile in sorted(all_ifiles):
            if ifile < len(filenames):
                filepath = filenames[ifile]
                filename = os.path.basename(filepath)
            else:
                filepath = f"<unknown file {ifile}>"
                filename = filepath

            self._groups.append(GroupInfo(
                ifile=ifile,
                filename=filename,
                filepath=filepath,
                original_mass=mass_by_ifile.get(ifile, 0.0),
                material_ids=mats_by_ifile.get(ifile, set()),
                property_ids=props_by_ifile.get(ifile, set()),
                mass_elem_ids=mass_elems_by_ifile.get(ifile, set()),
                conrod_ids=conrods_by_ifile.get(ifile, set()),
            ))

    # ------------------------------------------------- Table population
    #
    # Columns:
    #   0  File Name
    #   1  Original Mass
    #   2  Scale Factor  (editable)
    #   3  Scaled Mass
    #   4  Delta

    def _populate_sheet(self):
        """Fill the tksheet with group data."""
        self._sheet.dehighlight_all(redraw=False)
        # Clear stale cell-level readonly state from previous populate calls
        self._sheet.readonly_columns(columns=[], readonly=False)
        multiplier = 386.1 if self._divide_386.get() else 1.0
        hide_zero = self._hide_zero.get()

        data = []
        self._visible_indices = []
        for i, group in enumerate(self._groups):
            if hide_zero and group.original_mass == 0.0:
                continue
            self._visible_indices.append(i)
            scale = self._scale_overrides.get(i, 1.0)
            orig = group.original_mass * multiplier
            scaled = group.original_mass * scale * multiplier
            delta = (scale - 1.0) * 100 if group.original_mass != 0 else 0.0
            data.append([
                group.filename, f"{orig:.4e}", f"{scale:.4f}",
                f"{scaled:.4e}", f"{delta:+.0f}%",
            ])

        # TOTAL row (always uses ALL groups, including hidden)
        total_orig = sum(g.original_mass for g in self._groups) * multiplier
        total_scaled = sum(
            g.original_mass * self._scale_overrides.get(i, 1.0)
            for i, g in enumerate(self._groups)) * multiplier
        total_delta = ((total_scaled / total_orig - 1.0) * 100
                       if total_orig != 0 else 0.0)
        data.append(["TOTAL", f"{total_orig:.4e}", "",
                     f"{total_scaled:.4e}", f"{total_delta:+.0f}%"])

        self._sheet.set_sheet_data(data)
        self._sheet.readonly_columns(columns=[0, 1, 3, 4])

        total_row = len(self._visible_indices)
        self._sheet.readonly_cells(row=total_row, column=2)
        self._sheet.highlight_rows(rows=[total_row], bg="gray30", fg="white")

        for vi, gi in enumerate(self._visible_indices):
            group = self._groups[gi]
            has_scalable = (group.material_ids or group.property_ids
                           or group.mass_elem_ids or group.conrod_ids
                           or group.original_mass != 0.0)
            if not has_scalable:
                self._sheet.readonly_cells(row=vi, column=2)

        # Gray out zero-mass rows (when visible)
        if not hide_zero:
            zero_rows = [vi for vi, gi in enumerate(self._visible_indices)
                         if self._groups[gi].original_mass == 0.0]
            if zero_rows:
                self._sheet.highlight_rows(rows=zero_rows,
                                           bg="gray25", fg="gray60")

        self._update_summary()

    # --------------------------------------------- Live preview

    def _on_sheet_modified(self, event=None):
        """Called when user edits a cell — sync to _scale_overrides."""
        for vi, gi in enumerate(self._visible_indices):
            try:
                val = float(self._sheet.get_cell_data(vi, 2))
                self._scale_overrides[gi] = val
            except (ValueError, TypeError):
                pass
        self._refresh_display()

    def _refresh_display(self):
        """Rebuild the table from _scale_overrides."""
        if not self._groups:
            return
        self._populate_sheet()

    def _update_summary(self):
        """Update the summary label at the bottom."""
        if not self._groups:
            self._summary_label.configure(text="")
            return

        multiplier = 386.1 if self._divide_386.get() else 1.0
        total_orig = sum(g.original_mass for g in self._groups)
        total_scaled = sum(
            g.original_mass * self._scale_overrides.get(i, 1.0)
            for i, g in enumerate(self._groups))

        disp_orig = total_orig * multiplier
        disp_scaled = total_scaled * multiplier

        self._summary_label.configure(
            text=f"Total: {disp_scaled:.4e}  "
                 f"(original: {disp_orig:.4e})")

    # ------------------------------------------------ Row select detail

    def _on_row_select(self, event=None):
        """Show entity breakdown for the selected row."""
        try:
            row = self._sheet.get_currently_selected().row
        except Exception:
            return
        if row is None or row >= len(self._visible_indices):
            self._detail_var.set("")
            return
        gi = self._visible_indices[row]
        group = self._groups[gi]
        parts = []
        if group.material_ids:
            parts.append(f"{len(group.material_ids)} MATs (rho)")
        if group.property_ids:
            parts.append(f"{len(group.property_ids)} PROPs (nsm)")
        if group.mass_elem_ids:
            parts.append(f"{len(group.mass_elem_ids)} Mass Elems")
        if group.conrod_ids:
            parts.append(f"{len(group.conrod_ids)} CONRODs")
        if parts:
            self._detail_var.set(f"{group.filename} — " + ", ".join(parts))
        else:
            self._detail_var.set(f"{group.filename} — no scalable entities")

    # ------------------------------------------------ Reset

    def _reset_all(self):
        if not self._groups:
            return
        self._scale_overrides.clear()
        self._populate_sheet()

    # ----------------------------------------- Backup / restore originals

    def _capture_originals(self):
        model = self.model
        originals = {}

        for mid, mat in model.materials.items():
            if mat.type in _RHO_MAT_TYPES:
                originals[('mat', mid)] = getattr(mat, 'rho', None)

        for pid, prop in model.properties.items():
            if prop.type in _NSM_PROP_TYPES:
                originals[('prop', pid)] = getattr(prop, 'nsm', None)

        for eid, elem in model.elements.items():
            if elem.type == 'CONROD':
                originals[('conrod', eid)] = getattr(elem, 'nsm', None)

        for eid, mass_elem in model.masses.items():
            if mass_elem.type == 'CONM2':
                I = mass_elem.I
                I_copy = list(I) if I is not None else None
                originals[('conm2', eid)] = (mass_elem.mass, I_copy)
            elif mass_elem.type == 'CONM1':
                mm = getattr(mass_elem, 'mass_matrix', None)
                originals[('conm1', eid)] = copy.deepcopy(mm)
            elif mass_elem.type in ('CMASS1', 'CMASS2'):
                originals[('cmass', eid)] = getattr(mass_elem, 'mass', None)

        return originals

    def _restore_originals(self, originals):
        model = self.model

        for key, val in originals.items():
            kind, card_id = key

            if kind == 'mat':
                model.materials[card_id].rho = val
            elif kind == 'prop':
                model.properties[card_id].nsm = val
            elif kind == 'conrod':
                model.elements[card_id].nsm = val
            elif kind == 'conm2':
                mass_val, I_copy = val
                model.masses[card_id].mass = mass_val
                if I_copy is not None:
                    model.masses[card_id].I = list(I_copy)
            elif kind == 'conm1':
                if val is not None:
                    model.masses[card_id].mass_matrix = copy.deepcopy(val)
            elif kind == 'cmass':
                model.masses[card_id].mass = val

    # --------------------------------------------- Apply scale factors

    def _apply_scale_factors_inplace(self, scale_by_ifile):
        model = self.model

        for group in self._groups:
            scale = scale_by_ifile.get(group.ifile, 1.0)
            if scale == 1.0:
                continue

            for mid in group.material_ids:
                mat = model.materials[mid]
                rho = getattr(mat, 'rho', None)
                if rho is not None and rho != 0.0:
                    mat.rho = rho * scale

            for pid in group.property_ids:
                prop = model.properties[pid]
                nsm = getattr(prop, 'nsm', None)
                if nsm is not None and nsm != 0.0:
                    prop.nsm = nsm * scale

            for eid in group.conrod_ids:
                elem = model.elements[eid]
                nsm = getattr(elem, 'nsm', None)
                if nsm is not None and nsm != 0.0:
                    elem.nsm = nsm * scale

            for eid in group.mass_elem_ids:
                mass_elem = model.masses[eid]
                if mass_elem.type == 'CONM2':
                    mass_elem.mass *= scale
                    if mass_elem.I is not None:
                        mass_elem.I = [x * scale for x in mass_elem.I]
                elif mass_elem.type == 'CONM1':
                    mm = getattr(mass_elem, 'mass_matrix', None)
                    if mm is not None:
                        try:
                            mass_elem.mass_matrix = [
                                [x * scale for x in row] for row in mm]
                        except TypeError:
                            mass_elem.mass_matrix = mm * scale
                elif mass_elem.type in ('CMASS1', 'CMASS2'):
                    m = getattr(mass_elem, 'mass', None)
                    if m is not None:
                        mass_elem.mass = m * scale

    # ------------------------------------------------ Write output

    def _write_summary(self, summary_path, written_files, scales):
        """Write a markdown summary of the scaling operation."""
        lines = ['# Mass Scale Summary', '']
        lines.append(f'**Date:** {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
        lines.append(f'**Original BDF:** {self._bdf_path}')
        lines.append(f'**WTMASS:** {self._wtmass:.4e}')
        lines.append('')

        # Scaled files table
        lines.append('## Scaled Files')
        lines.append('')
        lines.append('| File | Scale | Original Mass | Scaled Mass | Delta'
                     ' | MATs | PROPs | Mass Elems | CONRODs |')
        lines.append('|------|-------|---------------|-------------|------'
                     '|------|-------|------------|---------|')

        total_orig = 0.0
        total_scaled = 0.0
        for group, out_path in written_files:
            scale = scales.get(group.ifile, 1.0)
            orig_mass = group.original_mass
            scaled_mass = orig_mass * scale
            total_orig += orig_mass
            total_scaled += scaled_mass
            if orig_mass != 0:
                delta_pct = (scale - 1.0) * 100.0
                delta_str = f'{delta_pct:+.0f}%'
            else:
                delta_str = 'N/A'
            lines.append(
                f'| {group.filename} | {scale:.4f} '
                f'| {orig_mass:.4e} | {scaled_mass:.4e} | {delta_str} '
                f'| {len(group.material_ids)} | {len(group.property_ids)} '
                f'| {len(group.mass_elem_ids)} | {len(group.conrod_ids)} |')

        lines.append('')
        lines.append(f'**Total Original Mass:** {total_orig:.4e}')
        lines.append(f'**Total Scaled Mass:** {total_scaled:.4e}')
        lines.append('')

        # Entity types breakdown (only for scaled files)
        entity_lines = []
        for group, out_path in written_files:
            scale = scales.get(group.ifile, 1.0)
            if scale == 1.0:
                continue
            parts = []
            if group.material_ids:
                parts.append(f"{len(group.material_ids)} MATs (rho)")
            if group.property_ids:
                parts.append(f"{len(group.property_ids)} PROPs (nsm)")
            if group.mass_elem_ids:
                parts.append(f"{len(group.mass_elem_ids)} Mass Elems")
            if group.conrod_ids:
                parts.append(f"{len(group.conrod_ids)} CONRODs")
            if parts:
                entity_lines.append(
                    f'- **{group.filename}** — ' + ', '.join(parts))

        if entity_lines:
            lines.append('## Scaled Entity Types')
            lines.append('')
            lines.extend(entity_lines)
            lines.append('')

        # Output files list
        lines.append('## Output Files')
        lines.append('')
        for _group, out_path in written_files:
            lines.append(f'- `{out_path}`')
        lines.append('')

        # Unmodified files list
        scaled_ifiles = {g.ifile for g, _ in written_files}
        unmodified = [g for g in self._groups if g.ifile not in scaled_ifiles]
        if unmodified:
            lines.append('## Unmodified Files')
            lines.append('')
            for g in unmodified:
                lines.append(f'- `{g.filename}`')
            lines.append('')

        with open(summary_path, 'w') as f:
            f.write('\n'.join(lines))

    def _write_scaled(self):
        if self.model is None:
            return

        scales = {}
        for i, group in enumerate(self._groups):
            scales[group.ifile] = self._scale_overrides.get(i, 1.0)

        if all(v == 1.0 for v in scales.values()):
            if not messagebox.askyesno(
                    "No scaling",
                    "All scale factors are 1.0 (no changes).\n\n"
                    "Write anyway?"):
                return

        dlg = SaveModeDialog(self.winfo_toplevel(), self._bdf_path)
        if dlg.result is None:
            return
        mode, param = dlg.result

        model = self.model
        filenames = getattr(self, '_include_filenames', None)
        if not filenames:
            filenames = [self._bdf_path]

        out_filenames = {}
        if mode == 'suffix':
            for fp in filenames:
                base, ext = os.path.splitext(fp)
                out_filenames[fp] = f"{base}{param}{ext}"
        elif mode == 'directory':
            main_dir = os.path.dirname(filenames[0])
            for fp in filenames:
                rel = os.path.relpath(fp, main_dir)
                out_filenames[fp] = os.path.join(param, rel)
        elif mode == 'overwrite':
            for fp in filenames:
                out_filenames[fp] = fp

        scaled_fps = {g.filepath for g in self._groups
                      if scales.get(g.ifile, 1.0) != 1.0}
        existing = [p for fp, p in out_filenames.items()
                    if fp in scaled_fps and os.path.exists(p)]
        if existing:
            msg = (f"{len(existing)} output file(s) already exist "
                   "and will be overwritten:\n\n")
            msg += "\n".join(os.path.basename(p) for p in existing[:10])
            if len(existing) > 10:
                msg += f"\n... and {len(existing) - 10} more"
            if not messagebox.askyesno("Confirm overwrite", msg):
                return

        originals = self._capture_originals()
        written_files = []
        total_to_write = sum(
            1 for g in self._groups
            if out_filenames.get(g.filepath) is not None
            and scales.get(g.ifile, 1.0) != 1.0)
        try:
            self._apply_scale_factors_inplace(scales)
            model.uncross_reference()

            file_num = 0
            for i, group in enumerate(self._groups):
                fp = group.filepath
                dst = out_filenames.get(fp)
                if dst is None:
                    continue

                if scales.get(group.ifile, 1.0) == 1.0:
                    continue  # skip unscaled files entirely

                file_num += 1
                self._summary_label.configure(
                    text=f"Writing {group.filename}... "
                         f"({file_num}/{total_to_write})")
                self.update_idletasks()

                out_dir = os.path.dirname(dst)
                if out_dir:
                    os.makedirs(out_dir, exist_ok=True)

                lookup = _build_scaled_lookup(model, group)
                is_main = (i == 0)
                _rewrite_file_with_scaled_cards(fp, dst, lookup, is_main)
                written_files.append((group, dst))

        except Exception as exc:
            messagebox.showerror("Write failed", str(exc))
            return
        finally:
            self._restore_originals(originals)
            try:
                model.cross_reference()
            except Exception:
                pass

        self._update_summary()

        if written_files:
            summary_dir = os.path.dirname(self._bdf_path)
            summary_path = os.path.join(summary_dir, 'scale_summary.md')
            self._write_summary(summary_path, written_files, scales)
            messagebox.showinfo(
                "Success",
                f"Scaled BDF written ({len(written_files)} file(s)).\n"
                f"Summary: {summary_path}")
        else:
            messagebox.showinfo("Success", "No files needed scaling.")


def main():
    import logging
    logging.getLogger("customtkinter").setLevel(logging.ERROR)
    root = ctk.CTk()
    root.title("BDF Mass Scaling Tool")
    root.geometry("1050x500")
    app = MassScaleTool(root)
    app.pack(fill=tk.BOTH, expand=True)
    root.mainloop()


if __name__ == '__main__':
    main()
