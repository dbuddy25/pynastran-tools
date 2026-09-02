"""OP2 Export — dump frequency-response results to MATLAB .mat.

Tick the response types you want and hit Export. Everything the OP2 holds for
them — all entities, DOFs, frequencies and subcases — lands in a .mat file, so
a MATLAB script can plot it with no Python and no pyNastran on that side.

    d  = load('model_export.mat');
    sc = d.acceleration.psd.sc1;
    semilogy(sc.freq, sc.values(:, 1, 3))   % node sc.ids(1), DOF 3

Each subcase struct holds freq (nfreq x 1), ids (nent x 1), dof_labels,
values (nfreq x nent x ndof) and units. RMS has no freq axis; its values are
(nent x ndof). FRF values are complex.

Set "Model units" to match the OP2 (English = in/s^2, in, lbf; SI = m/s^2, m,
N) — it picks the divisor, and getting it wrong scales every value by ~1550x.
PSD is divided by that factor squared, RMS and FRF by the factor itself.
`sc.units` names the unit of the exported values; `sc.input_units` the OP2's.
"""

import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk
import numpy as np

from .asd_common import RESPONSE_TYPES, sc_int as _sc_int, read_op2_for_asd


# (.mat sub-field, RESPONSE_TYPES cfg key, where the dict lives on the OP2)
_FLAVORS = [
    ("psd", "psd_attr", "psd"),
    ("rms", "rms_attr", "rms"),
    ("frf", "frf_attr", "base"),
]

# MATLAB struct field names must be valid identifiers.
_RT_FIELD = {
    "Acceleration": "acceleration",
    "Displacement": "displacement",
    "SPC Force":    "spc_force",
    "CBUSH Force":  "cbush_force",
}

_ROT_LABELS = ["R1 (RX)", "R2 (RY)", "R3 (RZ)"]


def _result_dict(op2, cfg, where, flavor_key):
    """Return the {subcase: table} dict for one response type + flavour."""
    attr = cfg[flavor_key]
    if where == "base":
        return getattr(op2, attr, None) or {}
    container = getattr(getattr(op2, "op2_results", None), where, None)
    return getattr(container, attr, None) or {}


_UNITS_KEY = {"psd": "psd_units", "rms": "rms_units", "frf": "frf_units"}


def _table_to_struct(tbl, cfg, flavor, unit_factor, unit_label):
    """Convert one pyNastran result table into a savemat-ready dict."""
    id_attr = cfg.get("id_attr", "node_gridtype")
    ids = getattr(tbl, id_attr, None)
    if ids is None:
        return None
    ids = np.asarray(ids)
    if id_attr == "node_gridtype" and ids.ndim == 2:
        ids = ids[:, 0]

    data = np.asarray(tbl.data)
    if data.ndim != 3:
        return None

    # PSD is a squared quantity, so the unit factor squares too.
    values = data / (unit_factor ** 2 if flavor == "psd" else unit_factor)

    # RESPONSE_TYPES lists only the translations for node results; the OP2
    # table carries all six, so name the rotations rather than pad DOF4-6.
    ndof = values.shape[2]
    dof_labels = list(cfg["dof_labels"])[:ndof]
    while len(dof_labels) < ndof:
        i = len(dof_labels)
        dof_labels.append(_ROT_LABELS[i - 3] if 3 <= i < 6 else f"DOF{i + 1}")

    out = {"ids": ids.reshape(-1, 1),
           "dof_labels": np.array(dof_labels, dtype=object),
           # The unit of the exported VALUES (g^2/Hz, g, ...), not the input
           # unit the factor came from.
           "units": cfg.get(_UNITS_KEY[flavor], unit_label),
           "input_units": unit_label}
    if flavor == "rms":
        # RMS is one value per entity/DOF; drop the singleton frequency axis.
        out["values"] = values[0] if values.shape[0] == 1 else values
    else:
        out["freq"] = np.asarray(tbl._times, dtype=float).reshape(-1, 1)
        out["values"] = values
    return out


class Op2ExportModule:

    GUIDE = """OP2 EXPORT — MATLAB .mat

Dump frequency-response results to a .mat file so MATLAB can plot them with
no Python and no pyNastran on that side.  Tick the response types you want;
every entity, DOF, frequency and subcase present for them is exported.

Use ASD Overlay when you want a specific plot — this is a bulk dump.

MODEL UNITS
  Must match the OP2: English = in/s^2, in, lbf;  SI = m/s^2, m, N.
  This picks the divisor — wrong here silently scales everything by ~1550x.

IN MATLAB
  d  = load('model_export.mat');
  sc = d.acceleration.psd.sc1;
  semilogy(sc.freq, sc.values(:, 1, 3))   % first node, T3

  sc.ids(1)          % which node that was
  sc.dof_labels{3}   % which DOF
  sc.units           % g^2/Hz

Each subcase has freq, ids, dof_labels, values and units.  PSD/FRF values are
(nfreq x nentity x ndof); RMS drops the frequency axis.  FRF is complex.
"""

    def __init__(self, parent):
        self.frame = ctk.CTkFrame(parent)
        self._op2 = None
        self._op2_path = None
        self._rt_vars = {rt: ctk.BooleanVar(value=(rt == "Acceleration"))
                         for rt in RESPONSE_TYPES}
        # English picks the first unit choice, SI the last (in/s^2 vs m/s^2,
        # in vs m, lbf vs N). Wrong here scales every value by ~1550x.
        self._units_var = ctk.StringVar(value="English")
        self._build_ui()

    def _build_ui(self):
        bar = ctk.CTkFrame(self.frame)
        bar.pack(fill=tk.X, padx=10, pady=(10, 5))

        row = ctk.CTkFrame(bar, fg_color="transparent")
        row.pack(fill=tk.X, pady=4)

        self._open_btn = ctk.CTkButton(row, text="Open OP2", width=110,
                                       command=self._open_op2)
        self._open_btn.pack(side=tk.LEFT, padx=(4, 8))

        self._file_label = ctk.CTkLabel(row, text="(no file)", text_color="gray")
        self._file_label.pack(side=tk.LEFT)

        ctk.CTkButton(row, text="?", width=28,
                      command=self._show_guide).pack(side=tk.RIGHT, padx=4)

        self._export_btn = ctk.CTkButton(row, text="Export .mat", width=120,
                                         command=self._export, state=tk.DISABLED)
        self._export_btn.pack(side=tk.RIGHT, padx=4)

        rt_row = ctk.CTkFrame(bar, fg_color="transparent")
        rt_row.pack(fill=tk.X, pady=(0, 6))
        ctk.CTkLabel(rt_row, text="Export:").pack(side=tk.LEFT, padx=(4, 10))
        for rt in RESPONSE_TYPES:
            ctk.CTkCheckBox(rt_row, text=rt, variable=self._rt_vars[rt],
                            width=120).pack(side=tk.LEFT, padx=6)

        ctk.CTkLabel(rt_row, text="Model units:").pack(side=tk.LEFT, padx=(18, 4))
        ctk.CTkOptionMenu(rt_row, variable=self._units_var,
                          values=["English", "SI"], width=90).pack(side=tk.LEFT)

        self._status_label = ctk.CTkLabel(self.frame, text="Open an OP2 to begin.",
                                          text_color="gray", anchor="w")
        self._status_label.pack(fill=tk.X, padx=14, pady=(0, 10))

    def _show_guide(self):
        try:
            from structures_tools import show_guide
            show_guide(self.frame, "OP2 Export", self.GUIDE,
                       font=ctk.CTkFont(family="Courier", size=12))
        except Exception:
            messagebox.showinfo("OP2 Export", self.GUIDE)

    def _run_in_background(self, label, work_fn, done_fn):
        self._status_label.configure(text=label, text_color="gray")
        self._open_btn.configure(state=tk.DISABLED)
        self._export_btn.configure(state=tk.DISABLED)
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
                if self._op2 is not None:
                    self._export_btn.configure(state=tk.NORMAL)
                done_fn(container.get('result'), container.get('error'))

        thread = threading.Thread(target=_worker, daemon=True)
        thread.start()
        self.frame.after(50, _poll)

    def _open_op2(self):
        path = filedialog.askopenfilename(
            title="Open OP2",
            filetypes=[("OP2 files", "*.op2"), ("All files", "*.*")])
        if not path:
            return

        def _done(op2, error):
            if error is not None:
                messagebox.showerror("Error", f"Could not read OP2:\n{error}")
                self._status_label.configure(text="Load failed", text_color="red")
                return
            self._op2 = op2
            self._op2_path = path
            self._file_label.configure(text=os.path.basename(path),
                                       text_color=("gray10", "gray90"))
            self._export_btn.configure(state=tk.NORMAL)
            self._status_label.configure(text="Loaded. Pick response types and export.",
                                         text_color=("gray10", "gray90"))

        self._run_in_background("Loading OP2…", lambda: read_op2_for_asd(path), _done)

    def _build_export(self):
        """Assemble the nested dict handed to savemat."""
        data = {}
        for rt, cfg in RESPONSE_TYPES.items():
            if not self._rt_vars[rt].get():
                continue
            choices = cfg["unit_choices"]
            unit_label = choices[0] if self._units_var.get() == "English" else choices[-1]
            unit_factor = cfg["unit_factors"][unit_label]
            rt_out = {}
            for fkey, cfg_key, where in _FLAVORS:
                per_sc = {}
                for key, tbl in _result_dict(self._op2, cfg, where, cfg_key).items():
                    struct = _table_to_struct(tbl, cfg, fkey, unit_factor, unit_label)
                    if struct is None:
                        continue
                    # Random-result keys are tuples; sc_int collapses them, so
                    # two tables can land on one field. Keep both.
                    field = f"sc{_sc_int(key)}"
                    if field in per_sc:
                        n = 2
                        while f"{field}_{n}" in per_sc:
                            n += 1
                        field = f"{field}_{n}"
                    per_sc[field] = struct
                if per_sc:
                    rt_out[fkey] = per_sc
            if rt_out:
                data[_RT_FIELD[rt]] = rt_out
        return data

    def _export(self):
        if self._op2 is None:
            return
        if not any(v.get() for v in self._rt_vars.values()):
            messagebox.showwarning(
                "Nothing to Export", "No response types ticked.")
            return

        stem = os.path.splitext(os.path.basename(self._op2_path))[0]
        path = filedialog.asksaveasfilename(
            title="Export MATLAB .mat", defaultextension=".mat",
            filetypes=[("MATLAB files", "*.mat")],
            initialfile=f"{stem}_export.mat")
        if not path:
            return

        def _work():
            # Assembling copies every result array — keep it off the GUI
            # thread, and do it only once the user has committed to a path.
            data = self._build_export()
            if not data:
                return None
            from scipy.io import savemat
            savemat(path, data, do_compression=True, oned_as="column")
            return path

        def _done(result, error):
            if error is not None:
                messagebox.showerror("Export Error", str(error))
                self._status_label.configure(text="Export failed", text_color="red")
                return
            if result is None:
                messagebox.showwarning(
                    "Nothing to Export",
                    "The OP2 holds none of the ticked response types.")
                self._status_label.configure(text="Nothing exported",
                                             text_color="gray")
                return
            self._status_label.configure(
                text=f"Wrote {os.path.basename(result)} "
                     f"({os.path.getsize(result) / 1024 ** 2:.1f} MB)",
                text_color=("gray10", "gray90"))
            messagebox.showinfo("Saved", os.path.basename(result))

        self._run_in_background("Writing .mat…", _work, _done)
