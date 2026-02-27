#!/usr/bin/env python3
"""BDF Partitioner — Standalone GUI.

Reads a monolithic Nastran BDF, partitions it into component-level include files
using flood-fill (boundaries at RBE2-CBUSH-RBE2 interfaces and glue contact),
shows a pyvista 3D preview colored by part, and writes organized include files.

Usage:
    python partition_gui.py
"""
import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk
import numpy as np
if not hasattr(np, 'in1d'):
    np.in1d = np.isin
from tksheet import Sheet

try:
    from bdf_utils import make_model
    from partition_bdf import (
        partition_model, merge_parts, write_partition,
        build_pyvista_mesh, show_partition_preview,
        _CARDS_TO_SKIP,
    )
except ImportError:
    from preprocessing.bdf_utils import make_model
    from preprocessing.partition_bdf import (
        partition_model, merge_parts, write_partition,
        build_pyvista_mesh, show_partition_preview,
        _CARDS_TO_SKIP,
    )


# ── Guide text ─────────────────────────────────────────────────────────────

_GUIDE_TEXT = """\
BDF Partitioner — Quick Guide

PURPOSE
Split a monolithic Nastran BDF into component-level include files.
Components connected by RBE2-CBUSH-RBE2 chains or glue contact are
detected automatically.

WORKFLOW
1. Open BDF — loads and cross-references the model.
2. Click Partition — runs flood-fill to detect parts and joints.
3. Review the parts table — rename parts by editing the Name column.
4. (Optional) Select 2+ parts and click Merge Selected to combine them.
5. (Optional) Click 3D Preview to visualize the partition in pyvista.
6. Set the Output Directory and click Write Include Files.

OUTPUT
  master.bdf       — exec/case control + INCLUDE statements
  shared.bdf       — materials, properties, coordinate systems
  <part>.bdf       — GRIDs, elements, mass elements, SPCs, loads
  <partA>-to-<partB>.bdf — boundary CBUSH + RBE2 pairs + PBUSH

MERGE
Select multiple rows in the table and click Merge Selected.
Joints between merged parts are absorbed (CBUSHes become interior).

3D PREVIEW
Requires pyvista (pip install pyvista). If not installed, the button
is disabled. Parts are colored categorically with edge display.\
"""


class PartitionTool(ctk.CTkFrame):
    def __init__(self, parent=None):
        super().__init__(parent)

        self._model = None
        self._bdf_path = None
        self._result = None      # PartitionResult
        self._mesh = None        # pyvista mesh
        self._pyvista_ok = None  # None = not checked, True/False

        self._build_ui()

    # ------------------------------------------------------------------ UI

    def _build_ui(self):
        # ── Top toolbar ──
        toolbar = ctk.CTkFrame(self, fg_color="transparent")
        toolbar.pack(fill=tk.X, padx=5, pady=(5, 0))

        ctk.CTkLabel(toolbar, text="Input BDF:").pack(side=tk.LEFT)
        self._path_var = tk.StringVar()
        self._path_entry = ctk.CTkEntry(
            toolbar, textvariable=self._path_var, width=400)
        self._path_entry.pack(side=tk.LEFT, padx=(5, 0))

        ctk.CTkButton(
            toolbar, text="Browse\u2026", width=80,
            command=self._browse_bdf).pack(side=tk.LEFT, padx=5)

        self._partition_btn = ctk.CTkButton(
            toolbar, text="Partition", width=100,
            command=self._do_partition, state=tk.DISABLED)
        self._partition_btn.pack(side=tk.LEFT)

        ctk.CTkButton(
            toolbar, text="?", width=30, font=ctk.CTkFont(weight="bold"),
            command=self._show_guide,
        ).pack(side=tk.RIGHT, padx=(5, 0))

        # ── Status bar ──
        self._status_label = ctk.CTkLabel(
            self, text="Ready", text_color="gray", anchor="w")
        self._status_label.pack(fill=tk.X, padx=10, pady=(4, 0))

        # ── Parts table (tksheet) ──
        self._sheet = Sheet(
            self,
            headers=["#", "Name", "Elems", "Nodes", "PIDs"],
            show_top_left=False,
            show_row_index=False,
            height=300,
        )
        self._sheet.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self._sheet.enable_bindings(
            "single_select", "row_select", "edit_cell", "copy",
            "arrowkeys", "column_width_resize",
        )
        # Only Name column (1) is editable
        self._sheet.readonly_columns(columns=[0, 2, 3, 4])
        self._sheet.align_columns(
            columns=[0, 2, 3, 4], align="center", align_header=True)
        self._sheet.bind("<<SheetModified>>", self._on_name_edited)

        # ── Joints summary ──
        self._joints_label = ctk.CTkLabel(
            self, text="", text_color="gray", anchor="w")
        self._joints_label.pack(fill=tk.X, padx=10, pady=(0, 2))

        # ── Output bar ──
        out_bar = ctk.CTkFrame(self, fg_color="transparent")
        out_bar.pack(fill=tk.X, padx=5, pady=(0, 2))

        ctk.CTkLabel(out_bar, text="Output Dir:").pack(side=tk.LEFT)
        self._outdir_var = tk.StringVar()
        ctk.CTkEntry(
            out_bar, textvariable=self._outdir_var, width=400,
        ).pack(side=tk.LEFT, padx=(5, 0))
        ctk.CTkButton(
            out_bar, text="Browse\u2026", width=80,
            command=self._browse_outdir).pack(side=tk.LEFT, padx=5)

        # ── Action buttons ──
        btn_bar = ctk.CTkFrame(self, fg_color="transparent")
        btn_bar.pack(fill=tk.X, padx=5, pady=(0, 5))

        self._preview_btn = ctk.CTkButton(
            btn_bar, text="3D Preview", width=110,
            command=self._show_preview, state=tk.DISABLED)
        self._preview_btn.pack(side=tk.LEFT)

        self._merge_btn = ctk.CTkButton(
            btn_bar, text="Merge Selected", width=130,
            command=self._merge_selected, state=tk.DISABLED)
        self._merge_btn.pack(side=tk.LEFT, padx=10)

        self._write_btn = ctk.CTkButton(
            btn_bar, text="Write Include Files", width=160,
            command=self._write_output, state=tk.DISABLED)
        self._write_btn.pack(side=tk.RIGHT)

    # ---------------------------------------------------------- Guide

    def _show_guide(self):
        try:
            from nastran_tools import show_guide
        except ImportError:
            # Fallback: simple messagebox
            messagebox.showinfo("BDF Partitioner Guide", _GUIDE_TEXT)
            return
        show_guide(self.winfo_toplevel(), "BDF Partitioner Guide", _GUIDE_TEXT)

    # ---------------------------------------------------------- Threading

    def _run_in_background(self, label, work_fn, done_fn):
        """Run work_fn in a background thread, keeping UI responsive."""
        self._status_label.configure(text=label, text_color="gray")
        self._partition_btn.configure(state=tk.DISABLED)
        self._write_btn.configure(state=tk.DISABLED)
        self._merge_btn.configure(state=tk.DISABLED)
        self._preview_btn.configure(state=tk.DISABLED)

        container = {}

        def _worker():
            try:
                container['result'] = work_fn()
            except Exception as exc:
                container['error'] = exc

        def _poll():
            if thread.is_alive():
                self.after(50, _poll)
            else:
                done_fn(container.get('result'), container.get('error'))

        thread = threading.Thread(target=_worker, daemon=True)
        thread.start()
        self.after(50, _poll)

    def _restore_buttons(self):
        """Re-enable buttons based on current state."""
        has_bdf = self._bdf_path is not None
        has_result = self._result is not None

        self._partition_btn.configure(
            state=tk.NORMAL if has_bdf else tk.DISABLED)
        self._write_btn.configure(
            state=tk.NORMAL if has_result else tk.DISABLED)
        self._merge_btn.configure(
            state=tk.NORMAL if has_result else tk.DISABLED)

        # Check pyvista availability
        if self._pyvista_ok is None:
            try:
                import pyvista  # noqa: F401
                self._pyvista_ok = True
            except ImportError:
                self._pyvista_ok = False

        self._preview_btn.configure(
            state=tk.NORMAL if (has_result and self._pyvista_ok) else tk.DISABLED)

    # ---------------------------------------------------------- BDF loading

    def _browse_bdf(self):
        path = filedialog.askopenfilename(
            title="Open BDF File",
            filetypes=[("BDF files", "*.bdf *.dat *.nas *.pch"),
                       ("All files", "*.*")])
        if not path:
            return
        self._path_var.set(path)
        self._bdf_path = path
        self._result = None
        self._mesh = None
        self._sheet.set_sheet_data([])
        self._joints_label.configure(text="")

        # Set default output dir
        base_dir = os.path.dirname(path)
        base_name = os.path.splitext(os.path.basename(path))[0]
        self._outdir_var.set(os.path.join(base_dir, base_name + '_partitioned'))

        self._restore_buttons()
        self._status_label.configure(
            text=f"Loaded path: {os.path.basename(path)}. Click Partition.",
            text_color=("gray10", "gray90"))

    # ---------------------------------------------------------- Partition

    def _do_partition(self):
        path = self._bdf_path
        if not path:
            return

        def _work():
            model = make_model(_CARDS_TO_SKIP)
            model.read_bdf(path)
            model.cross_reference()
            result = partition_model(model)
            return model, result

        def _done(res, error):
            if error:
                import traceback
                messagebox.showerror(
                    "Partition Error",
                    f"Could not partition BDF:\n{traceback.format_exc()}")
                self._status_label.configure(
                    text="Partition failed", text_color="red")
                self._restore_buttons()
                return

            model, result = res
            self._model = model
            self._result = result
            self._mesh = None  # invalidate cached mesh

            self._populate_table()
            self._update_joints_label()

            warn_text = ""
            if result.warnings:
                warn_text = f" ({len(result.warnings)} warning(s))"
            self._status_label.configure(
                text=f"Found {len(result.parts)} parts, "
                     f"{len(result.joints)} joints{warn_text}",
                text_color=("gray10", "gray90"))

            self._restore_buttons()

        self._run_in_background(
            f"Partitioning {os.path.basename(path)}\u2026", _work, _done)

    # ---------------------------------------------------------- Table

    def _populate_table(self):
        if not self._result:
            return
        data = []
        for part in self._result.parts:
            data.append([
                str(part.part_id),
                part.name,
                str(len(part.element_ids)),
                str(len(part.node_ids)),
                ', '.join(str(p) for p in sorted(part.property_ids)[:8])
                + ('...' if len(part.property_ids) > 8 else ''),
            ])
        self._sheet.set_sheet_data(data)
        self._sheet.readonly_columns(columns=[0, 2, 3, 4])

    def _update_joints_label(self):
        if not self._result:
            self._joints_label.configure(text="")
            return
        n_joints = len(self._result.joints)
        n_cbush = sum(len(j.chains) for j in self._result.joints)
        n_contact = sum(len(j.contact_pairs) for j in self._result.joints)
        parts = []
        if n_cbush:
            parts.append(f"{n_cbush} CBUSHes")
        if n_contact:
            parts.append(f"{n_contact} glue contacts")
        detail = f" ({', '.join(parts)})" if parts else ""
        self._joints_label.configure(
            text=f"Joints: {n_joints}{detail}")

    def _on_name_edited(self, event=None):
        """Sync edited Name column back to Part objects."""
        if not self._result:
            return
        for i, part in enumerate(self._result.parts):
            try:
                new_name = self._sheet.get_cell_data(i, 1)
                if new_name and new_name.strip():
                    part.name = new_name.strip()
            except (IndexError, TypeError):
                pass

    # ---------------------------------------------------------- Merge

    def _merge_selected(self):
        if not self._result:
            return

        selected = self._get_selected_rows()
        if len(selected) < 2:
            messagebox.showwarning(
                "Merge", "Select 2 or more rows to merge.")
            return

        part_ids = set()
        for row in selected:
            try:
                pid = int(self._sheet.get_cell_data(row, 0))
                part_ids.add(pid)
            except (ValueError, TypeError, IndexError):
                pass

        if len(part_ids) < 2:
            return

        names = [p.name for p in self._result.parts if p.part_id in part_ids]
        if not messagebox.askyesno(
                "Merge Parts",
                f"Merge {len(part_ids)} parts?\n\n" +
                "\n".join(f"  - {n}" for n in names)):
            return

        self._result = merge_parts(self._result, part_ids)
        self._mesh = None  # invalidate
        self._populate_table()
        self._update_joints_label()
        self._status_label.configure(
            text=f"Merged → {len(self._result.parts)} parts, "
                 f"{len(self._result.joints)} joints",
            text_color=("gray10", "gray90"))

    def _get_selected_rows(self):
        """Get list of selected row indices from tksheet."""
        rows = set()
        try:
            currently = self._sheet.get_currently_selected()
            if currently is not None:
                if hasattr(currently, 'row') and currently.row is not None:
                    rows.add(currently.row)
        except Exception:
            pass

        try:
            for item in self._sheet.get_selected_rows():
                if isinstance(item, int):
                    rows.add(item)
                elif hasattr(item, 'row'):
                    rows.add(item.row)
        except Exception:
            pass

        return sorted(rows)

    # ---------------------------------------------------------- 3D Preview

    def _show_preview(self):
        if not self._result or not self._model:
            return

        if self._mesh is None:
            self._status_label.configure(
                text="Building 3D mesh\u2026", text_color="gray")
            self.update_idletasks()
            mesh, available = build_pyvista_mesh(self._model, self._result.parts)
            if not available:
                messagebox.showwarning(
                    "pyvista not available",
                    "Install pyvista for 3D preview:\n  pip install pyvista")
                return
            if mesh is None:
                messagebox.showwarning("Preview", "No displayable elements found.")
                return
            self._mesh = mesh

        self._status_label.configure(
            text="Showing 3D preview\u2026", text_color="gray")
        self.update_idletasks()

        try:
            show_partition_preview(self._mesh, self._result.parts)
        except Exception as exc:
            messagebox.showerror("Preview Error", str(exc))

        self._status_label.configure(
            text=f"{len(self._result.parts)} parts, "
                 f"{len(self._result.joints)} joints",
            text_color=("gray10", "gray90"))

    # ---------------------------------------------------------- Write

    def _write_output(self):
        if not self._result or not self._model:
            return

        outdir = self._outdir_var.get().strip()
        if not outdir:
            messagebox.showwarning("No output directory",
                                   "Please set an output directory.")
            return

        if os.path.exists(outdir) and os.listdir(outdir):
            if not messagebox.askyesno(
                    "Directory not empty",
                    f"Output directory is not empty:\n{outdir}\n\n"
                    "Files may be overwritten. Continue?"):
                return

        def _work():
            return write_partition(
                self._model, self._result, outdir, self._bdf_path,
                log_fn=lambda msg: None,  # background thread, no UI updates
            )

        def _done(stats, error):
            if error:
                import traceback
                messagebox.showerror(
                    "Write Error",
                    f"Could not write files:\n{traceback.format_exc()}")
                self._status_label.configure(
                    text="Write failed", text_color="red")
                self._restore_buttons()
                return

            self._restore_buttons()

            # Validation summary
            msg_parts = [f"Written to: {outdir}"]
            if stats:
                msg_parts.append(
                    f"Elements: {stats['written_elems']}/{stats['total_elems']}")
                msg_parts.append(
                    f"Nodes: {stats['written_nodes']}/{stats['total_nodes']}")
                missing_e = stats['total_elems'] - stats['written_elems']
                missing_n = stats['total_nodes'] - stats['written_nodes']
                if missing_e > 0 or missing_n > 0:
                    msg_parts.append(
                        f"\nNote: {missing_e} elements and {missing_n} nodes "
                        "in shared/joint files (materials, rigid elements, etc.)")
            messagebox.showinfo("Success", "\n".join(msg_parts))
            self._status_label.configure(
                text=f"Written {len(self._result.parts)} part files + "
                     f"{len(self._result.joints)} joint files",
                text_color=("gray10", "gray90"))

        self._run_in_background(f"Writing to {outdir}\u2026", _work, _done)

    # ---------------------------------------------------------- Browse outdir

    def _browse_outdir(self):
        d = filedialog.askdirectory(title="Select output directory")
        if d:
            self._outdir_var.set(d)


def main():
    import logging
    logging.getLogger("customtkinter").setLevel(logging.ERROR)
    root = ctk.CTk()
    root.title("BDF Partitioner")
    root.geometry("900x550")
    app = PartitionTool(root)
    app.pack(fill=tk.BOTH, expand=True)
    root.mainloop()


if __name__ == '__main__':
    main()
