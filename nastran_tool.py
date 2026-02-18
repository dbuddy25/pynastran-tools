#!/usr/bin/env python3
"""Nastran post-processing GUI tool.

Modular tab-based interface. Each module lives in modules/ and is
registered in _register_modules().

Usage:
    python nastran_tool.py
"""
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from pyNastran.op2.op2 import OP2

from modules.meff import MeffModule

_MACOS = sys.platform == 'darwin'
_MOD_KEY = 'Command' if _MACOS else 'Control'
_ACCEL = 'Cmd' if _MACOS else 'Ctrl'


class NastranTool(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Nastran Post-Processing Tool")
        self.geometry("1400x600")

        self.op2 = None
        self.op2_b = None
        self._op2_path = None
        self._op2_b_path = None
        self.modules = []

        self._build_menu()
        self._build_ui()
        self._register_modules()

    def _build_menu(self):
        menubar = tk.Menu(self)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Open OP2\u2026",
                              command=self._open_op2,
                              accelerator=f"{_ACCEL}+O")
        file_menu.add_command(label="Open Comparison OP2\u2026",
                              command=self._open_comparison_op2,
                              accelerator=f"{_ACCEL}+Shift+O",
                              state=tk.DISABLED)
        file_menu.add_command(label="Clear Comparison",
                              command=self._clear_comparison,
                              state=tk.DISABLED)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.quit)
        menubar.add_cascade(label="File", menu=file_menu)

        self._file_menu = file_menu
        self.config(menu=menubar)
        self.bind_all(f"<{_MOD_KEY}-o>", lambda e: self._open_op2())
        self.bind_all(f"<{_MOD_KEY}-Shift-o>", lambda e: self._open_comparison_op2())

    def _build_ui(self):
        self.status_var = tk.StringVar(value="No file loaded")
        status = ttk.Label(self, textvariable=self.status_var,
                           relief=tk.SUNKEN, anchor=tk.W, padding=(4, 2))
        status.pack(side=tk.BOTTOM, fill=tk.X)

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    def _register_modules(self):
        self._add_module(MeffModule)

    def _add_module(self, module_class):
        module = module_class(self.notebook)
        self.notebook.add(module.frame, text=module.name)
        self.modules.append(module)

    def _open_op2(self):
        path = filedialog.askopenfilename(
            title="Open OP2 File",
            filetypes=[("OP2 files", "*.op2"), ("All files", "*.*")])
        if not path:
            return

        self.status_var.set(f"Loading {path}\u2026")
        self.update_idletasks()

        try:
            op2 = OP2()
            op2.read_op2(path)
        except Exception as exc:
            messagebox.showerror("Error", f"Could not read OP2:\n{exc}")
            self.status_var.set("Load failed")
            return

        # Clear any existing comparison before loading new primary
        if self.op2_b is not None:
            self._clear_comparison()

        self.op2 = op2
        self._op2_path = path
        self.status_var.set(path)

        # Enable comparison menu item
        self._file_menu.entryconfigure(1, state=tk.NORMAL)

        for module in self.modules:
            module.load(op2)

    def _open_comparison_op2(self):
        if self.op2 is None:
            return

        path = filedialog.askopenfilename(
            title="Open Comparison OP2 File",
            filetypes=[("OP2 files", "*.op2"), ("All files", "*.*")])
        if not path:
            return

        self.status_var.set(f"Loading comparison {path}\u2026")
        self.update_idletasks()

        try:
            op2_b = OP2()
            op2_b.read_op2(path)
        except Exception as exc:
            messagebox.showerror("Error", f"Could not read OP2:\n{exc}")
            self.status_var.set(self._op2_path)
            return

        self.op2_b = op2_b
        self._op2_b_path = path
        self.status_var.set(f"A: {self._op2_path}  |  B: {self._op2_b_path}")

        # Enable clear comparison
        self._file_menu.entryconfigure(2, state=tk.NORMAL)

        for module in self.modules:
            module.load_comparison(op2_b)

    def _clear_comparison(self):
        self.op2_b = None
        self._op2_b_path = None
        self.status_var.set(self._op2_path or "No file loaded")

        # Disable clear comparison
        self._file_menu.entryconfigure(2, state=tk.DISABLED)

        for module in self.modules:
            module.clear_comparison()


def main():
    app = NastranTool()
    app.mainloop()


if __name__ == '__main__':
    main()
