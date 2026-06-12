#!/usr/bin/env python3
"""Standalone launcher for ONLY the ESE (strain energy) breakdown tool.

For a machine that has Python installed but doesn't need the whole suite. Keep
this file in the project folder (it pulls in energy_breakdown.py and bdf_utils.py
from here) and run:

    python run_ese.py

Requires (pip install): customtkinter  tksheet  numpy  pyNastran  openpyxl
(NOT scipy/matplotlib — the ESE tool doesn't use them.)
"""
import os
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
# energy_breakdown.py lives in postprocessing/modules; bdf_utils.py is at the
# repo root (imported lazily when grouping by a BDF). Put both on sys.path.
for _sub in ('', os.path.join('postprocessing', 'modules'), 'preprocessing'):
    _p = os.path.join(_HERE, _sub) if _sub else _HERE
    if _p not in sys.path:
        sys.path.insert(0, _p)


def _fail(missing):
    msg = (f"Missing required package: {missing}\n\n"
           "Install the ESE tool's dependencies with:\n"
           "    pip install customtkinter tksheet numpy pyNastran openpyxl")
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Missing dependency", msg)
    except Exception:
        print(msg, file=sys.stderr)
    sys.exit(1)


if __name__ == '__main__':
    try:
        from energy_breakdown import main
    except ModuleNotFoundError as exc:
        _fail(exc.name)
    main()
