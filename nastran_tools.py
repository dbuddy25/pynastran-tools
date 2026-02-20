#!/usr/bin/env python3
"""Nastran Tools — unified GUI application.

Sidebar-navigated host for Pre-Processing and Post-Processing tools:
  - Mass Scale (BDF mass scaling by include file)
  - Renumber (include file ID renumbering)
  - MEFFMASS (modal effective mass fractions from OP2)

Usage:
    python nastran_tools.py
"""
import os
import sys
import tkinter as tk

import customtkinter as ctk
import numpy as np
if not hasattr(np, 'in1d'):
    np.in1d = np.isin

# Ensure sub-packages are importable
_root = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(_root, 'preprocessing'))
sys.path.insert(0, os.path.join(_root, 'postprocessing'))

from mass_scale import MassScaleTool
from renumber_includes import RenumberIncludesTool

# MeffModule depends on numpy/scipy — lazy import with fallback
_meff_available = True
try:
    from modules.meff import MeffModule
except Exception:
    _meff_available = False


__version__ = "0.1.0"


def show_guide(parent, title, text):
    """Open a non-modal guide dialog with read-only text."""
    win = ctk.CTkToplevel(parent)
    win.title(title)
    win.geometry("550x420")
    win.resizable(True, True)
    win.transient(parent)

    tb = ctk.CTkTextbox(win, wrap="word")
    tb.pack(fill="both", expand=True, padx=10, pady=(10, 5))
    tb.insert("1.0", text)
    tb.configure(state="disabled")

    ctk.CTkButton(win, text="Close", width=80,
                  command=win.destroy).pack(pady=(0, 10))


class Sidebar(ctk.CTkFrame):
    """Fixed-width sidebar with section headers and tool buttons."""

    def __init__(self, parent, on_select):
        super().__init__(parent, width=200)
        self.pack_propagate(False)

        self._on_select = on_select
        self._buttons = {}
        self._active_key = None

        # App title
        ctk.CTkLabel(
            self, text="Nastran Tools",
            font=ctk.CTkFont(size=16, weight="bold"),
        ).pack(padx=12, pady=(16, 12))

        # Pre-Processing section
        self._add_section("Pre-Processing")
        self._add_tool("mass_scale", "Mass Scale")
        self._add_tool("renumber", "Renumber")

        # Post-Processing section
        self._add_section("Post-Processing")
        self._add_tool("meff", "MEFFMASS")

    def _add_section(self, label):
        ctk.CTkLabel(
            self, text=label,
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="gray",
        ).pack(anchor=tk.W, padx=16, pady=(14, 4))

    def _add_tool(self, key, label):
        btn = ctk.CTkButton(
            self, text=f"  {label}", anchor=tk.W,
            fg_color="transparent", text_color=("gray10", "gray90"),
            hover_color=("gray75", "gray30"),
            command=lambda k=key: self._select(k),
        )
        btn.pack(fill=tk.X, padx=8, pady=1)
        self._buttons[key] = btn

    def _select(self, key):
        # Deselect previous
        if self._active_key and self._active_key in self._buttons:
            self._buttons[self._active_key].configure(
                fg_color="transparent")

        # Highlight new
        self._active_key = key
        self._buttons[key].configure(
            fg_color=("gray75", "gray30"))

        self._on_select(key)

    def set_active(self, key):
        """Programmatically set the active button."""
        self._select(key)

    def disable_tool(self, key):
        """Disable a tool button (e.g. when dependencies are missing)."""
        if key in self._buttons:
            self._buttons[key].configure(state=tk.DISABLED)


class NastranToolsApp(ctk.CTk):
    """Unified Nastran tools application with sidebar navigation."""

    def __init__(self):
        super().__init__()
        self.title(f"Nastran Tools v{__version__}")
        self.geometry("1400x800")
        self.minsize(1000, 600)

        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        # Layout: sidebar + content
        self._sidebar = Sidebar(self, on_select=self._switch_tool)
        self._sidebar.pack(side=tk.LEFT, fill=tk.Y)

        self._content = ctk.CTkFrame(self)
        self._content.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Create tool modules
        self._tools = {}

        self._tools['mass_scale'] = MassScaleTool(self._content)
        self._tools['renumber'] = RenumberIncludesTool(self._content)

        if _meff_available:
            meff = MeffModule(self._content)
            self._tools['meff'] = meff.frame
        else:
            self._sidebar.disable_tool("meff")

        self._active_tool = None

        # Show first tool by default
        self._sidebar.set_active('mass_scale')

    def _switch_tool(self, key):
        """Hide the current tool frame and show the selected one."""
        if self._active_tool is not None:
            self._active_tool.pack_forget()

        tool_widget = self._tools.get(key)
        if tool_widget is not None:
            tool_widget.pack(fill=tk.BOTH, expand=True)
            self._active_tool = tool_widget


def main():
    import logging
    logging.getLogger("customtkinter").setLevel(logging.ERROR)
    app = NastranToolsApp()
    app.mainloop()


if __name__ == '__main__':
    main()
