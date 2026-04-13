#!/usr/bin/env python3
"""Structures Tools — unified GUI application.

Sidebar-navigated host for Pre-Processing and Post-Processing tools:
  - Mass Scale (BDF mass scaling by include file)
  - Renumber (include file ID renumbering)
  - MEFFMASS (modal effective mass fractions from OP2)
  - ESE Breakdown (element strain energy % by group from OP2)
  - CBUSH Forces (CBUSH element forces from OP2)
  - Mass Breakdown (BDF mass breakdown by group with GPWG validation)

Usage:
    python structures_tools.py
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
sys.path.insert(0, os.path.join(_root, 'calculators'))

from mass_scale import MassScaleTool
from renumber_includes import RenumberIncludesTool
from thermal_cte import ThermalCteTool

# MeffModule depends on numpy/scipy — lazy import with fallback
_meff_available = True
try:
    from modules.meff import MeffModule
except Exception:
    _meff_available = False

# EnergyBreakdownModule — lazy import with fallback
_energy_available = True
try:
    from modules.energy_breakdown import EnergyBreakdownModule
except Exception:
    _energy_available = False

# CbushForcesModule — lazy import with fallback
_cbush_available = True
try:
    from modules.cbush_forces import CbushForcesModule
except Exception:
    _cbush_available = False

# MassBreakdownModule — lazy import with fallback
_mass_breakdown_available = True
try:
    from modules.mass_breakdown import MassBreakdownModule
except Exception:
    _mass_breakdown_available = False

from miles_equation import MilesEquationTool


__version__ = "0.2.0"


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
            self, text="Structures Tools",
            font=ctk.CTkFont(size=16, weight="bold"),
        ).pack(padx=12, pady=(16, 12))

        # Pre-Processing section
        self._add_section("Pre-Processing")
        self._add_tool("mass_scale", "Mass Scale")
        self._add_tool("renumber", "Renumber")
        self._add_tool("thermal_cte", "Thermal CTE")

        # Post-Processing section
        self._add_section("Post-Processing")
        self._add_tool("meff", "MEFFMASS")
        self._add_tool("energy", "ESE Breakdown")
        self._add_tool("cbush", "CBUSH Forces")
        self._add_tool("mass_breakdown", "Mass Breakdown")

        # Hand Calcs section
        self._add_section("Hand Calcs")
        self._add_tool("miles", "Miles Eq — RMS Disp")

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


class StructuresToolsApp(ctk.CTk):
    """Unified structural analysis tools application with sidebar navigation."""

    def __init__(self):
        super().__init__()
        self.title(f"Structures Tools v{__version__}")
        self.geometry("1400x800")
        self.minsize(1000, 600)

        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        # Layout: sidebar + content
        self._sidebar = Sidebar(self, on_select=self._switch_tool)
        self._sidebar.pack(side=tk.LEFT, fill=tk.Y)

        self._content = ctk.CTkFrame(self)
        self._content.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Intro page
        self._intro = self._build_intro()
        self._intro.pack(fill=tk.BOTH, expand=True)

        # Create tool modules
        self._tools = {}

        self._tools['mass_scale'] = MassScaleTool(self._content)
        self._tools['renumber'] = RenumberIncludesTool(self._content)
        self._tools['thermal_cte'] = ThermalCteTool(self._content)

        if _meff_available:
            meff = MeffModule(self._content)
            self._tools['meff'] = meff.frame
        else:
            self._sidebar.disable_tool("meff")

        if _energy_available:
            energy = EnergyBreakdownModule(self._content)
            self._tools['energy'] = energy.frame
        else:
            self._sidebar.disable_tool("energy")

        if _cbush_available:
            cbush = CbushForcesModule(self._content)
            self._tools['cbush'] = cbush.frame
        else:
            self._sidebar.disable_tool("cbush")

        if _mass_breakdown_available:
            mass_bd = MassBreakdownModule(self._content)
            self._tools['mass_breakdown'] = mass_bd.frame
        else:
            self._sidebar.disable_tool("mass_breakdown")

        self._tools['miles'] = MilesEquationTool(self._content)

        self._active_tool = None

    def _build_intro(self):
        """Create the landing page shown on startup."""
        frame = ctk.CTkFrame(self._content)

        # Vertical centering spacer
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_rowconfigure(6, weight=1)
        frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            frame, text="Structures Tools",
            font=ctk.CTkFont(size=28, weight="bold"),
        ).grid(row=1, column=0, pady=(0, 4))

        ctk.CTkLabel(
            frame, text=f"v{__version__}",
            font=ctk.CTkFont(size=14),
            text_color="gray",
        ).grid(row=2, column=0, pady=(0, 30))

        categories = (
            "Pre-Processing — Mass Scale, Renumber, Thermal CTE",
            "Post-Processing — MEFFMASS, ESE Breakdown, CBUSH Forces, Mass Breakdown",
            "Hand Calcs — Miles Equation",
        )
        for i, cat in enumerate(categories):
            ctk.CTkLabel(
                frame, text=cat,
                font=ctk.CTkFont(size=13),
            ).grid(row=3 + i, column=0, pady=2)

        ctk.CTkLabel(
            frame, text="Select a tool from the sidebar to get started.",
            font=ctk.CTkFont(size=12),
            text_color="gray",
        ).grid(row=3 + len(categories), column=0, pady=(20, 0))

        return frame

    def _switch_tool(self, key):
        """Hide the current tool frame and show the selected one."""
        # Hide intro page on first tool selection
        if self._intro is not None:
            self._intro.pack_forget()
            self._intro.destroy()
            self._intro = None

        if self._active_tool is not None:
            self._active_tool.pack_forget()

        tool_widget = self._tools.get(key)
        if tool_widget is not None:
            tool_widget.pack(fill=tk.BOTH, expand=True)
            self._active_tool = tool_widget


def main():
    import logging
    logging.getLogger("customtkinter").setLevel(logging.ERROR)
    app = StructuresToolsApp()
    app.mainloop()


if __name__ == '__main__':
    main()
