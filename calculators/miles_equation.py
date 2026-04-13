#!/usr/bin/env python3
"""Miles Equation — RMS Displacement Calculator.

Computes SDOF random vibration displacement response using Miles' Equation:

    Y_rms = sqrt( Q * ASD * g^2 / (32 * pi^3 * fn^3) )

where ASD is input in G²/Hz and g converts to length units.

Reference: John W. Miles, "On Structural Fatigue Under Random Loading",
Journal of the Aeronautical Sciences, pg. 753, November, 1954.
"""
import math
import tkinter as tk

import customtkinter as ctk


# Gravity constants for unit conversion (ASD in G²/Hz → length/s²)
_G_ACCEL = {
    'in': 386.1,       # in/s²
    'm': 9.80665,      # m/s²
}

_UNIT_LABELS = {
    'in': 'inches',
    'm': 'meters',
}

_GUIDE_TEXT = """\
MILES' EQUATION — RMS DISPLACEMENT

Computes the RMS displacement response of a single degree of freedom
(SDOF) system subjected to random vibration (white noise input).

EQUATION
  Y_rms = sqrt( Q * ASD * g² / (32 * π³ * fn³) )

  where ASD is input in G²/Hz and g converts to length units:
    g = 386.1 in/s²  (for inches)
    g = 9.807 m/s²   (for meters)

INPUTS
  fn          Natural frequency of the SDOF system (Hz)
  Q           Transmissibility (amplification factor) at fn
              Q = 1 / (2 * zeta), where zeta is the critical
              damping ratio.  Typical values: 10–50.
  ASD Input   Input acceleration spectral density at fn
              in units of G²/Hz.
  Units       Length unit for displacement output (in or m)

OUTPUTS
  Y_rms       RMS displacement (in selected units)
  3σ Disp     Peak displacement = 3 × Y_rms

DAMPING INPUT
  Enter either Q (transmissibility) or ζ (critical damping ratio).
  The tool converts between them:
      Q = 1 / (2ζ)        ζ = 1 / (2Q)

NOTES
  • Miles' Equation assumes a flat (white noise) input spectrum.
    For shaped spectra it may underpredict the response.
  • The 3σ displacement is conservative for design purposes.
  • This form of the equation converts ASD from G²/Hz to
    (length/s²)²/Hz internally using the gravity constant.
"""


class MilesEquationTool(ctk.CTkFrame):
    """Miles Equation — RMS Displacement calculator."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._build_ui()

    def _build_ui(self):
        # Toolbar
        toolbar = ctk.CTkFrame(self, fg_color="transparent")
        toolbar.pack(fill=tk.X, padx=5, pady=(5, 0))

        ctk.CTkLabel(
            toolbar, text="Miles' Equation — RMS Displacement",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(side=tk.LEFT, padx=(5, 0))

        ctk.CTkButton(
            toolbar, text="?", width=30, font=ctk.CTkFont(weight="bold"),
            command=self._show_guide,
        ).pack(side=tk.RIGHT, padx=(5, 0))

        ctk.CTkButton(
            toolbar, text="Clear", width=70,
            command=self._clear,
        ).pack(side=tk.RIGHT)

        ctk.CTkButton(
            toolbar, text="Calculate", width=100,
            command=self._calculate,
        ).pack(side=tk.RIGHT, padx=(0, 5))

        # Main content
        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # --- Input section ---
        input_frame = ctk.CTkFrame(body)
        input_frame.pack(fill=tk.X, pady=(0, 10))

        ctk.CTkLabel(
            input_frame, text="Inputs",
            font=ctk.CTkFont(size=12, weight="bold"),
        ).grid(row=0, column=0, columnspan=4, sticky=tk.W,
               padx=10, pady=(10, 5))

        # Row 1: fn and ASD
        self._fn_var = tk.StringVar()
        ctk.CTkLabel(input_frame, text="fn (Hz):").grid(
            row=1, column=0, sticky=tk.E, padx=(10, 5), pady=4)
        fn_entry = ctk.CTkEntry(input_frame, textvariable=self._fn_var,
                                width=120)
        fn_entry.grid(row=1, column=1, padx=(0, 20), pady=4)

        self._asd_var = tk.StringVar()
        ctk.CTkLabel(input_frame, text="ASD (G\u00b2/Hz):").grid(
            row=1, column=2, sticky=tk.E, padx=(10, 5), pady=4)
        ctk.CTkEntry(input_frame, textvariable=self._asd_var,
                      width=120).grid(row=1, column=3, padx=(0, 10), pady=4)

        # Row 2: Damping
        self._damping_mode = tk.StringVar(value="Q")
        ctk.CTkLabel(input_frame, text="Damping:").grid(
            row=2, column=0, sticky=tk.E, padx=(10, 5), pady=4)

        damping_row = ctk.CTkFrame(input_frame, fg_color="transparent")
        damping_row.grid(row=2, column=1, columnspan=3, sticky=tk.W, pady=4)

        ctk.CTkRadioButton(
            damping_row, text="Q", variable=self._damping_mode,
            value="Q", command=self._on_damping_toggle,
        ).pack(side=tk.LEFT, padx=(0, 5))

        self._q_var = tk.StringVar()
        self._q_entry = ctk.CTkEntry(damping_row, textvariable=self._q_var,
                                      width=80, placeholder_text="e.g. 10")
        self._q_entry.pack(side=tk.LEFT, padx=(0, 20))

        ctk.CTkRadioButton(
            damping_row, text="\u03b6 (zeta)", variable=self._damping_mode,
            value="zeta", command=self._on_damping_toggle,
        ).pack(side=tk.LEFT, padx=(0, 5))

        self._zeta_var = tk.StringVar()
        self._zeta_entry = ctk.CTkEntry(
            damping_row, textvariable=self._zeta_var,
            width=80, placeholder_text="e.g. 0.05")
        self._zeta_entry.pack(side=tk.LEFT)

        self._zeta_entry.configure(state=tk.DISABLED)

        # Row 3: Units
        self._unit_var = tk.StringVar(value="in")
        ctk.CTkLabel(input_frame, text="Units:").grid(
            row=3, column=0, sticky=tk.E, padx=(10, 5), pady=(4, 10))

        unit_row = ctk.CTkFrame(input_frame, fg_color="transparent")
        unit_row.grid(row=3, column=1, columnspan=3, sticky=tk.W,
                      pady=(4, 10))

        ctk.CTkRadioButton(
            unit_row, text="inches", variable=self._unit_var, value="in",
        ).pack(side=tk.LEFT, padx=(0, 15))
        ctk.CTkRadioButton(
            unit_row, text="meters", variable=self._unit_var, value="m",
        ).pack(side=tk.LEFT)

        # --- Results section ---
        results_frame = ctk.CTkFrame(body)
        results_frame.pack(fill=tk.X, pady=(0, 10))

        ctk.CTkLabel(
            results_frame, text="Results",
            font=ctk.CTkFont(size=12, weight="bold"),
        ).grid(row=0, column=0, columnspan=4, sticky=tk.W,
               padx=10, pady=(10, 5))

        self._result_labels = {}
        result_defs = [
            ("Y_rms",    "Y_rms:"),
            ("3sigma_Y", "3\u03c3 Disp:"),
            ("Q_out",    "Q:"),
            ("zeta_out", "\u03b6:"),
        ]
        for i, (key, label) in enumerate(result_defs):
            r = i // 2 + 1
            c = (i % 2) * 2
            ctk.CTkLabel(results_frame, text=label).grid(
                row=r, column=c, sticky=tk.E, padx=(10, 5), pady=4)
            val_label = ctk.CTkLabel(
                results_frame, text="\u2014",
                font=ctk.CTkFont(size=13, weight="bold"),
            )
            val_label.grid(row=r, column=c + 1, sticky=tk.W,
                           padx=(0, 20), pady=4)
            self._result_labels[key] = val_label

        # Bottom padding
        ctk.CTkLabel(results_frame, text="").grid(
            row=3, column=0, pady=(0, 6))

        # Equation display
        eq_frame = ctk.CTkFrame(body)
        eq_frame.pack(fill=tk.X)
        ctk.CTkLabel(
            eq_frame,
            text="Y_rms = \u221a( Q \u00b7 ASD \u00b7 g\u00b2 / (32\u03c0\u00b3 \u00b7 fn\u00b3) )",
            font=ctk.CTkFont(family="Courier", size=14),
            text_color="gray",
        ).pack(padx=10, pady=10)

        # Focus first field
        fn_entry.focus_set()

    def _on_damping_toggle(self):
        mode = self._damping_mode.get()
        if mode == "Q":
            self._q_entry.configure(state=tk.NORMAL)
            self._zeta_entry.configure(state=tk.DISABLED)
        else:
            self._q_entry.configure(state=tk.DISABLED)
            self._zeta_entry.configure(state=tk.NORMAL)

    def _calculate(self):
        # Parse fn
        try:
            fn = float(self._fn_var.get())
            if fn <= 0:
                raise ValueError
        except (ValueError, TypeError):
            self._show_error("Enter a positive number for fn.")
            return

        # Parse ASD
        try:
            asd = float(self._asd_var.get())
            if asd <= 0:
                raise ValueError
        except (ValueError, TypeError):
            self._show_error("Enter a positive number for ASD.")
            return

        # Parse damping
        mode = self._damping_mode.get()
        if mode == "Q":
            try:
                Q = float(self._q_var.get())
                if Q <= 0:
                    raise ValueError
            except (ValueError, TypeError):
                self._show_error("Enter a positive number for Q.")
                return
            zeta = 1.0 / (2.0 * Q)
        else:
            try:
                zeta = float(self._zeta_var.get())
                if zeta <= 0 or zeta >= 1:
                    raise ValueError
            except (ValueError, TypeError):
                self._show_error("Enter \u03b6 between 0 and 1.")
                return
            Q = 1.0 / (2.0 * zeta)

        # Unit conversion
        unit = self._unit_var.get()
        g = _G_ACCEL[unit]
        unit_label = _UNIT_LABELS[unit]

        # Miles displacement equation:
        # Y_rms = sqrt( Q * ASD_g * g^2 / (32 * pi^3 * fn^3) )
        numerator = Q * asd * g ** 2
        denominator = 32.0 * math.pi ** 3 * fn ** 3
        y_rms = math.sqrt(numerator / denominator)
        three_sigma_y = 3.0 * y_rms

        # Update results
        self._result_labels["Y_rms"].configure(
            text=f"{y_rms:.6g} {unit_label}")
        self._result_labels["3sigma_Y"].configure(
            text=f"{three_sigma_y:.6g} {unit_label}")
        self._result_labels["Q_out"].configure(text=f"{Q:.2f}")
        self._result_labels["zeta_out"].configure(text=f"{zeta:.4f}")

    def _clear(self):
        for var in (self._fn_var, self._asd_var, self._q_var,
                    self._zeta_var):
            var.set("")
        for lbl in self._result_labels.values():
            lbl.configure(text="\u2014")

    def _show_error(self, msg):
        from tkinter import messagebox
        messagebox.showerror("Input Error", msg,
                             parent=self.winfo_toplevel())

    def _show_guide(self):
        try:
            from nastran_tools import show_guide
        except ImportError:
            from tkinter import messagebox
            messagebox.showinfo("Guide", _GUIDE_TEXT,
                                parent=self.winfo_toplevel())
            return
        show_guide(self.winfo_toplevel(), "Miles' Equation — RMS Disp",
                   _GUIDE_TEXT)


def main():
    import logging
    logging.getLogger("customtkinter").setLevel(logging.ERROR)
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")

    root = ctk.CTk()
    root.title("Miles' Equation — RMS Displacement")
    root.geometry("700x450")
    root.minsize(600, 400)

    tool = MilesEquationTool(root)
    tool.pack(fill=tk.BOTH, expand=True)

    root.mainloop()


if __name__ == '__main__':
    main()
