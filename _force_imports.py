"""Build-time helper for PyInstaller — NOT used at runtime.

PyInstaller reliably bundles a submodule when it follows a real `import` edge
to it (this is why `modules.meff` survived: postprocessing/modal_effective_mass.py
imports it directly). Hand-listed `hiddenimports` names for the generic `modules`
package did NOT resolve. So this module imports every suite submodule via real
edges; adding `_force_imports` to the spec's hiddenimports makes PyInstaller walk
these edges and bundle them all. The try/except keeps dev imports from failing.
"""
try:  # noqa: SIM105
    import modules.meff
    import modules.energy_breakdown
    import modules.cbush_forces
    import modules.mass_breakdown
    import modules.asd_overlay
    import modules.asd_common
    import modules.response_limiting
    import modules.random_vibe_env
except Exception:
    pass
