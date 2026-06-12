"""PyInstaller runtime hook: ensure sys.stdout/sys.stderr are never None.

In a windowed (console=False) PyInstaller build there is no console, so
PyInstaller sets sys.stdout and sys.stderr to None. Libraries that write a
warning to stderr at import time (e.g. customtkinter's font module) then crash
with "'NoneType' object has no attribute 'write'". Giving the streams a dummy
sink before any such import keeps the app alive. Runs before all other imports.
"""
import sys
import io

if sys.stderr is None:
    sys.stderr = io.StringIO()
if sys.stdout is None:
    sys.stdout = io.StringIO()
