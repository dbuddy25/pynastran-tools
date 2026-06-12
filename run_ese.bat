@echo off
REM One-click launcher for the ESE (strain energy) tool.
REM Double-click this file. First run installs the Python packages it needs;
REM after that it just launches. Requires Python 3 to be installed.
cd /d "%~dp0"

echo Checking/installing ESE dependencies (first run may take a minute)...
python -m pip install --quiet --disable-pip-version-check customtkinter tksheet numpy pyNastran openpyxl
if errorlevel 1 (
  echo.
  echo Dependency install failed. Make sure Python and pip are on PATH:
  echo     python --version
  echo If that fails, install Python 3 from python.org ^(tick "Add to PATH"^).
  pause
  exit /b 1
)

echo Launching ESE tool...
python run_ese.py
