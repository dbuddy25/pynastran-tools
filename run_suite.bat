@echo off
REM One-click launcher for the full Structures Tools suite.
REM Double-click this file. First run installs the Python packages it needs;
REM after that it just launches. Requires Python 3 to be installed.
cd /d "%~dp0"

echo Checking/installing suite dependencies (first run may take a few minutes)...
python -m pip install --quiet --disable-pip-version-check -r requirements.txt
if errorlevel 1 (
  echo.
  echo Dependency install failed. Make sure Python and pip are on PATH:
  echo     python --version
  echo If that fails, install Python 3 from python.org ^(tick "Add to PATH"^).
  pause
  exit /b 1
)

echo Launching Structures Tools...
python structures_tools.py
