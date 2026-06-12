@echo off
REM One-click launcher for the ESE (strain energy) tool.
REM Double-click this file. First run installs the Python packages it needs;
REM after that it just launches. Requires Python 3 (python.org or Anaconda).
cd /d "%~dp0"

REM Find a Python: prefer the 'py' launcher (most reliable on Windows), then
REM fall back to 'python' on PATH.
set "PYRUN="
where py >nul 2>nul && set "PYRUN=py"
if not defined PYRUN (
  where python >nul 2>nul && set "PYRUN=python"
)
if not defined PYRUN (
  echo.
  echo Python 3 was not found on this machine.
  echo Install it from https://python.org  ^(tick "Add Python to PATH"^),
  echo then double-click this file again.
  echo.
  pause
  exit /b 1
)

echo Using %PYRUN%
echo Checking/installing ESE dependencies (first run may take a minute)...
%PYRUN% -m pip install --quiet --disable-pip-version-check customtkinter tksheet numpy pyNastran openpyxl
if errorlevel 1 (
  echo.
  echo Dependency install failed. Check your internet/proxy, or install the
  echo packages manually:  %PYRUN% -m pip install customtkinter tksheet numpy pyNastran openpyxl
  pause
  exit /b 1
)

echo Launching ESE tool...
%PYRUN% run_ese.py
