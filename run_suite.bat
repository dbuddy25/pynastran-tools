@echo off
REM One-click launcher for the full Structures Tools suite.
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
echo Checking/installing suite dependencies (first run may take a few minutes)...
%PYRUN% -m pip install --quiet --disable-pip-version-check -r requirements.txt
if errorlevel 1 (
  echo.
  echo Dependency install failed. Check your internet/proxy, or install
  echo manually:  %PYRUN% -m pip install -r requirements.txt
  pause
  exit /b 1
)

echo Launching Structures Tools...
%PYRUN% structures_tools.py
