@echo off
REM === Math & Sciences Hub: Export tutors.json from Excel (Windows) ===
REM This script tries: py -> python -> python3
REM Place this file in the same folder as:
REM   1) export_tutors_json.py
REM   2) your Excel workbook

set XLSX=Math_and_Sciences_Hub_Full_Institution_System_ENHANCED_UNPROTECTED.xlsx
set OUT=tutors.json
set YEAR=2026

set PYTHON=

where py >nul 2>nul
if %errorlevel%==0 set PYTHON=py

if "%PYTHON%"=="" (
  where python >nul 2>nul
  if %errorlevel%==0 set PYTHON=python
)

if "%PYTHON%"=="" (
  where python3 >nul 2>nul
  if %errorlevel%==0 set PYTHON=python3
)

if "%PYTHON%"=="" (
  echo.
  echo ERROR: Python not found on this PC.
  echo.
  echo Fix options:
  echo 1^) Install Python 3 from https://www.python.org/downloads/  ^(tick "Add Python to PATH" during install^)
  echo 2^) OR install "Python" from Microsoft Store, then reopen this window
  echo.
  echo After installing, close and re-open this folder, then run this file again.
  echo.
  pause
  exit /b 1
)

echo Using: %PYTHON%
echo Installing openpyxl (required)...
%PYTHON% -m pip install --quiet openpyxl

echo Exporting tutors.json...
%PYTHON% export_tutors_json.py --xlsx "%XLSX%" --out "%OUT%" --year %YEAR% --writeback

echo.
echo Done. Upload/commit %OUT% to GitHub (same folder as index.html).
pause
