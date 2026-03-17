# Math & Sciences Hub: Export tutors.json from Excel (PowerShell)
# Run: Right-click -> Run with PowerShell (or open PowerShell here and run .\RUN_EXPORT_TUTORS.ps1)

$XLSX = "Math_and_Sciences_Hub_Full_Institution_System_ENHANCED_UNPROTECTED.xlsx"
$OUT  = "tutors.json"
$YEAR = 2026

# Prefer python, fallback to py
$pythonCmd = $null
if (Get-Command python -ErrorAction SilentlyContinue) { $pythonCmd = "python" }
elseif (Get-Command py -ErrorAction SilentlyContinue) { $pythonCmd = "py" }
elseif (Get-Command python3 -ErrorAction SilentlyContinue) { $pythonCmd = "python3" }

if (-not $pythonCmd) {
  Write-Host "ERROR: Python not found. Install Python 3 from https://www.python.org/downloads/ (tick Add to PATH) then retry." -ForegroundColor Red
  pause
  exit 1
}

Write-Host "Using: $pythonCmd"
& $pythonCmd -m pip install openpyxl
& $pythonCmd .\export_tutors_json.py --xlsx $XLSX --out $OUT --year $YEAR --writeback

Write-Host ""
Write-Host "Done. Upload/commit tutors.json to GitHub (same folder as index.html)." -ForegroundColor Green
pause
