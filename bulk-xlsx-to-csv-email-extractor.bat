@echo off
setlocal enabledelayedexpansion

set "DIR=%~dp0"

for %%F in ("%DIR%*.xlsx") do (
  echo Conversion: %%~nxF
  powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$xlsx='%%~fF'; $csv='%%~dpnF.csv';" ^
    "$excel = New-Object -ComObject Excel.Application;" ^
    "$excel.DisplayAlerts = $false; $excel.Visible = $false;" ^
    "$wb = $excel.Workbooks.Open($xlsx);" ^
    "$xlCSV = 6; $wb.SaveAs($csv, $xlCSV);" ^
    "$wb.Close($false); $excel.Quit();" ^
    "[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null;" ^
    "[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null;" ^
    "[GC]::Collect(); [GC]::WaitForPendingFinalizers();"
)

REM Email extraction from all CSV files in the .bat folder (except emails.csv)
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$dir='%DIR%'; $out=Join-Path $dir 'emails.csv';" ^
  "'Email' | Set-Content -Encoding UTF8 $out;" ^
  "Get-ChildItem -Path $dir -Filter '*.csv' | Where-Object { $_.Name -ne 'emails.csv' } |" ^
  "Select-String -Pattern '[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}' -AllMatches |" ^
  "ForEach-Object { $_.Matches.Value } | Sort-Object -Unique | Add-Content -Encoding UTF8 $out"

echo Extraction completed to emails.csv file created in %DIR%
pause