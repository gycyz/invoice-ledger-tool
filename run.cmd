@echo off
rem Launcher for invoice_ledger.py (ASCII only on purpose: cmd.exe parses batch files as OEM/GBK,
rem so non-ASCII literals here would be garbled and break the command line).
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
setlocal
cd /d "%~dp0"
set "PYFILE="
for /f "delims=" %%f in ('dir /b /s "%~dp0*invoice_ledger.py" 2^>nul') do set "PYFILE=%%f"
if not defined PYFILE (
  echo [ERROR] invoice_ledger.py not found under this folder.
  pause
  exit /b 1
)
python "%PYFILE%"
if errorlevel 1 (
  echo.
  echo [ERROR] Run failed. Make sure Python is on PATH and pandas + openpyxl are installed.
  pause
  exit /b 1
)
echo.
echo Done. Press any key to close.
pause >nul
exit /b 0
