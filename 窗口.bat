@echo off
setlocal
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"

rem ============================================================
rem  Windowed UI entry.  Equivalent to:  python window.py
rem    --debug      open the WebView devtools
rem    --no-window  serve only, no window
rem    --selftest   open for 4s then close (sanity check)
rem  Keep this file ASCII-only: batch + UTF-8 is fragile.
rem ============================================================

rem --- find a python that actually has openpyxl (build.py needs it) ---
set "PY="
where python >nul 2>nul
if not errorlevel 1 (
  python -c "import openpyxl" >nul 2>nul
  if not errorlevel 1 set "PY=python"
)
if not defined PY (
  where py >nul 2>nul
  if not errorlevel 1 (
    py -c "import openpyxl" >nul 2>nul
    if not errorlevel 1 set "PY=py"
  )
)
if not defined PY if exist "%USERPROFILE%\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe" set "PY=%USERPROFILE%\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe"

if not defined PY goto nopy

set "GANTT_HIDE_CONSOLE=1"
"%PY%" "%~dp0window.py" %*
set "RC=%errorlevel%"
if not "%RC%"=="0" echo [ERROR] Window exited with code %RC%. Details: logs\window.log
endlocal & exit /b %RC%

:nopy
echo.
echo [ERROR] No Python with openpyxl found.
echo         Install Python 3 and run:  pip install openpyxl
echo         The window also needs pywebview: pip install pywebview
echo.
pause
exit /b 1
