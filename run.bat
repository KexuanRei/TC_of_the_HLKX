@echo off
setlocal
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"

rem --- find a python that actually has openpyxl ---
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

if not defined PY (
  echo.
  echo [ERROR] No Python with openpyxl found.
  echo         Install Python 3 and run:  pip install openpyxl
  echo.
  pause
  exit /b 1
)

echo Using Python: %PY%
"%PY%" build.py %*
if errorlevel 1 (
  echo.
  echo [ERROR] build.py failed. See the message above.
  pause
  exit /b 1
)

echo.
echo [OK] data.js generated.
if not "%GANTT_NO_OPEN%"=="1" start "" "viewer.html"
endlocal
