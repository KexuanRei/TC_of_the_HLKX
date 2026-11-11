@echo off
setlocal
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"

set "PY="
where python >nul 2>nul
if not errorlevel 1 set "PY=python"
if not defined PY if exist "%USERPROFILE%\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe" set "PY=%USERPROFILE%\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe"

if not defined PY (
  echo.
  echo [ERROR] Python 3 not found.
  echo.
  pause
  exit /b 1
)

"%PY%" menu.py
if errorlevel 1 pause
endlocal
