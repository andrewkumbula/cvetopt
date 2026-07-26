@echo off
REM Debug launcher WITH console - shows errors. Run as BananaMan or Ilya.
setlocal EnableExtensions
cd /d "%~dp0"
echo [cvetopt] debug launcher from %CD%
echo.

if exist ".venv\Scripts\python.exe" (
  ".venv\Scripts\python.exe" -u "launcher\cvetopt_app.py"
) else (
  echo [cvetopt] missing .venv\Scripts\python.exe
  pause
  exit /b 1
)

echo.
echo [cvetopt] exit code %ERRORLEVEL%
echo [cvetopt] logs: data\launcher.log  and  data\launcher-server.log
pause
