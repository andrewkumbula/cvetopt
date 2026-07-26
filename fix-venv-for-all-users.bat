@echo off
REM Recreate .venv so it does NOT point to C:\Users\<admin>\AppData\...
REM Run as Administrator (BananaMan), then Ilya can start cvetopt.
setlocal EnableExtensions
cd /d "%~dp0"

echo [cvetopt] Project: %CD%
echo [cvetopt] Looking for a machine-wide Python (not per-user AppData)...

set "PY="
if exist "%ProgramFiles%\Python311\python.exe" set "PY=%ProgramFiles%\Python311\python.exe"
if not defined PY if exist "%ProgramFiles%\Python312\python.exe" set "PY=%ProgramFiles%\Python312\python.exe"
if not defined PY if exist "%LocalAppData%\Programs\Python\Python311\python.exe" set "PY=%LocalAPPDATA%\Programs\Python\Python311\python.exe"

where uv >nul 2>nul
if errorlevel 1 (
  echo [cvetopt] uv not in PATH. Open a console where "where uv" works, or install uv system-wide.
  pause
  exit /b 1
)

if not defined PY (
  echo [cvetopt] No Python in Program Files.
  echo [cvetopt] Install Python 3.11+ with "Install for all users" checked, then re-run this bat.
  echo [cvetopt] Or set PY to full path: set PY=C:\path\to\python.exe
  pause
  exit /b 1
)

echo [cvetopt] Using Python: %PY%
"%PY%" -c "import sys; print(sys.version)"

if exist ".venv\" (
  echo [cvetopt] Removing old .venv (bound to admin AppData)...
  rmdir /s /q ".venv"
)

echo [cvetopt] uv sync with shared Python...
set "UV_PYTHON=%PY%"
uv sync
if errorlevel 1 (
  echo [cvetopt] uv sync failed
  pause
  exit /b 1
)

echo [cvetopt] Playwright browsers into shared folder...
set "PLAYWRIGHT_BROWSERS_PATH=%CD%\ms-playwright"
uv run playwright install chromium

echo [cvetopt] Grant Ilya access...
icacls "%CD%\.venv" /grant "server-cvetopt\Ilya:(OI)(CI)(M)" /T >nul
if exist "%CD%\ms-playwright" icacls "%CD%\ms-playwright" /grant "server-cvetopt\Ilya:(OI)(CI)(M)" /T >nul

echo.
echo [cvetopt] OK. Now under Ilya run: C:\Apps\cvetopt\cvetopt.bat
pause
