@echo off
REM Recreate .venv for ALL Windows accounts (not tied to BananaMan AppData).
REM Run as Administrator. Needs "uv" in PATH.
setlocal EnableExtensions
cd /d "%~dp0"

echo [cvetopt] Project: %CD%

where uv >nul 2>nul
if errorlevel 1 (
  echo [cvetopt] ERROR: uv not found in PATH.
  echo [cvetopt] Open a console where "where uv" works, then re-run this file.
  pause
  exit /b 1
)

REM Shared Python inside the project - readable by Ilya and BananaMan.
set "SHARED_PY_DIR=%CD%\.python"
set "UV_PYTHON_INSTALL_DIR=%SHARED_PY_DIR%"

echo [cvetopt] Installing Python 3.11 into: %SHARED_PY_DIR%
uv python install 3.11
if errorlevel 1 (
  echo [cvetopt] ERROR: uv python install failed
  pause
  exit /b 1
)

if exist ".venv" (
  echo [cvetopt] Removing old .venv ...
  rmdir /s /q ".venv"
)

echo [cvetopt] Creating .venv with shared Python ...
set "UV_PYTHON=3.11"
uv sync
if errorlevel 1 (
  echo [cvetopt] ERROR: uv sync failed
  pause
  exit /b 1
)

echo [cvetopt] Checking pyvenv.cfg ...
type ".venv\pyvenv.cfg"
findstr /I /C:"\Users\BananaMan\" ".venv\pyvenv.cfg" >nul 2>nul
if not errorlevel 1 (
  echo [cvetopt] WARNING: venv still mentions BananaMan - may fail for Ilya
) else (
  echo [cvetopt] OK: venv is not bound to BananaMan AppData
)

echo [cvetopt] Playwright into shared folder ...
set "PLAYWRIGHT_BROWSERS_PATH=%CD%\ms-playwright"
uv run playwright install chromium

echo [cvetopt] Granting Ilya access ...
icacls "%CD%\.venv" /grant "server-cvetopt\Ilya:(OI)(CI)(M)" /T
icacls "%CD%\.python" /grant "server-cvetopt\Ilya:(OI)(CI)(M)" /T
if exist "%CD%\ms-playwright" icacls "%CD%\ms-playwright" /grant "server-cvetopt\Ilya:(OI)(CI)(M)" /T

echo.
echo [cvetopt] Done. Under Ilya run: cvetopt.bat
pause
