@echo off
REM Build cvetopt.exe in project root. Double-click or run from cmd.
setlocal EnableExtensions
cd /d "%~dp0"

if not exist "launcher\cvetopt_app.py" (
  echo [cvetopt] Missing launcher\cvetopt_app.py - run: git pull
  pause
  exit /b 1
)

set "UV_CMD=uv"
where uv >nul 2>nul
if errorlevel 1 (
  where python >nul 2>nul
  if errorlevel 1 (
    echo [cvetopt] Need uv or python in PATH
    pause
    exit /b 1
  )
  set "UV_CMD=python -m uv"
)

if exist "cvetopt.exe" del /f "cvetopt.exe"

if not exist "build\launcher" mkdir "build\launcher"

echo [cvetopt] Folder: %CD%
echo [cvetopt] Building cvetopt.exe with PyInstaller...
echo.

%UV_CMD% run --with pyinstaller pyinstaller ^
  --onefile ^
  --noconsole ^
  --name cvetopt ^
  --distpath "%CD%" ^
  --workpath "%CD%\build\launcher" ^
  --specpath "%CD%\build\launcher" ^
  --clean ^
  "%CD%\launcher\cvetopt_app.py"

if errorlevel 1 (
  echo.
  echo [cvetopt] Build failed.
  pause
  exit /b 1
)

if not exist "cvetopt.exe" (
  echo [cvetopt] cvetopt.exe was not created.
  pause
  exit /b 1
)

echo.
echo [cvetopt] OK: %CD%\cvetopt.exe
echo [cvetopt] Desktop shortcut: scripts\create-desktop-shortcut.ps1
pause
