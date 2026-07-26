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

echo [cvetopt] Closing running launcher, if any ...
taskkill /f /im cvetopt.exe >nul 2>nul
taskkill /f /im cvetopt-debug.exe >nul 2>nul

REM Work folder in per-user TEMP: project build\ may belong to another Windows user.
set "WORKDIR=%TEMP%\cvetopt-build"
if exist "%WORKDIR%" rmdir /s /q "%WORKDIR%" 2>nul
mkdir "%WORKDIR%" 2>nul

REM Old build\ is not used anymore; drop it if we still may.
if exist "build\launcher" rmdir /s /q "build" 2>nul
if exist "build\launcher" echo [cvetopt] Note: stale build\ folder from another user - can be deleted by Administrator.

call :unlock "cvetopt.exe"
call :unlock "cvetopt-debug.exe"
if exist "cvetopt.exe" (
  echo [cvetopt] Cannot replace cvetopt.exe - it is locked or owned by another user.
  echo [cvetopt] Close the app on all accounts, then run this script as Administrator.
  pause
  exit /b 1
)

echo [cvetopt] Folder: %CD%
echo [cvetopt] Work folder: %WORKDIR%
echo [cvetopt] Building cvetopt.exe with PyInstaller...
echo.

%UV_CMD% run --with pyinstaller pyinstaller ^
  --onefile ^
  --noconsole ^
  --name cvetopt ^
  --distpath "%CD%" ^
  --workpath "%WORKDIR%" ^
  --specpath "%WORKDIR%" ^
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
echo [cvetopt] Also building console debug exe...

%UV_CMD% run --with pyinstaller pyinstaller ^
  --onefile ^
  --console ^
  --name cvetopt-debug ^
  --distpath "%CD%" ^
  --workpath "%WORKDIR%" ^
  --specpath "%WORKDIR%" ^
  "%CD%\launcher\cvetopt_app.py"

if exist "cvetopt-debug.exe" (
  echo [cvetopt] OK: %CD%\cvetopt-debug.exe  - run this to see errors in console
)

REM Both accounts must be able to start the exe.
icacls "cvetopt.exe" /grant "*S-1-5-32-545:(RX)" /q >nul 2>nul
if exist "cvetopt-debug.exe" icacls "cvetopt-debug.exe" /grant "*S-1-5-32-545:(RX)" /q >nul 2>nul

rmdir /s /q "%WORKDIR%" 2>nul

echo [cvetopt] Desktop shortcut: scripts\create-desktop-shortcut.ps1
pause
exit /b 0

:unlock
if not exist "%~1" exit /b 0
del /f /q "%~1" 2>nul
if not exist "%~1" exit /b 0
echo [cvetopt] %~1 is protected - taking ownership ...
takeown /f "%~1" >nul 2>nul
icacls "%~1" /grant "*S-1-5-32-544:F" /q >nul 2>nul
del /f /q "%~1" 2>nul
exit /b 0
