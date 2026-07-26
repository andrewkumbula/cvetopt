@echo off
REM Build cvetopt.exe in project root. Double-click or run from cmd.
setlocal EnableExtensions
cd /d "%~dp0"

REM Replacing an exe owned by another Windows account needs elevation.
fltmc >nul 2>nul
if errorlevel 1 (
  echo [cvetopt] Need Administrator rights - restarting elevated ...
  powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
  exit /b 0
)

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
if exist "build\launcher" echo [cvetopt] Note: stale build\ folder left over - delete C:\Apps\cvetopt\build manually.

echo [cvetopt] Folder: %CD%
echo [cvetopt] Work folder: %WORKDIR%
echo [cvetopt] Building cvetopt.exe with PyInstaller...
echo.

%UV_CMD% run --with pyinstaller pyinstaller ^
  --onefile ^
  --noconsole ^
  --name cvetopt ^
  --distpath "%WORKDIR%\dist" ^
  --workpath "%WORKDIR%" ^
  --specpath "%WORKDIR%" ^
  --clean ^
  "%CD%\launcher\cvetopt_app.py"

if errorlevel 1 goto build_failed
if not exist "%WORKDIR%\dist\cvetopt.exe" goto build_failed

echo.
echo [cvetopt] Building console debug exe...
%UV_CMD% run --with pyinstaller pyinstaller ^
  --onefile ^
  --console ^
  --name cvetopt-debug ^
  --distpath "%WORKDIR%\dist" ^
  --workpath "%WORKDIR%" ^
  --specpath "%WORKDIR%" ^
  "%CD%\launcher\cvetopt_app.py"

echo.
call :deploy "cvetopt.exe"
if errorlevel 1 goto deploy_failed
call :deploy "cvetopt-debug.exe"

rmdir /s /q "%WORKDIR%" 2>nul

echo.
echo [cvetopt] OK: %CD%\cvetopt.exe
echo [cvetopt] cvetopt-debug.exe shows errors in a console window.
echo [cvetopt] Desktop shortcut: scripts\create-desktop-shortcut.ps1
pause
exit /b 0

:build_failed
echo.
echo [cvetopt] Build failed.
pause
exit /b 1

:deploy_failed
echo.
echo [cvetopt] Could not replace cvetopt.exe. Who holds it:
tasklist /v /fi "imagename eq cvetopt.exe"
icacls "cvetopt.exe"
echo [cvetopt] New build kept here: %WORKDIR%\dist\cvetopt.exe
pause
exit /b 1

REM Copy fresh exe over the old one, clearing ownership/locks if needed.
:deploy
if not exist "%WORKDIR%\dist\%~1" exit /b 0
if exist "%~1" (
  del /f /q "%~1" >nul 2>nul
  if exist "%~1" (
    echo [cvetopt] %~1 is protected - taking ownership ...
    takeown /f "%~1" >nul 2>nul
    icacls "%~1" /reset /q >nul 2>nul
    icacls "%~1" /grant "*S-1-5-32-544:F" /q >nul 2>nul
    del /f /q "%~1" >nul 2>nul
  )
)
copy /y "%WORKDIR%\dist\%~1" "%~1" >nul 2>nul
if not exist "%~1" exit /b 1
REM Both Windows accounts must be able to start it.
icacls "%~1" /grant "*S-1-5-32-545:(RX)" /q >nul 2>nul
echo [cvetopt] Updated: %~1
exit /b 0
