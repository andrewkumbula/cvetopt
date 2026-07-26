@echo off
REM Let every Windows account of this server read/write project data (logs, state, sessions).
setlocal EnableExtensions
cd /d "%~dp0"

fltmc >nul 2>nul
if errorlevel 1 (
  echo [cvetopt] Need Administrator rights - restarting elevated ...
  powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
  exit /b 0
)

echo [cvetopt] Folder: %CD%
echo [cvetopt] Granting Users modify rights on project files ...
icacls "%CD%" /grant "*S-1-5-32-545:(OI)(CI)M" /t /c /q >nul 2>nul

if exist "data" (
  echo [cvetopt] Resetting inherited rights inside data\ ...
  takeown /f "data" /r /d y >nul 2>nul
  icacls "data" /reset /t /c /q >nul 2>nul
  icacls "data" /grant "*S-1-5-32-545:(OI)(CI)M" /t /c /q >nul 2>nul
)

if exist "C:\Invoice" (
  echo [cvetopt] Granting Users modify rights on C:\Invoice ...
  icacls "C:\Invoice" /grant "*S-1-5-32-545:(OI)(CI)M" /t /c /q >nul 2>nul
)

echo.
echo [cvetopt] Done. Start the app again under the working account.
echo [cvetopt] For the sklad folder run: scripts\grant-user-access.ps1
pause
