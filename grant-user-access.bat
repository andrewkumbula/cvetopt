@echo off
REM Grant Ilya (or other user) access to cvetopt. Run as Administrator.
setlocal EnableExtensions
cd /d "%~dp0"

if "%~1"=="" (
  set "USER_ARG=.\Ilya"
) else (
  set "USER_ARG=%~1"
)

echo [cvetopt] Grant access to %USER_ARG% ...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\grant-user-access.ps1" -User "%USER_ARG%" -ProjectRoot "%~dp0"
set "EC=%ERRORLEVEL%"
if %EC% NEQ 0 (
  echo [cvetopt] Failed, code %EC%
  pause
  exit /b %EC%
)
pause
