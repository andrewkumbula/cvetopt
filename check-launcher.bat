@echo off
REM Collect everything needed to diagnose "exe does not start the server".
setlocal EnableExtensions
cd /d "%~dp0"

echo ==== who / where ====
whoami
echo folder: %CD%
echo.

echo ==== files ====
if exist "cvetopt.exe" (echo cvetopt.exe: yes) else (echo cvetopt.exe: MISSING)
if exist "cvetopt-debug.exe" (echo cvetopt-debug.exe: yes) else (echo cvetopt-debug.exe: MISSING)
if exist "cvetopt.bat" (echo cvetopt.bat: yes) else (echo cvetopt.bat: MISSING)
if exist ".venv\Scripts\python.exe" (echo .venv python: yes) else (echo .venv python: MISSING)
echo.

echo ==== git ====
git log --oneline -1
echo.

echo ==== port 8000 ====
netstat -ano | findstr ":8000"
echo.

echo ==== processes ====
tasklist /fi "imagename eq cvetopt.exe" 2>nul | findstr /I cvetopt
tasklist /fi "imagename eq python.exe" 2>nul | findstr /I python
echo.

echo ==== data\launcher.log (last lines) ====
if exist "data\launcher.log" (
  powershell -NoProfile -Command "Get-Content -Tail 40 -LiteralPath 'data\launcher.log'"
) else (
  echo no data\launcher.log - launcher never started
)
echo.

echo ==== data\launcher-server.log (last lines) ====
if exist "data\launcher-server.log" (
  powershell -NoProfile -Command "Get-Content -Tail 40 -LiteralPath 'data\launcher-server.log'"
) else (
  echo no data\launcher-server.log - server was never spawned
)
echo.

echo ==== done: copy everything above ====
pause
