@echo off
REM Grant Ilya full access to project runtime folders. Run as BananaMan / Admin.
setlocal EnableExtensions
cd /d "%~dp0"

set "USER=server-cvetopt\Ilya"

echo [cvetopt] Granting %USER% on runtime folders...
echo [cvetopt] Stopping python/cvetopt that may lock files...
taskkill /F /IM cvetopt.exe >nul 2>nul
for /f "tokens=5" %%P in ('netstat -aon ^| findstr /R /C:":8000 .*LISTENING"') do taskkill /F /PID %%P >nul 2>nul
taskkill /F /IM python.exe >nul 2>nul
timeout /t 2 /nobreak >nul

for %%D in (".venv" ".python" "src" "data" "ms-playwright" ".") do (
  if exist "%%~D" (
    echo.
    echo [cvetopt] == %%~D
    takeown /F "%%~D" /R /D Y >nul 2>nul
    icacls "%%~D" /reset /T /C >nul 2>nul
    icacls "%%~D" /inheritance:e /T /C >nul 2>nul
    icacls "%%~D" /grant "Administrators:(OI)(CI)F" /T /C
    icacls "%%~D" /grant "%USER%:(OI)(CI)(M)" /T /C
  )
)

echo.
echo [cvetopt] Sample ACL on uvicorn:
icacls ".venv\Lib\site-packages\uvicorn\__init__.py"
echo.
echo [cvetopt] Done. Under Ilya run:
echo   type .venv\Lib\site-packages\uvicorn\__init__.py
echo   cvetopt.bat
pause
