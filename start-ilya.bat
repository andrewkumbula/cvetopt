@echo off
REM Minimal start for Ilya - no "where", no browser helper. Shows errors clearly.
setlocal EnableExtensions
cd /d "%~dp0"
title cvetopt
set "PYTHONUNBUFFERED=1"
set "PLAYWRIGHT_BROWSERS_PATH=%~dp0ms-playwright"
echo [cvetopt] starting...
"%~dp0.venv\Scripts\python.exe" -u -m uvicorn cvetopt.app:app --host 127.0.0.1 --port 8000 --app-dir src --log-level info
echo [cvetopt] exited %ERRORLEVEL%
pause
