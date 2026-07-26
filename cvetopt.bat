@echo off
REM cvetopt launcher for Windows Server.
REM Starts uvicorn on http://127.0.0.1:8000. Exit code 42 = git pull + restart.
setlocal EnableExtensions EnableDelayedExpansion

set "ROOT=%~dp0"
cd /d "%ROOT%"

title cvetopt - do not close this window

echo.
echo ============================================================
echo   cvetopt: server runs in THIS console window.
echo   Do NOT close this window while using the app.
echo   Closing it stops the server.
echo   To stop: press Ctrl+C here, wait until it exits.
echo   Tip: do not click inside this window - it pauses output.
echo ============================================================
echo.

if not defined PLAYWRIGHT_BROWSERS_PATH (
  set "PLAYWRIGHT_BROWSERS_PATH=%ROOT%ms-playwright"
)

set "UV_MODE="
set "UV_CMD="
set "VENV_PY=%ROOT%.venv\Scripts\python.exe"

REM Prefer project .venv FIRST - avoid "where uv/python" hanging on bad PATH entries.
echo [cvetopt] checking .venv ...
if exist "%VENV_PY%" (
  findstr /I /C:"\Users\" "%ROOT%.venv\pyvenv.cfg" >nul 2>nul
  if not errorlevel 1 (
    findstr /I /C:"\AppData\" "%ROOT%.venv\pyvenv.cfg" >nul 2>nul
    if not errorlevel 1 (
      echo [cvetopt] .venv points to per-user AppData Python - other accounts cannot use it.
      echo [cvetopt] Under admin run: fix-venv-for-all-users.bat
      if not "%CVETOPT_HIDDEN%"=="1" pause
      exit /b 1
    )
  )
  set "UV_MODE=venv"
  echo [cvetopt] mode: venv
  goto ready
)

echo [cvetopt] no .venv - looking for uv ...
where uv >nul 2>nul
if not errorlevel 1 (
  set "UV_MODE=uv"
  set "UV_CMD=uv"
  echo [cvetopt] mode: uv
  goto ready
)

echo [cvetopt] looking for python -m uv ...
where python >nul 2>nul
if not errorlevel 1 (
  python -m uv --version >nul 2>nul
  if not errorlevel 1 (
    set "UV_MODE=uv"
    set "UV_CMD=python -m uv"
    echo [cvetopt] mode: python -m uv
    goto ready
  )
)

echo [cvetopt] Missing uv, python, and .venv\Scripts\python.exe
echo [cvetopt] Under admin: fix-venv-for-all-users.bat or install uv system-wide
if not "%CVETOPT_HIDDEN%"=="1" pause
exit /b 1

:ready
reg add "HKCU\Console" /v QuickEdit /t REG_DWORD /d 0 /f >nul 2>nul

set "PYTHONUNBUFFERED=1"
set "PYTHONIOENCODING=utf-8"

if not "%CVETOPT_NO_BROWSER%"=="1" (
  start "" /b cmd /c "timeout /t 3 /nobreak >nul & start "" http://127.0.0.1:8000/"
)

:loop
echo.
echo [cvetopt] %DATE% %TIME% - starting uvicorn (Ctrl+C to stop)
echo [cvetopt] mode=%UV_MODE%  open http://127.0.0.1:8000/
echo [cvetopt] loading Python modules - wait up to 60 sec...
if "%UV_MODE%"=="venv" (
  "%VENV_PY%" -u -m uvicorn cvetopt.app:app --host 127.0.0.1 --port 8000 --app-dir src --log-level info
) else (
  %UV_CMD% run python -u -m uvicorn cvetopt.app:app --host 127.0.0.1 --port 8000 --app-dir src --log-level info
)
set "EXIT_CODE=%ERRORLEVEL%"

if "%EXIT_CODE%"=="42" (
  echo [cvetopt] Update requested. git pull...
  if exist ".git\" (
    git pull --ff-only
  ) else (
    echo [cvetopt] .git not found - skip git pull
  )
  if "%UV_MODE%"=="venv" (
    echo [cvetopt] .venv mode: looking for uv to sync...
    where uv >nul 2>nul
    if not errorlevel 1 (
      uv sync
      uv run playwright install chromium
    ) else (
      echo [cvetopt] uv not found - sync skipped. Run sync under admin later.
    )
  ) else (
    echo [cvetopt] uv sync...
    %UV_CMD% sync
    echo [cvetopt] Playwright Chromium...
    %UV_CMD% run playwright install chromium
  )
  echo [cvetopt] Restarting...
  goto loop
)

echo [cvetopt] uvicorn exited with code %EXIT_CODE%
if not "%CVETOPT_HIDDEN%"=="1" pause
exit /b %EXIT_CODE%
