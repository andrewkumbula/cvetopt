@echo off
REM Лаунчер cvetopt для Windows Server.
REM Двойной клик / cvetopt.exe: поднимает uvicorn на http://127.0.0.1:8000.
REM Цикл while — после «обновить программу» (exit 42) делает git pull и перезапуск.
setlocal EnableExtensions EnableDelayedExpansion

set "ROOT=%~dp0"
cd /d "%ROOT%"

title cvetopt - do not close this window

echo.
echo ============================================================
echo   cvetopt: сервер работает в ЭТОМ окне консоли.
echo   Пока пользуетесь сайтом в браузере — НЕ ЗАКРЫВАЙТЕ окно.
echo   Если закрыть окно, сервер остановится и сайт перестанет открываться.
echo   Завершить работу: в этом окне нажмите Ctrl+C, дождитесь остановки.
echo ============================================================
echo.

REM Общий кэш браузеров Playwright (чтобы работало у всех учёток, не только у админа).
if not defined PLAYWRIGHT_BROWSERS_PATH (
  set "PLAYWRIGHT_BROWSERS_PATH=%ROOT%ms-playwright"
)

REM Режим запуска:
REM   UV_MODE=uv      — "uv run ..." (предпочтительно)
REM   UV_MODE=venv    — ".venv\Scripts\python.exe -m uvicorn ..." (для учётки без uv в PATH)
set "UV_MODE="
set "UV_CMD="
set "VENV_PY=%ROOT%.venv\Scripts\python.exe"

where uv >nul 2>nul
if not errorlevel 1 (
  set "UV_MODE=uv"
  set "UV_CMD=uv"
  echo [cvetopt] Режим: uv
  goto ready
)

where python >nul 2>nul
if not errorlevel 1 (
  python -m uv --version >nul 2>nul
  if not errorlevel 1 (
    set "UV_MODE=uv"
    set "UV_CMD=python -m uv"
    echo [cvetopt] Режим: python -m uv
    goto ready
  )
)

if exist "%VENV_PY%" (
  set "UV_MODE=venv"
  echo [cvetopt] uv/python не в PATH — использую .venv: %VENV_PY%
  goto ready
)

echo [cvetopt] Не найдены uv, python и .venv\Scripts\python.exe
echo [cvetopt] Под админом: установите uv в SYSTEM PATH или выполните uv sync в C:\Apps\cvetopt
if not "%CVETOPT_HIDDEN%"=="1" pause
exit /b 1

:ready
REM Открываем браузер один раз через 3 секунды (uvicorn ещё стартует).
REM Лаунчер cvetopt.exe / cvetopt-launcher.vbs сам открывает окно — не дублируем.
if not "%CVETOPT_NO_BROWSER%"=="1" (
  start "" /b cmd /c "timeout /t 3 /nobreak >nul & start "" http://127.0.0.1:8000/"
)

:loop
echo.
echo [cvetopt] %DATE% %TIME% — запускаю uvicorn (Ctrl+C для выхода)
if "%UV_MODE%"=="venv" (
  "%VENV_PY%" -m uvicorn cvetopt.app:app --host 127.0.0.1 --port 8000 --app-dir src
) else (
  %UV_CMD% run uvicorn cvetopt.app:app --host 127.0.0.1 --port 8000 --app-dir src
)
set "EXIT_CODE=%ERRORLEVEL%"

REM Код выхода 42 — «обновись и перезапустись».
if "%EXIT_CODE%"=="42" (
  echo [cvetopt] Получен запрос на обновление. Делаю git pull…
  if exist ".git\" (
    git pull --ff-only
  ) else (
    echo [cvetopt] .git не найден. Пропускаю git pull.
  )
  if "%UV_MODE%"=="venv" (
    echo [cvetopt] Режим .venv: ищу uv для sync…
    where uv >nul 2>nul
    if not errorlevel 1 (
      uv sync
      uv run playwright install chromium
    ) else (
      echo [cvetopt] uv не найден — sync пропущен. Зависимости обновятся при следующем запуске под админом.
    )
  ) else (
    echo [cvetopt] uv sync…
    %UV_CMD% sync
    echo [cvetopt] Playwright Chromium…
    %UV_CMD% run playwright install chromium
  )
  echo [cvetopt] Перезапуск…
  goto loop
)

echo [cvetopt] uvicorn завершился с кодом %EXIT_CODE%. Выход.
if not "%CVETOPT_HIDDEN%"=="1" pause
exit /b %EXIT_CODE%
