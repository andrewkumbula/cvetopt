@echo off
REM Лаунчер cvetopt для Windows Server.
REM Двойной клик: поднимает локальный uvicorn, открывает браузер на http://127.0.0.1:8000.
REM Цикл while — после команды "обновить программу" приложение само рестартует на новой версии.
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

REM uv в PATH или через python -m uv (типично после pip install uv на Server 2019).
set "UV_CMD=uv"
where uv >nul 2>nul
if errorlevel 1 (
  where python >nul 2>nul
  if errorlevel 1 (
    echo [cvetopt] Не найдены ни uv, ни python. Установите Python 3.11+ и: python -m pip install uv
    if not "%CVETOPT_HIDDEN%"=="1" pause
    exit /b 1
  )
  set "UV_CMD=python -m uv"
  echo [cvetopt] uv не в PATH — использую: python -m uv
)

REM Открываем браузер один раз через 3 секунды (uvicorn ещё стартует).
REM Лаунчер cvetopt-launcher.vbs сам открывает окно — не дублируем.
if not "%CVETOPT_NO_BROWSER%"=="1" (
  start "" /b cmd /c "timeout /t 3 /nobreak >nul & start "" http://127.0.0.1:8000/"
)

:loop
echo.
echo [cvetopt] %DATE% %TIME% — запускаю uvicorn (Ctrl+C для выхода)
%UV_CMD% run uvicorn cvetopt.app:app --host 127.0.0.1 --port 8000 --app-dir src
set "EXIT_CODE=%ERRORLEVEL%"

REM Код выхода 42 — наш «обновись и перезапустись». Любой другой код = ручной Ctrl+C / краш.
if "%EXIT_CODE%"=="42" (
  echo [cvetopt] Получен запрос на обновление. Делаю git pull + uv sync…
  if exist ".git\" (
    git pull --ff-only
  ) else (
    echo [cvetopt] .git не найден. Пропускаю git pull.
  )
  %UV_CMD% sync
  echo [cvetopt] Проверяю Chromium для Playwright…
  %UV_CMD% run playwright install chromium
  echo [cvetopt] Перезапуск…
  goto loop
)

REM Любой другой код — выходим, не зацикливаемся.
echo [cvetopt] uvicorn завершился с кодом %EXIT_CODE%. Выход.
if not "%CVETOPT_HIDDEN%"=="1" pause
exit /b %EXIT_CODE%
