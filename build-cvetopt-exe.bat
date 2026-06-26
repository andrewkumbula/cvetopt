@echo off
REM Сборка cvetopt.exe в корне проекта. Запуск: двойной клик или из cmd.
setlocal EnableExtensions
cd /d "%~dp0"

if not exist "scripts\build-launcher-exe.ps1" (
  echo [cvetopt] Не найден scripts\build-launcher-exe.ps1
  echo Сначала обновите проект: git pull  ^(или кнопка «Обновить программу»^)
  pause
  exit /b 1
)

if not exist "launcher\cvetopt_app.py" (
  echo [cvetopt] Не найден launcher\cvetopt_app.py — нужен git pull с последней версией.
  pause
  exit /b 1
)

echo [cvetopt] Папка: %CD%
echo [cvetopt] Сборка cvetopt.exe …
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\build-launcher-exe.ps1"
set "EC=%ERRORLEVEL%"
if %EC% NEQ 0 (
  echo [cvetopt] Ошибка сборки, код %EC%
  pause
  exit /b %EC%
)

if exist "%~dp0cvetopt.exe" (
  echo.
  echo [cvetopt] Готово: %~dp0cvetopt.exe
) else (
  echo [cvetopt] cvetopt.exe не появился — см. сообщения выше.
  pause
  exit /b 1
)

pause
