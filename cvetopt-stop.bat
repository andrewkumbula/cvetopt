@echo off
REM Остановка cvetopt, запущенного скрыто (cvetopt-hidden.vbs).
REM Завершает процесс, слушающий порт 8000 (uvicorn/python).
setlocal EnableExtensions

set "FOUND="
for /f "tokens=5" %%P in ('netstat -aon ^| findstr /R /C:":8000 .*LISTENING"') do (
  set "FOUND=1"
  echo [cvetopt] Останавливаю процесс PID %%P на порту 8000…
  taskkill /F /PID %%P >nul 2>nul
)

if not defined FOUND (
  echo [cvetopt] Сервер на порту 8000 не найден — возможно уже остановлен.
)

echo [cvetopt] Готово.
timeout /t 2 /nobreak >nul
