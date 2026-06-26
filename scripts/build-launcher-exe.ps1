# Сборка cvetopt.exe (лаунчер Windows-приложения).
# Запуск на сервере: powershell -ExecutionPolicy Bypass -File scripts\build-launcher-exe.ps1
param(
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
)

$ErrorActionPreference = "Stop"
Set-Location $ProjectRoot

$launcherPy = Join-Path $ProjectRoot "launcher\cvetopt_app.py"
if (-not (Test-Path $launcherPy)) {
    Write-Error "Не найден $launcherPy"
}

$buildDir = Join-Path $ProjectRoot "build\launcher"
$oldExe = Join-Path $ProjectRoot "cvetopt.exe"
if (Test-Path $oldExe) {
    Remove-Item -Force $oldExe
}

Write-Host "==> Сборка cvetopt.exe (PyInstaller, один файл, без консоли)…"
uv run --with pyinstaller pyinstaller `
    --onefile `
    --noconsole `
    --name cvetopt `
    --distpath $ProjectRoot `
    --workpath $buildDir `
    --specpath $buildDir `
    --clean `
    $launcherPy

if (-not (Test-Path $oldExe)) {
    Write-Error "Сборка не создала cvetopt.exe"
}

Write-Host ""
Write-Host "Готово: $oldExe"
Write-Host "Ярлык на рабочий стол: powershell -ExecutionPolicy Bypass -File scripts\create-desktop-shortcut.ps1"
