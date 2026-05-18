# Первичная установка cvetopt на Windows Server.
# Запуск (PowerShell от админа): Set-ExecutionPolicy -Scope Process Bypass; .\scripts\deploy-server.ps1
# Переменные: $RepoUrl — URL git-репозитория; $InstallDir — каталог (по умолчанию C:\Apps\cvetopt).

param(
    [Parameter(Mandatory = $true)]
    [string] $RepoUrl,
    [string] $InstallDir = "C:\Apps\cvetopt"
)

$ErrorActionPreference = "Stop"

Write-Host "==> Проверка uv и git..."
if (-not (Get-Command uv -ErrorAction SilentlyContinue)) {
    Write-Host "uv не найден. Установите: winget install astral-sh.uv"
    exit 1
}
if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    Write-Host "git не найден. Установите: winget install Git.Git"
    exit 1
}

$parent = Split-Path -Parent $InstallDir
if (-not (Test-Path $parent)) {
    New-Item -ItemType Directory -Force -Path $parent | Out-Null
}

if (Test-Path $InstallDir) {
    Write-Host "Каталог $InstallDir уже существует. Для обновления используйте git pull в нём или удалите каталог."
    exit 1
}

Write-Host "==> git clone $RepoUrl -> $InstallDir"
git clone $RepoUrl $InstallDir
Set-Location $InstallDir

Write-Host "==> uv sync"
uv sync

Write-Host "==> playwright install chromium (под текущей учёткой!)"
uv run playwright install chromium

Write-Host "==> Desktop-папки для Excel COM (Server)"
$dirs = @(
    "C:\Windows\System32\config\systemprofile\Desktop",
    "C:\Windows\SysWOW64\config\systemprofile\Desktop"
)
foreach ($d in $dirs) {
    if (-not (Test-Path $d)) {
        New-Item -ItemType Directory -Force -Path $d | Out-Null
        Write-Host "  создано: $d"
    }
}

if (-not (Test-Path ".env")) {
    Copy-Item ".env.example" ".env"
    Write-Host ""
    Write-Host "Создан .env из .env.example — ЗАПОЛНИТЕ логины и пароли в $InstallDir\.env"
} else {
    Write-Host ".env уже есть."
}

Write-Host ""
Write-Host "Дальше вручную:"
Write-Host "  1. Положите Auto_new.xls в $InstallDir"
Write-Host "  2. Отредактируйте .env (BIFLORICA_*, DELMIR_*, MAIL_*)"
Write-Host "  3. Один раз запустите Excel под сервисной учёткой"
Write-Host "  4. Запуск: двойной клик $InstallDir\cvetopt.bat"
Write-Host "  5. Автозапуск: Task Scheduler At log on -> cvetopt.bat (см. README_WIN.md)"
