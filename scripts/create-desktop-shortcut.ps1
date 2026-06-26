# Создаёт ярлык «cvetopt» на рабочем столе (лаунчер + иконка).
# Запуск: powershell -ExecutionPolicy Bypass -File scripts\create-desktop-shortcut.ps1
param(
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
)

$launcher = Join-Path $ProjectRoot "cvetopt-launcher.vbs"
if (-not (Test-Path $launcher)) {
    Write-Error "Не найден $launcher"
    exit 1
}

$wsh = New-Object -ComObject WScript.Shell
$desktop = [Environment]::GetFolderPath("Desktop")
$lnkPath = Join-Path $desktop "cvetopt.lnk"

$shortcut = $wsh.CreateShortcut($lnkPath)
$shortcut.TargetPath = "$env:SystemRoot\System32\wscript.exe"
$shortcut.Arguments = "`"$launcher`""
$shortcut.WorkingDirectory = $ProjectRoot
$shortcut.Description = "cvetopt — скачивание отчётов и обработка инвойсов"
$shortcut.Save()

Write-Host "Готово: $lnkPath"
Write-Host "Двойной клик: откроет окно cvetopt. Закрытие окна остановит сервер."
