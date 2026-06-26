# Создаёт ярлык «cvetopt» на рабочем столе.
# Предпочитает cvetopt.exe; если нет — cvetopt-launcher.vbs.
# Запуск: powershell -ExecutionPolicy Bypass -File scripts\create-desktop-shortcut.ps1
param(
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
)

$exe = Join-Path $ProjectRoot "cvetopt.exe"
$vbs = Join-Path $ProjectRoot "cvetopt-launcher.vbs"

if (Test-Path $exe) {
    $target = $exe
    $arguments = ""
} elseif (Test-Path $vbs) {
    $target = "$env:SystemRoot\System32\wscript.exe"
    $arguments = "`"$vbs`""
    Write-Warning "cvetopt.exe не найден — ярлык на VBS. Соберите exe: scripts\build-launcher-exe.ps1"
} else {
    Write-Error "Нет ни cvetopt.exe, ни cvetopt-launcher.vbs в $ProjectRoot"
}

$wsh = New-Object -ComObject WScript.Shell
$desktop = [Environment]::GetFolderPath("Desktop")
$lnkPath = Join-Path $desktop "cvetopt.lnk"

$shortcut = $wsh.CreateShortcut($lnkPath)
$shortcut.TargetPath = $target
$shortcut.Arguments = $arguments
$shortcut.WorkingDirectory = $ProjectRoot
$shortcut.Description = "cvetopt — скачивание отчётов и обработка инвойсов"
$shortcut.Save()

Write-Host "Готово: $lnkPath"
Write-Host "Двойной клик: откроет окно cvetopt. Закрытие окна остановит сервер."
