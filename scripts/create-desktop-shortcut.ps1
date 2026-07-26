# Desktop shortcut for cvetopt (prefers cvetopt.exe, else cvetopt-launcher.vbs).
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
    Write-Warning "cvetopt.exe not found - using VBS. Run build-cvetopt-exe.bat first."
} else {
    Write-Error "Missing cvetopt.exe and cvetopt-launcher.vbs in $ProjectRoot"
}

$wsh = New-Object -ComObject WScript.Shell
$desktop = [Environment]::GetFolderPath("Desktop")
$lnkPath = Join-Path $desktop "cvetopt.lnk"

$shortcut = $wsh.CreateShortcut($lnkPath)
$shortcut.TargetPath = $target
$shortcut.Arguments = $arguments
$shortcut.WorkingDirectory = $ProjectRoot
$shortcut.Description = "cvetopt"
$shortcut.Save()

Write-Host "OK: $lnkPath"
Write-Host "Double-click starts server + app window; closing the window stops the server."
if (Test-Path (Join-Path $ProjectRoot "cvetopt.exe")) {
    Write-Host "Tip: rebuild exe after launcher changes: build-cvetopt-exe.bat"
}
