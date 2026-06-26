# Build cvetopt.exe (Windows app launcher).
# Run: powershell -ExecutionPolicy Bypass -File scripts\build-launcher-exe.ps1
param(
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
)

$ErrorActionPreference = "Stop"
Set-Location $ProjectRoot

$launcherPy = Join-Path $ProjectRoot "launcher\cvetopt_app.py"
if (-not (Test-Path $launcherPy)) {
    Write-Error "Missing $launcherPy"
}

$buildDir = Join-Path $ProjectRoot "build\launcher"
$oldExe = Join-Path $ProjectRoot "cvetopt.exe"
if (Test-Path $oldExe) {
    Remove-Item -Force $oldExe
}

Write-Host "==> Building cvetopt.exe (PyInstaller, onefile, noconsole)..."
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
    Write-Error "Build did not create cvetopt.exe"
}

Write-Host ""
Write-Host "OK: $oldExe"
Write-Host "Shortcut: powershell -ExecutionPolicy Bypass -File scripts\create-desktop-shortcut.ps1"
