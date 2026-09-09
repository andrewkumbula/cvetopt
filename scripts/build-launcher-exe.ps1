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

$buildDir = Join-Path $env:TEMP "cvetopt-build"
$distDir = Join-Path $buildDir "dist"
$newExe = Join-Path $distDir "cvetopt.exe"
$finalExe = Join-Path $ProjectRoot "cvetopt.exe"
if (Test-Path $buildDir) {
    Remove-Item -Recurse -Force $buildDir
}

Write-Host "==> Building cvetopt.exe (PyInstaller, onefile, noconsole)..."
uv run --with pyinstaller pyinstaller `
    --onefile `
    --noconsole `
    --name cvetopt `
    --distpath $distDir `
    --workpath $buildDir `
    --specpath $buildDir `
    --clean `
    $launcherPy

if (-not (Test-Path $newExe)) {
    Write-Error "Build did not create cvetopt.exe - old $finalExe was left untouched"
}

# Only replace the old exe now that the new build is confirmed to exist.
Copy-Item -Force $newExe $finalExe
Remove-Item -Recurse -Force $buildDir

Write-Host ""
Write-Host "OK: $finalExe"
Write-Host "Shortcut: powershell -ExecutionPolicy Bypass -File scripts\create-desktop-shortcut.ps1"
