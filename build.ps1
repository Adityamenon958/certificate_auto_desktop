# Build Certificate Auto desktop app (Windows)
# Run from project folder: .\build.ps1

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$venvActivate = Join-Path $PSScriptRoot "then\Scripts\Activate.ps1"
if (-not (Test-Path $venvActivate)) {
    Write-Error "Virtual env 'then' not found. Create it: python -m venv then"
}

Write-Host "Activating venv..."
. $venvActivate

Write-Host "Installing dependencies..."
pip install -q -r requirements.txt
pip install -q -r requirements-build.txt

Write-Host "Building with PyInstaller..."
pyinstaller certificate_auto.spec --noconfirm

$out = Join-Path $PSScriptRoot "dist\CertificateAuto"
if (Test-Path (Join-Path $out "CertificateAuto.exe")) {
    Write-Host ""
    Write-Host "Build OK: $out"
    Write-Host "Copy .env into that folder before sharing with client."
    Write-Host "See CLIENT_TESTING.md for packaging steps."
} else {
    Write-Error "Build failed - CertificateAuto.exe not found."
}
