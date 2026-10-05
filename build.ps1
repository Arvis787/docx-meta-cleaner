$ErrorActionPreference = 'Stop'
Push-Location $PSScriptRoot
try {
    python -m PyInstaller --noconfirm --clean --onefile --windowed --name DocxMetaCleaner-v1.2.0 --icon assets/app.png --add-data 'assets/app.png;assets' DocxMetaCleaner.py
    if ($LASTEXITCODE -ne 0) { throw 'PyInstaller failed.' }
} finally {
    Pop-Location
}
