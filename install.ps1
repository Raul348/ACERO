# Instala ACERO Refuerzo (compilado en .\dist) en las versiones de Revit indicadas que estén instaladas.
#   .\install.ps1                  -> 2024..2027 (las que existan en %AppData%\Autodesk\Revit\Addins)
#   .\install.ps1 -Versions 2025   -> solo Revit 2025
# Requiere haber ejecutado antes .\build.ps1 (o descargar el paquete de GitHub Actions en .\dist).
param([string[]]$Versions = @("2024", "2025", "2026", "2027"))
$ErrorActionPreference = "Stop"

foreach ($v in $Versions) {
    $src = Join-Path $PSScriptRoot "dist\Revit$v"
    $addins = Join-Path $env:APPDATA "Autodesk\Revit\Addins\$v"
    if (-not (Test-Path $src)) { Write-Host "Revit ${v}: no hay paquete en $src (ejecute build.ps1)" -ForegroundColor Yellow; continue }
    if (-not (Test-Path $addins)) { Write-Host "Revit ${v}: no está instalado, se omite" -ForegroundColor DarkGray; continue }

    foreach ($manifest in Get-ChildItem $src -Filter *.addin) {
        $name = $manifest.BaseName
        New-Item -ItemType Directory -Force -Path (Join-Path $addins $name) | Out-Null
        Copy-Item (Join-Path $src "$name\$name.dll") (Join-Path $addins $name) -Force
        Copy-Item $manifest.FullName $addins -Force
        Write-Host "Revit ${v}: $name instalado" -ForegroundColor Green
    }
}
