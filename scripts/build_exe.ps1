# 打包 SMART发票助手（目录版，自带 Python 运行时）
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot\..

Write-Host "Installing PyInstaller..."
python -m pip install -q pyinstaller

Write-Host "Building..."
python -m PyInstaller --noconfirm --clean SMART_Invoice_App.spec
if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed" }

$dist = Join-Path (Get-Location) "dist\SMART发票助手"
if (-not (Test-Path $dist)) {
    # fallback ascii folder if name mangled
    $cand = Get-ChildItem (Join-Path (Get-Location) "dist") -Directory | Select-Object -First 1
    if ($cand) { $dist = $cand.FullName } else { throw "dist folder missing" }
}

New-Item -ItemType Directory -Force -Path (Join-Path $dist "config") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $dist "templates") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $dist "input") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $dist "output") | Out-Null

Copy-Item "config\config.yaml" (Join-Path $dist "config\config.yaml") -Force
Copy-Item "templates\*" (Join-Path $dist "templates") -Recurse -Force
Copy-Item "使用说明.txt" (Join-Path $dist "使用说明.txt") -Force

Write-Host ""
Write-Host "DONE. Distribute this folder:"
Write-Host $dist
