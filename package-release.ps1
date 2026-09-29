param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[0-9]+\.[0-9]+\.[0-9]+$')]
    [string]$Version
)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$output = Join-Path $root 'dist'
$assetDir = Join-Path $root 'release-assets'
$staging = Join-Path ([System.IO.Path]::GetTempPath()) ("Docx2Pdf-release-" + [guid]::NewGuid().ToString('N'))
$zipPath = Join-Path $assetDir 'Docx2Pdf-win-x64.zip'
$hashPath = "$zipPath.sha256"

try {
    & (Join-Path $root 'build-windows.ps1') -Version $Version
    if ($LASTEXITCODE -ne 0) { throw "Build failed with exit code $LASTEXITCODE" }

    $exe = Join-Path $output 'Docx2Pdf.exe'
    if (-not (Test-Path $exe)) { throw "Expected executable was not produced: $exe" }
    New-Item -ItemType Directory -Force -Path $staging, $assetDir | Out-Null
    Copy-Item $exe, (Join-Path $root 'README.md'), (Join-Path $root 'LICENSE') -Destination $staging
    Set-Content -Path (Join-Path $staging 'VERSION.txt') -Value "Docx2Pdf v$Version - Windows x64" -Encoding UTF8
    if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
    Compress-Archive -Path (Join-Path $staging '*') -DestinationPath $zipPath -CompressionLevel Optimal

    $hash = (Get-FileHash $zipPath -Algorithm SHA256).Hash.ToLowerInvariant()
    Set-Content -Path $hashPath -Value "$hash  Docx2Pdf-win-x64.zip" -Encoding ASCII
    Write-Host "Release package: $zipPath" -ForegroundColor Green
    Write-Host "SHA-256: $hash" -ForegroundColor Green
}
finally {
    if (Test-Path $staging) { Remove-Item $staging -Recurse -Force }
}
