param(
    [ValidatePattern('^[0-9]+\.[0-9]+\.[0-9]+$')]
    [string]$Version
)

$ErrorActionPreference = 'Stop'

$project = Join-Path $PSScriptRoot 'Docx2PdfGui.csproj'
$output = Join-Path $PSScriptRoot 'dist'

if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) {
    throw 'The .NET 8 SDK is required. Install it from https://dotnet.microsoft.com/download/dotnet/8.0'
}

New-Item -ItemType Directory -Force -Path $output | Out-Null
$publishArgs = @(
    'publish', $project, '-c', 'Release', '-r', 'win-x64', '--self-contained', 'true',
    '-p:PublishSingleFile=true', '-p:IncludeNativeLibrariesForSelfExtract=true',
    '-p:EnableCompressionInSingleFile=true', '-o', $output
)
if ($Version) { $publishArgs += "-p:Version=$Version" }
dotnet @publishArgs

if ($LASTEXITCODE -ne 0) {
    throw "dotnet publish failed with exit code $LASTEXITCODE"
}

$exe = Join-Path $output 'Docx2Pdf.exe'
if (-not (Test-Path $exe)) {
    throw "Build reported success but the executable is missing: $exe"
}
Get-ChildItem $output -Filter '*.pdb' -File | Remove-Item -Force
Write-Host "Built: $exe" -ForegroundColor Green
