# Copies docx2pdf to %LOCALAPPDATA%\docx2pdf and adds it to the user PATH.
# Run once:  powershell -ExecutionPolicy Bypass -File install.ps1

$target = Join-Path $env:LOCALAPPDATA "docx2pdf"
New-Item -ItemType Directory -Force -Path $target | Out-Null
Copy-Item -Path (Join-Path $PSScriptRoot "docx2pdf.ps1"), (Join-Path $PSScriptRoot "docx2pdf.cmd") -Destination $target -Force

$userPath = [Environment]::GetEnvironmentVariable("Path", "User")
if ($userPath -notlike "*$target*") {
    [Environment]::SetEnvironmentVariable("Path", "$userPath;$target", "User")
    Write-Host "Added $target to your user PATH." -ForegroundColor Green
}

Write-Host "Installed. Open a new terminal and run:  docx2pdf <folder>" -ForegroundColor Cyan
