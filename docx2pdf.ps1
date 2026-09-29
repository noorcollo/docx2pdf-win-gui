<#
.SYNOPSIS
    Batch-convert .docx files to PDF using Microsoft Word (identical to File > Save As > PDF).

.DESCRIPTION
    Drives the locally installed Microsoft Word via COM automation and exports each
    document with Word's own PDF engine, so the output is exactly what you would get
    by saving manually. Recurses into subfolders by default.

.PARAMETER Path
    A folder containing .docx files, or a single .docx file.

.PARAMETER OutDir
    Optional. Destination folder for PDFs. The subfolder structure under Path is preserved.
    If omitted, each PDF is written next to its source .docx.

.PARAMETER NoRecurse
    Only process the top-level folder; skip subfolders.

.EXAMPLE
    docx2pdf "C:\Docs"
    docx2pdf "C:\Docs" -OutDir "D:\PDF"
    docx2pdf "C:\Docs" -NoRecurse
    docx2pdf "C:\Docs\report.docx"
#>

param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Path,

    [string]$OutDir,

    [switch]$NoRecurse
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $Path)) {
    Write-Host "Path not found: $Path" -ForegroundColor Red
    exit 1
}

$Path = (Resolve-Path $Path).Path

if (Test-Path $Path -PathType Leaf) {
    if ([System.IO.Path]::GetExtension($Path) -ne ".docx") {
        Write-Host "Not a .docx file: $Path" -ForegroundColor Red
        exit 1
    }
    $files = @(Get-Item $Path)
    $root  = Split-Path $Path
} else {
    $files = @(Get-ChildItem -Path $Path -Filter *.docx -Recurse:(-not $NoRecurse) -File |
               Where-Object { $_.Name -notlike "~$*" })   # skip Word lock files
    $root  = $Path
}

if ($files.Count -eq 0) {
    Write-Host "No .docx files found in: $Path" -ForegroundColor Yellow
    exit 0
}

Write-Host "Found $($files.Count) file(s). Converting..." -ForegroundColor Cyan

try {
    $word = New-Object -ComObject Word.Application
} catch {
    Write-Host "Microsoft Word is not installed or could not be started." -ForegroundColor Red
    exit 1
}
$word.Visible = $false
$word.DisplayAlerts = 0   # wdAlertsNone

$ok = 0; $fail = 0

try {
    foreach ($f in $files) {
        if ($OutDir) {
            $rel = $f.FullName.Substring($root.Length).TrimStart('\')
            $pdf = Join-Path $OutDir ([System.IO.Path]::ChangeExtension($rel, ".pdf"))
            New-Item -ItemType Directory -Force -Path (Split-Path $pdf) | Out-Null
        } else {
            $pdf = [System.IO.Path]::ChangeExtension($f.FullName, ".pdf")
        }

        try {
            $doc = $word.Documents.Open($f.FullName, $false, $true)   # ConfirmConversions, ReadOnly
            $doc.ExportAsFixedFormat($pdf, 17)                        # 17 = wdExportFormatPDF
            $doc.Close($false)
            Write-Host "  OK    $($f.FullName)" -ForegroundColor Green
            $ok++
        } catch {
            Write-Host "  FAIL  $($f.FullName)" -ForegroundColor Red
            Write-Host "        $($_.Exception.Message)" -ForegroundColor DarkRed
            $fail++
        }
    }
} finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}

Write-Host "Done: $ok succeeded, $fail failed." -ForegroundColor Cyan
if ($fail -gt 0) { exit 1 }
