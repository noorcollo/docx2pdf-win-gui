# DOCX to PDF for Windows

A lightweight Windows desktop app for batch-converting `.docx` files to PDF with **Microsoft Word's own PDF engine**. This preserves Word's layout and rendering rather than re-creating documents with a different converter.

> **Requires desktop Microsoft Word.** This app automates Word locally; it is not a standalone DOCX renderer.

## Download

**[Download the latest Windows x64 release](https://github.com/noorcollo/docx2pdf-win-gui/releases/latest/download/Docx2Pdf-win-x64.zip)** from [GitHub Releases](https://github.com/noorcollo/docx2pdf-win-gui/releases).

1. Download and extract the ZIP.
2. Run `Docx2Pdf.exe`.
3. Add Word files or folders, choose an output folder if desired, and select **Convert to PDF**.

The app is self-contained: it does not require a separate .NET installation. The download is unsigned, so Windows may show a SmartScreen notice. Verify that you downloaded it from this repository's Releases page; do not disable Windows security protections to run software you do not trust.

## Requirements

- Windows 10 or 11, 64-bit
- Desktop Microsoft Word 2013 or later, installed and activated

Word is controlled through its Windows COM automation interface. Word Online, WordPad, LibreOffice, and Office in a browser are not substitutes for the required desktop Word installation.

## Features

- Add multiple `.docx` files, add a folder, or drag files and folders into the window
- Scan subfolders recursively, or limit a folder scan to its top level
- Save each PDF beside its source, or choose a separate destination folder
- Preserve folder structure beneath the selected destination
- View per-file status and batch progress
- Continue past a failed document and review the reported error
- Cancel safely after the currently converting document finishes
- Confirm before replacing existing PDFs
- Skip Word's temporary lock files (`~$...`)

## Output behavior

- **No output folder selected:** each PDF is saved beside its source DOCX.
- **Output folder selected:** files imported from a folder are placed beneath a folder with that source folder's name, preserving subfolders. Individually selected files are placed directly in the chosen output folder.
- If a target PDF already exists, the app asks before replacing any existing output files.

## Privacy

Documents are opened and exported on your own PC by the locally installed Microsoft Word. The app does not upload your documents to a server.

## Troubleshooting

- **Microsoft Word was not found or could not be started:** install and activate the desktop version of Word, then try again.
- **A document fails:** open it in Word manually. Password protection, file damage, permissions, or another process holding the file can prevent conversion.
- **The app displays a startup error:** diagnostic details are written to `%LOCALAPPDATA%\Docx2Pdf\startup.log`.
- **The app appears not to open:** check Windows Security's Protection history and any SmartScreen prompt. Only run a copy obtained from this repository's Releases page. If a startup log was created, include it when reporting the issue.
- **Word remains open after an interrupted run:** close the app normally when possible. If Word remains hidden, close the extra `WINWORD.EXE` process from Task Manager after saving any other Word work.

## Build from source

### Requirements

- .NET 8 SDK
- Windows 10/11 x64 for running the app
- Microsoft Word 2013 or later for conversion

From PowerShell in the repository folder:

```powershell
.\build-windows.ps1
```

The self-contained executable is written to `dist\Docx2Pdf.exe`. The build can also be run manually:

```powershell
dotnet publish .\Docx2PdfGui.csproj -c Release -r win-x64 --self-contained true `
  -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true `
  -p:EnableCompressionInSingleFile=true -o .\dist
```

To prepare a release ZIP and SHA-256 checksum after building:

```powershell
.\package-release.ps1 -Version 1.0.0
```

The package is written to `release-assets\Docx2Pdf-win-x64.zip` with a `.sha256` file beside it. Compare a downloaded package's hash in PowerShell with:

```powershell
Get-FileHash .\Docx2Pdf-win-x64.zip -Algorithm SHA256
```

## Original command-line converter

This repository also retains the upstream PowerShell batch tool for scripted use:

```powershell
.\docx2pdf.cmd "C:\Docs"
.\docx2pdf.cmd "C:\Docs" -OutDir "D:\PDF"
.\docx2pdf.cmd "C:\Docs\report.docx"
```

It requires Windows PowerShell and desktop Microsoft Word. Unless `-OutDir` is provided, PDFs are saved next to their source files. See the script's help comments for `-NoRecurse` and other details.

## Attribution and license

The original Word COM batch-conversion implementation is from [ncdai/docx2pdf-win](https://github.com/ncdai/docx2pdf-win) and is distributed under the MIT License. This repository retains the upstream license and command-line scripts and adds a Windows Forms GUI. See [LICENSE](LICENSE).
