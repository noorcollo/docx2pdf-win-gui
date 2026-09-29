@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0docx2pdf.ps1" %*
