@echo off
chcp 65001 >nul
title ABBYY Auto OCR

echo =============================================
echo   ABBYY FINEREADER 16 AUTO OCR
echo =============================================
echo.

REM Chay PowerShell script o cung thu muc
echo Dang chay script...
echo.

powershell.exe -ExecutionPolicy Bypass -NoProfile -NoExit -File "%~dp0ABBYY-AutoOCR.ps1"

REM Neu PowerShell dong, van giu cua so bat
echo.
echo =============================================
pause