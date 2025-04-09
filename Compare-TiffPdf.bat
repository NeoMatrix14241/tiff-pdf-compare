@echo off
title TIFF and PDF Comparison Tool
color 0A

:: Check if PowerShell is available
where powershell >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo PowerShell is not installed or not available in PATH
    echo Please install PowerShell to use this tool
    pause
    exit /b 1
)

:: Unblock files
powershell -ExecutionPolicy Bypass -Command & '%~dp0unblocker.ps1'

:: Launch PowerShell script with execution policy bypass
powershell.exe -ExecutionPolicy Bypass -File "%~dp0Compare-TiffPdfPages.ps1"

exit /b 0
