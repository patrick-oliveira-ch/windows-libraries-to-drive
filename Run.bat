@echo off
REM Lance Install-GoogleDriveSync.ps1. Le script PowerShell s'auto-élève via UAC.
REM Tous les arguments du .bat sont forwardés au script.

setlocal
set "SCRIPT=%~dp0Install-GoogleDriveSync.ps1"

if not exist "%SCRIPT%" (
    echo [ERREUR] Script introuvable : %SCRIPT%
    pause
    exit /b 1
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" %*

endlocal
