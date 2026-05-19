@echo off
REM Lance la GUI de configuration. La GUI elle-meme appelle Install-GoogleDriveSync.ps1
REM qui s'auto-eleve via UAC si necessaire.

setlocal
set "SCRIPT=%~dp0Show-GoogleDriveSyncGUI.ps1"

if not exist "%SCRIPT%" (
    echo [ERREUR] GUI introuvable : %SCRIPT%
    pause
    exit /b 1
)

start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%SCRIPT%"

endlocal
