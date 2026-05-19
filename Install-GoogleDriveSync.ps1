<#
.SYNOPSIS
    Redirige les dossiers utilisateur Windows (Documents, Images, Vidéos, Musique, Objets 3D, Bureau)
    vers une arborescence Google Drive pour synchroniser automatiquement entre plusieurs PC.

.DESCRIPTION
    1. Vérifie les droits administrateur.
    2. Installe Google Drive for Desktop si absent (téléchargement + vérif signature Authenticode).
    3. Attend que Google Drive soit connecté et monté.
    4. Crée l'arborescence cible sur Google Drive (ex: G:\Mon Drive\WindowsLibraries\...).
    5. Pour chaque Known Folder : migre le contenu via robocopy, met à jour le registre,
       puis appelle SHSetKnownFolderPath et notifie l'Explorateur.
    6. Pour chaque mapping "extra" (Scripts, .ssh, .gitconfig, Templates Office, Signatures
       Outlook) : migre et crée un symlink.
    7. Optionnellement : désinstalle OneDrive et bloque sa réinstallation.

    Idempotent. Crée un journal UTF-8 dans %TEMP% (ou %USERPROFILE% en fallback).

    Impact système notable :
      - Modifie HKCU\...\User Shell Folders et HKCU\...\Shell Folders.
      - Si -DisableOneDrive : modifie HKLM\SOFTWARE\Policies\Microsoft\Windows\OneDrive
        (DisableFileSyncNGSC=1), ce qui désactive aussi OneDrive for Business.
        Voir -RestoreOneDrive pour annuler.

.PARAMETER RootName
    Nom du dossier racine créé sous "Mon Drive". Défaut : WindowsLibraries.
    Validé : alphanumériques, tirets, underscores, espaces simples (max 64 car.).

.PARAMETER IncludeDesktop
    Inclure le Bureau dans la synchronisation.

.PARAMETER Force3DObjects
    Force la création du dossier "3D Objects" même s'il n'existe pas (cas Windows 11 22H2+).
    Utile si tu utilises Blender, Paint 3D, etc.

.PARAMETER DriveLetter
    Lettre du drive Google Drive si auto-détection impossible (ex: 'G').

.PARAMETER SkipInstall
    Ne pas tenter d'installer Google Drive for Desktop.

.PARAMETER DisableOneDrive
    Après migration, désinstalle OneDrive et bloque sa réinstallation via policy.

.PARAMETER RestoreOneDrive
    Annule les policies HKLM appliquées par -DisableOneDrive. Ne réinstalle pas OneDrive.

.PARAMETER RefreshQuickAccess
    Épingle à l'Accès rapide Windows tous les dossiers présents sous WindowsLibraries\
    qui ne correspondent pas à un Known Folder déjà visible. Utile après avoir créé
    de nouveaux dossiers sur Drive depuis un autre PC. Ne nécessite pas les droits admin.

.PARAMETER InstallScheduledRefresh
    Installe une tâche planifiée Windows qui lance -RefreshQuickAccess automatiquement
    toutes les N minutes (voir -IntervalMinutes). Démarre à la connexion utilisateur.
    Tâche nommée 'GoogleDriveSync-RefreshQuickAccess', tourne en arrière-plan caché.

.PARAMETER IntervalMinutes
    Intervalle (en minutes) entre deux exécutions du auto-refresh. Défaut : 15. Range : 5-1440.

.PARAMETER UninstallScheduledRefresh
    Supprime la tâche planifiée installée par -InstallScheduledRefresh.

.PARAMETER IncludeScripts
    Synchronise %USERPROFILE%\Scripts\ via symlink.

.PARAMETER IncludeDevConfig
    Synchronise %USERPROFILE%\.ssh\ et %USERPROFILE%\.gitconfig via symlinks.
    BLOQUE l'opération si une clé SSH privée non chiffrée est détectée (sauf -Force).

.PARAMETER IncludeOfficeTemplates
    Synchronise %APPDATA%\Microsoft\Templates\ et %APPDATA%\Microsoft\Signatures\.

.PARAMETER MountTimeoutSeconds
    Délai max d'attente du montage Google Drive après installation. Défaut : 600.

.PARAMETER Force
    Supprime les confirmations interactives ET autorise les opérations risquées
    (ex: SSH non chiffrée, désinstall OneDrive avec fichiers cloud-only).

.EXAMPLE
    .\Install-GoogleDriveSync.ps1
    Installation standard : Documents, Images, Vidéos, Musique, Objets 3D.

.EXAMPLE
    .\Install-GoogleDriveSync.ps1 -IncludeDesktop -IncludeScripts -IncludeOfficeTemplates
    Inclut Bureau, Scripts et templates/signatures Office.

.EXAMPLE
    .\Install-GoogleDriveSync.ps1 -DisableOneDrive -Force
    Migre tout vers Google Drive ET désinstalle OneDrive sans demander confirmation.

.EXAMPLE
    .\Install-GoogleDriveSync.ps1 -RestoreOneDrive
    Annule uniquement les policies de blocage OneDrive (n'affecte pas les redirections).

.NOTES
    Licence : MIT.
    Nécessite : Windows 10/11, PowerShell 5.1+, droits admin, compte Google.
#>

[CmdletBinding(DefaultParameterSetName='Install')]
param(
    [Parameter(ParameterSetName='Install')]
    [ValidateScript({
        if ($_ -notmatch '^[A-Za-z0-9_\-]([A-Za-z0-9 _\-]{0,62}[A-Za-z0-9_\-])?$') {
            throw "RootName invalide : 1-64 caractères alphanum/espace/-/_, commence et finit par alphanum/-/_."
        }
        if ($_ -match '^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$') {
            throw "RootName ne peut pas être un nom Windows réservé ($_)."
        }
        $true
    })]
    [string]$RootName = 'WindowsLibraries',

    [Parameter(ParameterSetName='Install')]
    [switch]$IncludeDesktop,

    [Parameter(ParameterSetName='Install')]
    [switch]$Force3DObjects,

    [Parameter(ParameterSetName='Install')]
    [ValidatePattern('^[A-Za-z]$')]
    [string]$DriveLetter,

    [Parameter(ParameterSetName='Install')][switch]$SkipInstall,
    [Parameter(ParameterSetName='Install')][switch]$DisableOneDrive,
    [Parameter(ParameterSetName='Install')][switch]$IncludeScripts,
    [Parameter(ParameterSetName='Install')][switch]$IncludeDevConfig,
    [Parameter(ParameterSetName='Install')][switch]$IncludeOfficeTemplates,

    [Parameter(ParameterSetName='Install')]
    [ValidateRange(60, 3600)]
    [int]$MountTimeoutSeconds = 600,

    [Parameter(ParameterSetName='Restore', Mandatory=$true)]
    [switch]$RestoreOneDrive,

    [Parameter(ParameterSetName='Refresh', Mandatory=$true)]
    [switch]$RefreshQuickAccess,

    [Parameter(ParameterSetName='ScheduleInstall', Mandatory=$true)]
    [switch]$InstallScheduledRefresh,

    [Parameter(ParameterSetName='ScheduleInstall')]
    [ValidateRange(5, 1440)]
    [int]$IntervalMinutes = 15,

    [Parameter(ParameterSetName='ScheduleUninstall', Mandatory=$true)]
    [switch]$UninstallScheduledRefresh,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'

# Auto-élévation : relance avec UAC si pas admin (préserve les arguments)
# Le mode -RefreshQuickAccess ne touche que HKCU → pas besoin d'admin.
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal(
    [Security.Principal.WindowsIdentity]::GetCurrent())
$needsAdmin = -not $RefreshQuickAccess
if ($needsAdmin -and -not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    if (-not $PSCommandPath) { throw "Impossible de relancer avec UAC : script appelé sans fichier." }
    Write-Host "Élévation administrateur requise pour : $PSCommandPath" -ForegroundColor Yellow
    if (-not $Force) {
        $reply = Read-Host "Confirmer le lancement en admin de CE script ? (o/N)"
        if ($reply -notmatch '^[oOyY]') { Write-Host "Annulé." -ForegroundColor Red; exit 1 }
    }
    # Quote chaque argument pour préserver espaces / caractères spéciaux à travers Start-Process.
    function Quote-Arg { param([string]$Value) '"' + ($Value -replace '"','""') + '"' }
    $relaunch = @('-NoProfile','-ExecutionPolicy','Bypass','-File',(Quote-Arg $PSCommandPath))
    foreach ($k in $PSBoundParameters.Keys) {
        $v = $PSBoundParameters[$k]
        if ($v -is [switch]) {
            if ($v.IsPresent) { $relaunch += "-$k" }
        } else {
            $relaunch += "-$k"
            $relaunch += (Quote-Arg "$v")
        }
    }
    try {
        Start-Process -FilePath powershell.exe -Verb RunAs -ArgumentList $relaunch -ErrorAction Stop
    } catch {
        Write-Host "Élévation refusée ou échouée : $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
    exit 0
}

# TLS 1.2/1.3 (PS 5.1 défaut peut être TLS 1.0)
try {
    [Net.ServicePointManager]::SecurityProtocol =
        [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13
} catch {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
}

# Log file (fallback si %TEMP% indisponible)
$logDir = $env:TEMP
if (-not (Test-Path -LiteralPath $logDir)) { $logDir = $env:USERPROFILE }
$script:LogFile = Join-Path $logDir "GoogleDriveSync_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# Purge anciens logs (>30j)
Get-ChildItem -LiteralPath $logDir -Filter 'GoogleDriveSync_*.log' -ErrorAction SilentlyContinue |
    Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) } |
    Remove-Item -Force -ErrorAction SilentlyContinue

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR','OK')][string]$Level = 'INFO'
    )
    $stamp = Get-Date -Format 'HH:mm:ss'
    $line  = "[$stamp][$Level] $Message"
    $color = switch ($Level) { 'WARN' {'Yellow'} 'ERROR' {'Red'} 'OK' {'Green'} default {'Cyan'} }
    Write-Host $line -ForegroundColor $color
    try { Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8 } catch {}
}

function Assert-Admin {
    # Filet de sécurité ; l'auto-élévation au top a normalement traité le cas.
    if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw "Droits administrateur requis."
    }
}

function New-AdminOnlyTempDir {
    # Crée un sous-dossier avec ACL admins-only pour le download de l'installeur.
    # Préfère %SystemRoot%\Temp (déjà admin-only) à %TEMP% (writable user) si possible.
    $base = if (Test-Path -LiteralPath "$env:SystemRoot\Temp") { "$env:SystemRoot\Temp" } else { $env:TEMP }
    $name = [System.IO.Path]::GetRandomFileName()
    $dir  = Join-Path $base "gdsync_$name"
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
    try {
        $acl = New-Object System.Security.AccessControl.DirectorySecurity
        $acl.SetAccessRuleProtection($true, $false)
        foreach ($who in @(
            (New-Object System.Security.Principal.NTAccount('BUILTIN\Administrators')),
            (New-Object System.Security.Principal.NTAccount('NT AUTHORITY\SYSTEM'))
        )) {
            $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule(
                $who, 'FullControl',
                'ContainerInherit,ObjectInherit', 'None', 'Allow')))
        }
        Set-Acl -LiteralPath $dir -AclObject $acl
    } catch {
        # Fail-closed : si on ne peut pas durcir l'ACL, on n'utilise pas le dossier.
        Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue
        throw "Impossible de durcir l'ACL sur $dir : $($_.Exception.Message)"
    }
    return $dir
}

# --- Helpers chemins / liens ---

function ConvertTo-NormalPath {
    param([string]$Path)
    if (-not $Path) { return $Path }
    # Refuse les UNC : on ne migre/symlink que vers du local ou drive monté.
    if ($Path -match '^(\\\\|//)') { throw "Chemin UNC refusé : $Path" }
    try { return [System.IO.Path]::GetFullPath($Path).TrimEnd('\') } catch { return $Path.TrimEnd('\') }
}

function Remove-LinkSafely {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return }
    $item = Get-Item -LiteralPath $Path -Force
    if ($item.LinkType -in @('SymbolicLink','Junction')) {
        # NE PAS utiliser Remove-Item -Recurse : suit la cible sur PS 5.1.
        if ($item.PSIsContainer) {
            [System.IO.Directory]::Delete($Path, $false)
        } else {
            [System.IO.File]::Delete($Path)
        }
        return
    }
    if ($item.PSIsContainer) {
        Remove-Item -LiteralPath $Path -Recurse -Force
    } else {
        Remove-Item -LiteralPath $Path -Force
    }
}

function Test-OneDriveSourcePath {
    param([string]$Path)
    if (-not $Path) { return $false }
    return ($Path -match '\\OneDrive(?:\b|\\| - )')
}

# --- Détection / installation Google Drive ---

function Test-GoogleDriveInstalled {
    $candidates = @(
        "${env:ProgramFiles}\Google\Drive File Stream\launch.bat",
        "${env:ProgramFiles}\Google\Drive File Stream\GoogleDriveFS.exe",
        "${env:ProgramFiles(x86)}\Google\Drive File Stream\GoogleDriveFS.exe",
        "${env:ProgramFiles}\Google\DriveFS\GoogleDriveFS.exe"
    )
    foreach ($p in $candidates) { if (Test-Path -LiteralPath $p) { return $true } }
    if (Get-Process -Name 'GoogleDriveFS' -ErrorAction SilentlyContinue) { return $true }
    if (Get-ItemProperty 'HKLM:\SOFTWARE\Google\DriveFS' -ErrorAction SilentlyContinue) { return $true }
    return $false
}

function Get-DriveLauncher {
    foreach ($p in @(
        "${env:ProgramFiles}\Google\Drive File Stream\launch.bat",
        "${env:ProgramFiles}\Google\DriveFS\launch.bat"
    )) { if (Test-Path -LiteralPath $p) { return $p } }
    return $null
}

function Install-GoogleDrive {
    if (Test-GoogleDriveInstalled) {
        Write-Log "Google Drive for Desktop déjà installé." 'OK'
        return
    }
    Write-Log "Téléchargement de Google Drive for Desktop..."
    $tempDir = New-AdminOnlyTempDir
    $installer = Join-Path $tempDir 'GoogleDriveSetup.exe'
    $url = 'https://dl.google.com/drive-file-stream/GoogleDriveSetup.exe'
    try {
        Invoke-WebRequest -Uri $url -OutFile $installer -UseBasicParsing
    } catch {
        throw "Échec du téléchargement depuis $url : $($_.Exception.Message)"
    }

    # Vérification signature Authenticode
    Write-Log "Vérification de la signature Authenticode..."
    $sig = Get-AuthenticodeSignature -FilePath $installer
    if ($sig.Status -ne 'Valid') {
        Remove-Item -LiteralPath $installer -Force -ErrorAction SilentlyContinue
        throw "Signature Authenticode invalide ($($sig.Status)). Installeur supprimé."
    }
    # Match ancré : évite CN trompeur du type "NotGoogle LLC"
    if ($sig.SignerCertificate.Subject -notmatch '(?:^|, )(?:CN|O)=Google LLC(?:,|$)') {
        Remove-Item -LiteralPath $installer -Force -ErrorAction SilentlyContinue
        throw "Signataire inattendu : $($sig.SignerCertificate.Subject)."
    }
    Write-Log "Signature valide : $($sig.SignerCertificate.Subject)" 'OK'

    Write-Log "Installation silencieuse..."
    # Google Drive for Desktop accepte --silent (alias officiel) ; rétrocompat /S également supporté.
    $proc = Start-Process -FilePath $installer `
        -ArgumentList '--silent','--desktop_shortcut','--gsuite_shortcuts=false' `
        -Wait -PassThru
    if ($proc.ExitCode -ne 0) {
        Write-Log "Installeur exit=$($proc.ExitCode) — vérifier manuellement." 'WARN'
    }
    try { Remove-Item -LiteralPath $tempDir -Recurse -Force } catch {
        Write-Log "Impossible de nettoyer $tempDir : $($_.Exception.Message)" 'WARN'
    }
    Write-Log "Google Drive installé. Connecte-toi à ton compte Google maintenant." 'WARN'
    $launcher = Get-DriveLauncher
    if ($launcher) { Start-Process -FilePath $launcher -ErrorAction SilentlyContinue }
}

function Find-GoogleDriveMount {
    param([int]$TimeoutSeconds = 600)
    if ($DriveLetter) {
        foreach ($name in @('Mon Drive','My Drive')) {
            $explicit = "${DriveLetter}:\$name"
            if (Test-Path -LiteralPath $explicit) { return $explicit }
        }
    }
    Write-Log "Attente du montage Google Drive (max $TimeoutSeconds s)..."
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        foreach ($d in (Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue)) {
            foreach ($name in @('Mon Drive','My Drive')) {
                $candidate = Join-Path $d.Root $name
                if (Test-Path -LiteralPath $candidate) {
                    Write-Log "Google Drive détecté." 'OK'
                    return $candidate
                }
            }
        }
        Start-Sleep -Seconds 5
    }
    throw "Google Drive non monté dans le délai imparti. Connecte-toi puis relance."
}

# --- Tests sécurité préalables ---

function Test-SshKeyEncrypted {
    param([Parameter(Mandatory)][string]$KeyPath)
    if (-not (Test-Path -LiteralPath $KeyPath -PathType Leaf)) { return $true }
    # Garde-fou DoS : clé > 1 Mo = anomalie, on ne charge pas.
    if ((Get-Item -LiteralPath $KeyPath -Force).Length -gt 1MB) {
        Write-Log "Fichier > 1 Mo ignoré : $KeyPath" 'WARN'
        return $true
    }
    try { $content = Get-Content -LiteralPath $KeyPath -Raw -ErrorAction Stop } catch { return $true }
    if ($content -match 'Proc-Type:\s*4,ENCRYPTED') { return $true }
    if ($content -match '-----BEGIN OPENSSH PRIVATE KEY-----') {
        if (Get-Command ssh-keygen -ErrorAction SilentlyContinue) {
            $out = & ssh-keygen -y -P '' -f $KeyPath 2>$null
            # Non-chiffrée si exit=0 ET output ressemble à une clé publique SSH.
            if ($LASTEXITCODE -eq 0 -and ($out -join "`n") -match '^(ssh-rsa|ssh-ed25519|ecdsa-sha2|ssh-dss) ') {
                return $false
            }
            return $true
        }
        return $null
    }
    return $true
}

function Assert-SshKeysSafe {
    $sshDir = "$env:USERPROFILE\.ssh"
    if (-not (Test-Path -LiteralPath $sshDir)) { return }
    $candidates = Get-ChildItem -LiteralPath $sshDir -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^id_(rsa|ed25519|ecdsa|dsa)$' -or $_.Name -like '*.pem' }
    $unprotected = @()
    $unknown     = @()
    foreach ($k in $candidates) {
        $r = Test-SshKeyEncrypted -KeyPath $k.FullName
        if ($r -eq $false) { $unprotected += $k.Name }
        elseif ($null -eq $r) { $unknown += $k.Name }
    }
    if ($unprotected.Count -gt 0) {
        $msg = "Clé(s) SSH SANS passphrase détectée(s) : $($unprotected -join ', '). " +
               "Les synchroniser vers Google Drive expose tes accès SSH."
        if (-not $Force) {
            throw "$msg`nProtège-les avec : ssh-keygen -p -f <chemin> -- puis relance, OU utilise -Force."
        }
        Write-Log "$msg (-Force actif → on continue)" 'WARN'
    }
    if ($unknown.Count -gt 0) {
        Write-Log "Impossible de vérifier le chiffrement de : $($unknown -join ', ') (ssh-keygen absent)." 'WARN'
    }
}

# --- Conflit OneDrive ---

function Test-OneDriveConflict {
    if ($DisableOneDrive) {
        Write-Log "OneDrive sera désinstallé après migration." 'INFO'
        return
    }
    $kfm = Get-ItemProperty -Path 'HKCU:\Software\Microsoft\OneDrive\Accounts\Personal' `
                            -Name 'UserFolder' -ErrorAction SilentlyContinue
    if (-not $kfm) { return }
    $oneDrivePath = $kfm.UserFolder
    $userShell = Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' `
                                  -ErrorAction SilentlyContinue
    $conflicts = @()
    foreach ($prop in $userShell.PSObject.Properties) {
        if ($prop.Value -is [string] -and $prop.Value -like "$oneDrivePath*") {
            $conflicts += $prop.Name
        }
    }
    if ($conflicts.Count -gt 0) {
        Write-Log "OneDrive synchronise déjà : $($conflicts -join ', ')" 'WARN'
        Write-Log "Désactive Known Folder Move (Param. OneDrive > Sauvegarde) ou utilise -DisableOneDrive." 'WARN'
        if (-not $Force) {
            $reply = Read-Host "Continuer malgré tout ? (o/N)"
            if ($reply -notmatch '^[oOyY]') { throw "Annulé par l'utilisateur." }
        }
    }
}

# --- API Shell ---

function Initialize-ShellApi {
    if ('KnownFolderPath' -as [type]) { return }
    Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class KnownFolderPath {
    [DllImport("shell32.dll", CharSet = CharSet.Auto)]
    public static extern int SHSetKnownFolderPath(
        ref Guid rfid,
        uint dwFlags,
        IntPtr hToken,
        [MarshalAs(UnmanagedType.LPWStr)] string pszPath);

    [DllImport("shell32.dll", CharSet = CharSet.Auto)]
    public static extern void SHChangeNotify(
        int wEventId,
        uint uFlags,
        IntPtr dwItem1,
        IntPtr dwItem2);
}
'@
}

function Invoke-ExplorerRefresh {
    param([string[]]$UpdatedPaths)
    Initialize-ShellApi
    # SHCNE_UPDATEDIR = 0x00001000, SHCNF_PATHW = 0x0005
    foreach ($p in $UpdatedPaths) {
        $ptr = [System.Runtime.InteropServices.Marshal]::StringToHGlobalUni($p)
        try {
            [KnownFolderPath]::SHChangeNotify(0x00001000, 0x0005, $ptr, [IntPtr]::Zero)
        } finally {
            [System.Runtime.InteropServices.Marshal]::FreeHGlobal($ptr)
        }
    }
    # Fallback global pour recharger les associations
    [KnownFolderPath]::SHChangeNotify(0x08000000, 0x0000, [IntPtr]::Zero, [IntPtr]::Zero)
}

# --- Accès rapide Windows (Quick Access) ---

function Test-IsPinnedToQuickAccess {
    param([Parameter(Mandatory)][string]$Path)
    try {
        $shell = New-Object -ComObject Shell.Application
        # Namespace de l'Accès rapide
        $qa = $shell.Namespace('shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}')
        if (-not $qa) { return $false }
        $target = (ConvertTo-NormalPath $Path).ToLowerInvariant()
        foreach ($item in $qa.Items()) {
            try {
                $itemPath = (ConvertTo-NormalPath $item.Path).ToLowerInvariant()
                if ($itemPath -eq $target) { return $true }
            } catch {}
        }
        return $false
    } catch {
        return $false
    }
}

function Add-ToQuickAccess {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) { return }
    if (Test-IsPinnedToQuickAccess -Path $Path) {
        Write-Log "Déjà épinglé : $(Split-Path -Leaf $Path)"
        return
    }
    try {
        $shell = New-Object -ComObject Shell.Application
        $folder = $shell.Namespace($Path)
        if ($folder -and $folder.Self) {
            $folder.Self.InvokeVerb('pintohome')
            Write-Log "Épinglé à l'Accès rapide : $(Split-Path -Leaf $Path)" 'OK'
        }
    } catch {
        Write-Log "Échec épinglage $Path : $($_.Exception.Message)" 'WARN'
    }
}

function Update-QuickAccessPins {
    param([Parameter(Mandatory)][string]$RootPath)
    if (-not (Test-Path -LiteralPath $RootPath)) {
        Write-Log "Root introuvable, skip Accès rapide : $RootPath" 'WARN'
        return
    }
    Write-Log "=== Épinglage Accès rapide ===" 'INFO'
    # Ces Known Folders sont déjà épinglés par défaut dans l'Accès rapide Win10/11.
    # 3D Objects NE l'est PAS (retiré du défaut depuis Win11 22H2) → on l'épingle.
    $skipNames = @('Documents','Pictures','Videos','Music','Desktop')
    foreach ($d in Get-ChildItem -LiteralPath $RootPath -Directory -ErrorAction SilentlyContinue) {
        if ($skipNames -contains $d.Name) {
            Write-Log "Skip $($d.Name) (Known Folder, déjà visible)"
            continue
        }
        Add-ToQuickAccess -Path $d.FullName
    }
}

function Get-DriveRootPath {
    # Détecte le root à partir du Known Folder Documents (déjà redirigé).
    $docPath = Get-CurrentKnownFolderPath -RegName 'Personal'
    if (-not $docPath -or -not (Test-Path -LiteralPath $docPath)) { return $null }
    return Split-Path -Parent $docPath
}

# --- Tâche planifiée auto-refresh ---

$script:ScheduledTaskName = 'GoogleDriveSync-RefreshQuickAccess'

function Test-ScheduledRefreshInstalled {
    return $null -ne (Get-ScheduledTask -TaskName $script:ScheduledTaskName -ErrorAction SilentlyContinue)
}

function Register-DriveSyncRefreshTask {
    param([Parameter(Mandatory)][int]$IntervalMinutes)

    if (-not $PSCommandPath) { throw "Chemin du script inconnu — relance via -File <chemin>." }
    $scriptPath = $PSCommandPath
    Write-Log "Installation tâche planifiée : $script:ScheduledTaskName"
    Write-Log "  Script  : $scriptPath"
    Write-Log "  Cadence : toutes les $IntervalMinutes min, dès la connexion"

    $action = New-ScheduledTaskAction `
        -Execute 'powershell.exe' `
        -Argument ("-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`" -RefreshQuickAccess")

    # Deux triggers séparés (assigner Repetition entre triggers ne marche pas en PS 5.1) :
    #   1. AtLogOn : démarre dès la connexion utilisateur
    #   2. TimeTrigger Once + Repetition : répète toutes les N min (durée 9999 jours = quasi indéfini)
    $startTime = (Get-Date).AddMinutes(1)
    $triggers = @(
        (New-ScheduledTaskTrigger -AtLogOn -User "$env:USERDOMAIN\$env:USERNAME"),
        (New-ScheduledTaskTrigger -Once -At $startTime `
            -RepetitionInterval (New-TimeSpan -Minutes $IntervalMinutes) `
            -RepetitionDuration (New-TimeSpan -Days 9999))
    )

    $principal = New-ScheduledTaskPrincipal `
        -UserId "$env:USERDOMAIN\$env:USERNAME" `
        -LogonType Interactive `
        -RunLevel Limited

    $settings = New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -StartWhenAvailable `
        -MultipleInstances IgnoreNew `
        -ExecutionTimeLimit (New-TimeSpan -Minutes 5) `
        -RestartCount 0

    Register-ScheduledTask `
        -TaskName $script:ScheduledTaskName `
        -Action $action -Trigger $triggers `
        -Principal $principal -Settings $settings `
        -Description "Auto-épingle les dossiers Drive sous WindowsLibraries à l'Accès rapide Explorer." `
        -Force | Out-Null

    Write-Log "Tâche planifiée enregistrée." 'OK'
    Write-Log "Prochain run dans ~1 min, puis toutes les $IntervalMinutes min."
}

function Unregister-DriveSyncRefreshTask {
    if (Test-ScheduledRefreshInstalled) {
        Unregister-ScheduledTask -TaskName $script:ScheduledTaskName -Confirm:$false
        Write-Log "Tâche planifiée '$script:ScheduledTaskName' supprimée." 'OK'
    } else {
        Write-Log "Pas de tâche '$script:ScheduledTaskName' à supprimer." 'INFO'
    }
}

# --- Known Folders ---

$KnownFolders = @{
    'Documents'  = @{ Guid = 'FDD39AD0-238F-46AF-ADB4-6C85480369C7'; Reg = 'Personal'; Local = "$env:USERPROFILE\Documents" }
    'Pictures'   = @{ Guid = '33E28130-4E1E-4676-835A-98395C3BC3BB'; Reg = 'My Pictures'; Local = "$env:USERPROFILE\Pictures" }
    'Videos'     = @{ Guid = '18989B1D-99B5-455B-841C-AB7C74E4DDFC'; Reg = 'My Video'; Local = "$env:USERPROFILE\Videos" }
    'Music'      = @{ Guid = '4BD8D571-6D19-48D3-BE97-422220080E43'; Reg = 'My Music'; Local = "$env:USERPROFILE\Music" }
    '3D Objects' = @{ Guid = '31C0DD25-9439-4F12-BF41-7FF4EDA38722'; Reg = '{31C0DD25-9439-4F12-BF41-7FF4EDA38722}'; Local = "$env:USERPROFILE\3D Objects" }
    'Desktop'    = @{ Guid = 'B4BFCC3A-DB2C-424C-B029-7FE99A87C641'; Reg = 'Desktop'; Local = "$env:USERPROFILE\Desktop" }
}

$ExtraMappings = @{
    'Scripts'         = @{ Source = "$env:USERPROFILE\Scripts";          Type = 'Dir';  Subdir = 'Scripts' }
    'SSH'             = @{ Source = "$env:USERPROFILE\.ssh";             Type = 'Dir';  Subdir = 'DevConfig\.ssh' }
    'GitConfig'       = @{ Source = "$env:USERPROFILE\.gitconfig";       Type = 'File'; Subdir = 'DevConfig\.gitconfig' }
    'OfficeTemplates' = @{ Source = "$env:APPDATA\Microsoft\Templates";  Type = 'Dir';  Subdir = 'Office\Templates' }
    'OutlookSig'      = @{ Source = "$env:APPDATA\Microsoft\Signatures"; Type = 'Dir';  Subdir = 'Office\Signatures' }
}

function Get-CurrentKnownFolderPath {
    param([Parameter(Mandatory)][string]$RegName)
    $userShell = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'
    $val = (Get-ItemProperty -Path $userShell -Name $RegName -ErrorAction SilentlyContinue).$RegName
    if (-not $val) { return $null }
    return [System.Environment]::ExpandEnvironmentVariables($val)
}

function Set-KnownFolder {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$TargetPath
    )
    Initialize-ShellApi
    $kf = $KnownFolders[$Name]
    if (-not $kf) { throw "Known folder inconnu : $Name" }
    if (-not (Test-Path -LiteralPath $TargetPath)) {
        New-Item -ItemType Directory -Path $TargetPath -Force | Out-Null
    }
    $userShell = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'
    $shell     = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
    Set-ItemProperty -Path $userShell -Name $kf.Reg -Value $TargetPath -Type ExpandString
    Set-ItemProperty -Path $shell     -Name $kf.Reg -Value $TargetPath -Type String -ErrorAction SilentlyContinue

    $guid = [Guid]$kf.Guid
    $hr = [KnownFolderPath]::SHSetKnownFolderPath([ref]$guid, 0, [IntPtr]::Zero, $TargetPath)
    if ($hr -ne 0) {
        Write-Log "SHSetKnownFolderPath HRESULT 0x$($hr.ToString('X')) pour $Name (registre OK)." 'WARN'
    } else {
        Write-Log "Known Folder '$Name' redirigé." 'OK'
    }
}

# --- Migration de fichiers ---

function Invoke-Robocopy {
    param(
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][string]$Destination,
        # /XJ exclut les junctions (évite la récursion via My Pictures, My Music, etc. dans Documents)
        [string[]]$Options = @('/E','/R:5','/W:5','/XJ'),
        [string]$LogName = 'robocopy'
    )
    $rcLog = Join-Path $env:TEMP "${LogName}_$(Get-Date -Format 'HHmmss').log"
    $rcArgs = @($Source, $Destination) + $Options +
              @('/NFL','/NDL','/NJH','/NJS','/NP',"/LOG:$rcLog")
    & robocopy @rcArgs | Out-Null
    $code = $LASTEXITCODE
    Write-Log "robocopy exit=$code (log: $rcLog)"
    return $code
}

function Move-FolderContents {
    param([Parameter(Mandatory)][string]$Source, [Parameter(Mandatory)][string]$Destination)
    if (-not (Test-Path -LiteralPath $Source)) { return }
    try {
        $srcResolved = ConvertTo-NormalPath (Resolve-Path -LiteralPath $Source -ErrorAction Stop).Path
        $dstResolved = ConvertTo-NormalPath (Resolve-Path -LiteralPath $Destination -ErrorAction Stop).Path
        if ($srcResolved -eq $dstResolved) { return }
    } catch { return }

    if (Test-OneDriveSourcePath -Path $Source) {
        Write-Log "Source dans OneDrive : la migration va matérialiser les fichiers cloud-only (download)." 'WARN'
        if (-not $Force) {
            $reply = Read-Host "Continuer la migration depuis OneDrive ? (o/N)"
            if ($reply -notmatch '^[oOyY]') { throw "Migration interrompue." }
        }
    }

    $hasItems = [System.IO.Directory]::EnumerateFileSystemEntries($Source) | Select-Object -First 1
    if (-not $hasItems) {
        try { Remove-Item -LiteralPath $Source -Force -Recurse -ErrorAction SilentlyContinue } catch {}
        return
    }

    Write-Log "robocopy /MOVE : $Source -> $Destination"
    $code = Invoke-Robocopy -Source $Source -Destination $Destination `
            -Options @('/MOVE','/E','/R:5','/W:5','/XJ') -LogName "rc_$(Split-Path $Source -Leaf)"
    if ($code -ge 8) {
        throw "Robocopy a échoué (code $code) pour $Source. Voir log."
    }
    # /MOVE supprime déjà la source si tout est copié ; nettoyage final seulement si dir vide.
    if ((Test-Path -LiteralPath $Source) -and
        -not ([System.IO.Directory]::EnumerateFileSystemEntries($Source) | Select-Object -First 1)) {
        try { Remove-Item -LiteralPath $Source -Force -Recurse } catch {
            Write-Log "Source résiduelle non supprimée : $Source" 'WARN'
        }
    }
}

# --- Symlinks (mappings extras) ---

function New-DriveSymLink {
    param(
        [Parameter(Mandatory)][string]$LinkPath,
        [Parameter(Mandatory)][string]$TargetPath,
        [switch]$IsFile
    )
    if (Test-Path -LiteralPath $LinkPath) {
        $item = Get-Item -LiteralPath $LinkPath -Force
        if ($item.LinkType -in @('SymbolicLink','Junction')) {
            $current = @($item.Target) | Select-Object -First 1
            if ((ConvertTo-NormalPath $current) -eq (ConvertTo-NormalPath $TargetPath)) {
                Write-Log "Symlink déjà en place : $LinkPath" 'OK'
                return
            }
            Write-Log "Symlink existant pointe ailleurs — remplacement."
            Remove-LinkSafely -Path $LinkPath
        } else {
            throw "$LinkPath existe et n'est pas un lien — migrer son contenu avant relance."
        }
    }
    $parent = Split-Path -Parent $LinkPath
    if (-not (Test-Path -LiteralPath $parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
    New-Item -ItemType SymbolicLink -Path $LinkPath -Target $TargetPath | Out-Null
    Write-Log "Symlink créé : $(Split-Path $LinkPath -Leaf)" 'OK'
}

function Update-ExtraMapping {
    param([Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][string]$RootPath)
    $m = $ExtraMappings[$Name]
    if (-not $m) { throw "Mapping inconnu : $Name" }
    $source = $m.Source
    $target = Join-Path $RootPath $m.Subdir
    $isFile = ($m.Type -eq 'File')

    Write-Log "--- Extra : $Name ---"

    if (Test-Path -LiteralPath $source) {
        $srcItem = Get-Item -LiteralPath $source -Force
        $isLink  = $srcItem.LinkType -in @('SymbolicLink','Junction')
        if (-not $isLink) {
            if ($isFile) {
                $targetParent = Split-Path -Parent $target
                if (-not (Test-Path -LiteralPath $targetParent)) {
                    New-Item -ItemType Directory -Path $targetParent -Force | Out-Null
                }
                if (Test-Path -LiteralPath $target) {
                    $srcDate = (Get-Item -LiteralPath $source).LastWriteTime
                    $tgtDate = (Get-Item -LiteralPath $target).LastWriteTime
                    if ($srcDate -gt $tgtDate) {
                        Copy-Item -LiteralPath $source -Destination $target -Force
                        Write-Log "Source plus récente — cible écrasée."
                    } else {
                        Write-Log "Cible plus récente ou identique — conservée."
                    }
                } else {
                    Copy-Item -LiteralPath $source -Destination $target -Force
                }
                Remove-Item -LiteralPath $source -Force
            } else {
                if (-not (Test-Path -LiteralPath $target)) {
                    New-Item -ItemType Directory -Path $target -Force | Out-Null
                }
                Write-Log "Fusion contenu (robocopy /E /XO)..."
                $code = Invoke-Robocopy -Source $source -Destination $target `
                        -Options @('/E','/XO','/R:5','/W:5','/XJ') -LogName "extra_$Name"
                if ($code -ge 8) {
                    throw "Robocopy a échoué pour mapping extra '$Name' (code $code)."
                }
                # Suppression de la source SEULEMENT si tous fichiers présents sur la cible.
                $sourceFull = (ConvertTo-NormalPath $source)
                $hasMissing = $false
                try {
                    foreach ($f in [System.IO.Directory]::EnumerateFiles($source, '*', 'AllDirectories')) {
                        $fullF = (ConvertTo-NormalPath $f)
                        # StartsWith en mode case-insensitive (NTFS est case-insensitive par défaut)
                        if (-not $fullF.StartsWith($sourceFull, [System.StringComparison]::OrdinalIgnoreCase)) {
                            $hasMissing = $true; break
                        }
                        $rel = $fullF.Substring($sourceFull.Length).TrimStart([System.IO.Path]::DirectorySeparatorChar)
                        if (-not (Test-Path -LiteralPath (Join-Path $target $rel))) {
                            $hasMissing = $true; break
                        }
                    }
                } catch {
                    Write-Log "Énumération impossible ($($_.Exception.GetType().Name)) — source conservée par sécurité." 'WARN'
                    $hasMissing = $true
                }
                if ($hasMissing) {
                    Write-Log "Certains fichiers non vérifiés côté cible — source conservée." 'WARN'
                } else {
                    Remove-LinkSafely -Path $source
                }
            }
        }
    } else {
        if (-not $isFile -and -not (Test-Path -LiteralPath $target)) {
            New-Item -ItemType Directory -Path $target -Force | Out-Null
        }
    }

    if ($isFile -and -not (Test-Path -LiteralPath $target)) {
        Write-Log "Cible fichier absente — symlink différé." 'WARN'
        return
    }
    if (Test-Path -LiteralPath $source) {
        Write-Log "Source $source toujours présente (suppression échouée ?) — symlink non créé." 'WARN'
        return
    }
    New-DriveSymLink -LinkPath $source -TargetPath $target -IsFile:$isFile
}

# --- OneDrive ---

function Disable-OneDrive {
    Write-Log "=== Désinstallation OneDrive ===" 'INFO'

    # Garde-fou : applications dépendantes de OneDrive ouvertes
    $openOffice = Get-Process -ErrorAction SilentlyContinue -Name @(
        'WINWORD','EXCEL','POWERPNT','OUTLOOK','ONENOTE',
        'MSACCESS','MSPUB','VISIO','WINPROJ',
        'Teams','ms-teams','lync'
    )
    if ($openOffice) {
        $names = ($openOffice | Select-Object -Expand ProcessName -Unique) -join ', '
        Write-Log "Applications Office ouvertes : $names — ferme-les pour éviter toute corruption." 'WARN'
        if (-not $Force) {
            $reply = Read-Host "Continuer quand même ? (o/N)"
            if ($reply -notmatch '^[oOyY]') { Write-Log "Désinstall OneDrive annulée." 'WARN'; return }
        }
    }

    # Confirmation finale
    if (-not $Force) {
        $reply = Read-Host "Confirmer désinstallation COMPLÈTE de OneDrive ? (o/N)"
        if ($reply -notmatch '^[oOyY]') { Write-Log "Désinstall OneDrive annulée." 'WARN'; return }
    }

    Write-Log "Arrêt des processus OneDrive..."
    Get-Process -Name 'OneDrive','FileCoAuth' -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2

    Write-Log "Désactivation tâches planifiées OneDrive..."
    Get-ScheduledTask -TaskName 'OneDrive*' -ErrorAction SilentlyContinue |
        ForEach-Object {
            try { Disable-ScheduledTask -TaskName $_.TaskName -TaskPath $_.TaskPath -ErrorAction Stop | Out-Null } catch {}
        }

    $uninstaller = $null
    foreach ($p in @(
        "$env:SystemRoot\System32\OneDriveSetup.exe",
        "$env:SystemRoot\SysWOW64\OneDriveSetup.exe"
    )) { if (Test-Path -LiteralPath $p) { $uninstaller = $p; break } }
    if ($uninstaller) {
        Write-Log "Lancement $uninstaller /uninstall ..."
        Start-Process -FilePath $uninstaller -ArgumentList '/uninstall' -Wait
    } else {
        Write-Log "OneDriveSetup.exe introuvable." 'WARN'
    }

    Remove-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run' `
                        -Name 'OneDrive' -ErrorAction SilentlyContinue

    $policyPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive'
    if (-not (Test-Path $policyPath)) { New-Item -Path $policyPath -Force | Out-Null }
    Set-ItemProperty -Path $policyPath -Name 'DisableFileSyncNGSC'              -Value 1 -Type DWord
    Set-ItemProperty -Path $policyPath -Name 'DisableFileSync'                  -Value 1 -Type DWord
    Set-ItemProperty -Path $policyPath -Name 'PreventNetworkTrafficPreUserSignIn' -Value 1 -Type DWord
    Write-Log "Policy DisableFileSyncNGSC activée (impacte aussi OneDrive Business)." 'OK'

    foreach ($clsid in @(
        'Registry::HKEY_CLASSES_ROOT\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}',
        'Registry::HKEY_CLASSES_ROOT\Wow6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}'
    )) {
        if (Test-Path $clsid) {
            Set-ItemProperty -Path $clsid -Name 'System.IsPinnedToNameSpaceTree' `
                -Value 0 -Type DWord -ErrorAction SilentlyContinue
        }
    }

    foreach ($p in @(
        "$env:LOCALAPPDATA\Microsoft\OneDrive",
        "$env:PROGRAMDATA\Microsoft OneDrive",
        "$env:SystemDrive\OneDriveTemp"
    )) {
        if (Test-Path -LiteralPath $p) {
            Remove-Item -LiteralPath $p -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    $userOneDrive = "$env:USERPROFILE\OneDrive"
    if (Test-Path -LiteralPath $userOneDrive) {
        $hasContent = [System.IO.Directory]::EnumerateFileSystemEntries($userOneDrive) | Select-Object -First 1
        if ($hasContent) {
            Write-Log "Résidu dans $userOneDrive — vérifier manuellement." 'WARN'
        } else {
            Remove-Item -LiteralPath $userOneDrive -Recurse -Force -ErrorAction SilentlyContinue
            Write-Log "$userOneDrive vide — supprimé." 'OK'
        }
    }
    Write-Log "OneDrive désinstallé et bloqué." 'OK'
}

function Restore-OneDrivePolicies {
    Write-Log "=== Restauration policies OneDrive ===" 'INFO'
    $tasks = @(Get-ScheduledTask -TaskName 'OneDrive*' -ErrorAction SilentlyContinue)
    if ($tasks.Count -gt 0) {
        Write-Log "Tâches qui seront réactivées : $($tasks.TaskName -join ', ')" 'WARN'
    }
    if (-not $Force) {
        $reply = Read-Host "Confirmer la restauration ? (o/N)"
        if ($reply -notmatch '^[oOyY]') { Write-Log "Restauration annulée." 'WARN'; return }
    }
    $policyPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive'
    foreach ($v in @('DisableFileSyncNGSC','DisableFileSync','PreventNetworkTrafficPreUserSignIn')) {
        Remove-ItemProperty -Path $policyPath -Name $v -ErrorAction SilentlyContinue
    }
    foreach ($clsid in @(
        'Registry::HKEY_CLASSES_ROOT\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}',
        'Registry::HKEY_CLASSES_ROOT\Wow6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}'
    )) {
        if (Test-Path $clsid) {
            Set-ItemProperty -Path $clsid -Name 'System.IsPinnedToNameSpaceTree' `
                -Value 1 -Type DWord -ErrorAction SilentlyContinue
        }
    }
    Get-ScheduledTask -TaskName 'OneDrive*' -ErrorAction SilentlyContinue |
        ForEach-Object {
            try { Enable-ScheduledTask -TaskName $_.TaskName -TaskPath $_.TaskPath -ErrorAction Stop | Out-Null } catch {}
        }
    Write-Log "Policies OneDrive restaurées. Réinstalle OneDrive depuis https://onedrive.com si besoin." 'OK'
}

# --- Main ---

try {
    Write-Log "=== Google Drive Sync Setup ===" 'INFO'
    Write-Log "Journal : $script:LogFile"

    if ($RefreshQuickAccess) {
        $root = Get-DriveRootPath
        if (-not $root) {
            throw "Impossible de détecter le dossier racine. Documents n'a pas l'air d'être redirigé vers Drive."
        }
        Update-QuickAccessPins -RootPath $root
        Write-Log "=== Refresh Accès rapide terminé ===" 'OK'
        return
    }

    Assert-Admin

    if ($InstallScheduledRefresh) {
        Register-DriveSyncRefreshTask -IntervalMinutes $IntervalMinutes
        Write-Log "=== Installation tâche planifiée terminée ===" 'OK'
        return
    }

    if ($UninstallScheduledRefresh) {
        Unregister-DriveSyncRefreshTask
        Write-Log "=== Suppression tâche planifiée terminée ===" 'OK'
        return
    }

    if ($RestoreOneDrive) {
        Restore-OneDrivePolicies
        Write-Log "=== Restauration terminée ===" 'OK'
        return
    }

    if ($IncludeDevConfig) { Assert-SshKeysSafe }

    if (-not $SkipInstall) { Install-GoogleDrive }
    Test-OneDriveConflict

    $drivePath = Find-GoogleDriveMount -TimeoutSeconds $MountTimeoutSeconds
    $rootPath  = Join-Path $drivePath $RootName
    if (-not (Test-Path -LiteralPath $rootPath)) {
        New-Item -ItemType Directory -Path $rootPath -Force | Out-Null
        Write-Log "Dossier racine créé : $rootPath" 'OK'
    }

    $toRedirect = @('Documents','Pictures','Videos','Music','3D Objects')
    if ($IncludeDesktop) { $toRedirect += 'Desktop' }
    $refreshList = @()

    foreach ($name in $toRedirect) {
        $kf      = $KnownFolders[$name]
        $oldPath = Get-CurrentKnownFolderPath -RegName $kf.Reg
        if (-not $oldPath) { $oldPath = $kf.Local }
        $newPath = Join-Path $rootPath $name

        # 3D Objects désactivé par défaut sur Win11 22H2+ : skip sauf si -Force3DObjects
        if ($name -eq '3D Objects' -and -not (Test-Path -LiteralPath $oldPath)) {
            if ($Force3DObjects) {
                New-Item -ItemType Directory -Path $oldPath -Force | Out-Null
                Write-Log "3D Objects créé (Force3DObjects) : $oldPath" 'OK'
            } else {
                Write-Log "3D Objects absent (Win11 22H2+) — skip. Utilise -Force3DObjects pour le créer."
                continue
            }
        }

        Write-Log "--- $name ---"
        if (-not (Test-Path -LiteralPath $newPath)) {
            New-Item -ItemType Directory -Path $newPath -Force | Out-Null
        }
        Move-FolderContents -Source $oldPath -Destination $newPath
        Set-KnownFolder -Name $name -TargetPath $newPath
        $refreshList += $newPath
    }

    $extras = @()
    if ($IncludeScripts)         { $extras += 'Scripts' }
    if ($IncludeDevConfig)       { $extras += @('SSH','GitConfig') }
    if ($IncludeOfficeTemplates) { $extras += @('OfficeTemplates','OutlookSig') }

    if ($extras.Count -gt 0) {
        Write-Log "=== Mappings extras (symlinks) ===" 'INFO'
        foreach ($e in $extras) {
            try { Update-ExtraMapping -Name $e -RootPath $rootPath }
            catch { Write-Log "Échec sur '$e' : $($_.Exception.Message)" 'ERROR' }
        }
    }

    if ($DisableOneDrive) { Disable-OneDrive }

    Update-QuickAccessPins -RootPath $rootPath

    Invoke-ExplorerRefresh -UpdatedPaths $refreshList

    Write-Log "=== Terminé ===" 'OK'
    Write-Log "Pour que tous les apps voient les nouveaux chemins : redémarre l'Explorateur" 'INFO'
    Write-Log "  Stop-Process -Name explorer -Force"

} catch [System.Management.Automation.PipelineStoppedException] {
    Write-Log "Interrompu par l'utilisateur (Ctrl+C)." 'WARN'
    exit 130
} catch {
    Write-Log $_.Exception.Message 'ERROR'
    Write-Log "Voir : $script:LogFile" 'ERROR'
    exit 1
}
