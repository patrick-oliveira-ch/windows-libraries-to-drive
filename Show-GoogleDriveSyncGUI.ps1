#Requires -Version 5.1
<#
.SYNOPSIS
    Interface graphique de configuration pour Install-GoogleDriveSync.ps1.
.DESCRIPTION
    Génère la ligne de commande à partir des choix utilisateur, puis lance le script
    principal dans une console séparée (le script s'auto-élève via UAC si besoin).
    Fenêtre scrollable pour s'adapter aux écrans plus petits.
#>
[CmdletBinding()]
param()

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

$scriptDir  = Split-Path -Parent $PSCommandPath
$mainScript = Join-Path $scriptDir 'Install-GoogleDriveSync.ps1'
if (-not (Test-Path -LiteralPath $mainScript)) {
    [void][System.Windows.Forms.MessageBox]::Show(
        "Install-GoogleDriveSync.ps1 introuvable dans :`n$scriptDir",
        'Erreur', 'OK', 'Error')
    exit 1
}

# --- Form ---
$form = New-Object System.Windows.Forms.Form
$form.Text            = 'Google Drive Sync — Configuration'
$form.ClientSize      = New-Object System.Drawing.Size(580, 600)
$form.MinimumSize     = New-Object System.Drawing.Size(600, 400)
$form.StartPosition   = 'CenterScreen'
$form.FormBorderStyle = 'Sizable'
$form.MaximizeBox     = $true
$form.Font            = New-Object System.Drawing.Font('Segoe UI', 9)

# --- Helpers ---
function New-Check {
    param([string]$Text, [int]$X, [int]$Y, [int]$W = 480, [string]$Tip)
    $c = New-Object System.Windows.Forms.CheckBox
    $c.Text = $Text
    $c.Location = New-Object System.Drawing.Point($X, $Y)
    $c.Size     = New-Object System.Drawing.Size($W, 22)
    if ($Tip) { $script:tooltip.SetToolTip($c, $Tip) }
    return $c
}

function New-Label {
    param([string]$Text, [int]$X, [int]$Y, [int]$W = 200)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Text
    $l.Location = New-Object System.Drawing.Point($X, $Y)
    $l.Size = New-Object System.Drawing.Size($W, 20)
    return $l
}

function New-GroupBox {
    param([string]$Text, [int]$X, [int]$Y, [int]$W, [int]$H)
    $g = New-Object System.Windows.Forms.GroupBox
    $g.Text = $Text
    $g.Location = New-Object System.Drawing.Point($X, $Y)
    $g.Size = New-Object System.Drawing.Size($W, $H)
    return $g
}

$script:tooltip = New-Object System.Windows.Forms.ToolTip
$tooltip.AutoPopDelay = 12000
$tooltip.InitialDelay = 400

# === Panneau scrollable (Dock=Fill) ===
$scrollPanel = New-Object System.Windows.Forms.Panel
$scrollPanel.Dock       = 'Fill'
$scrollPanel.AutoScroll = $true
$scrollPanel.Padding    = New-Object System.Windows.Forms.Padding(0, 0, 0, 0)

# === Section 1 : Dossiers extras ===
$grpFolders = New-GroupBox 'Dossiers à synchroniser (extras)' 10 10 540 195

$cbDesktop = New-Check '-IncludeDesktop  (synchronise le Bureau)' 15 25 510 `
    "Ajoute le dossier Bureau aux Known Folders redirigés. Les raccourcis (.lnk/.url) sont exclus par défaut et déplacés vers Public Desktop."
$cbDesktopShortcuts = New-Check '   ↳ -IncludeDesktopShortcuts  (syncer aussi les .lnk/.url)' 30 47 495 `
    "DÉCONSEILLÉ : les raccourcis pointent vers des programmes locaux et seront cassés sur les autres PC."
$cbDesktopShortcuts.Enabled = $false
$cb3D = New-Check '-Force3DObjects  (crée Objets 3D si absent — Win11 22H2+)' 15 72 510 `
    "Sans cette option, 3D Objects est skippé s'il n'existe pas."
$cbScripts = New-Check '-IncludeScripts  (~\Scripts\)' 15 97 510 `
    "Symlink vers C:\Users\<user>\Scripts\ — utile pour porter tes scripts multi-PC."
$cbDev = New-Check '-IncludeDevConfig  (.ssh\ + .gitconfig)' 15 122 510 `
    "BLOQUÉ si une clé SSH privée non chiffrée est détectée (sauf -Force)."
$cbOffice = New-Check '-IncludeOfficeTemplates  (Templates Word/Excel + Signatures Outlook)' 15 147 510 `
    "Symlinks vers %APPDATA%\Microsoft\Templates et \Signatures."

# Active/désactive le sub-toggle shortcuts selon Desktop
$cbDesktop.Add_CheckedChanged({
    $cbDesktopShortcuts.Enabled = $cbDesktop.Checked
    if (-not $cbDesktop.Checked) { $cbDesktopShortcuts.Checked = $false }
})

$grpFolders.Controls.AddRange(@($cbDesktop, $cbDesktopShortcuts, $cb3D, $cbScripts, $cbDev, $cbOffice))

# === Section 2 : Configuration ===
$grpConfig = New-GroupBox 'Configuration' 10 215 540 145

$grpConfig.Controls.Add((New-Label 'Nom du dossier racine sur Drive :' 15 30 220))
$txtRoot = New-Object System.Windows.Forms.TextBox
$txtRoot.Location = New-Object System.Drawing.Point(245, 28)
$txtRoot.Size = New-Object System.Drawing.Size(280, 22)
$txtRoot.Text = 'WindowsLibraries'
$grpConfig.Controls.Add($txtRoot)

$grpConfig.Controls.Add((New-Label 'Lettre du drive Google Drive (vide = auto) :' 15 60 220))
$txtLetter = New-Object System.Windows.Forms.TextBox
$txtLetter.Location = New-Object System.Drawing.Point(245, 58)
$txtLetter.Size = New-Object System.Drawing.Size(50, 22)
$txtLetter.MaxLength = 1
$grpConfig.Controls.Add($txtLetter)

$grpConfig.Controls.Add((New-Label 'Timeout de montage (secondes) :' 15 90 220))
$numTimeout = New-Object System.Windows.Forms.NumericUpDown
$numTimeout.Location = New-Object System.Drawing.Point(245, 88)
$numTimeout.Size = New-Object System.Drawing.Size(80, 22)
$numTimeout.Minimum = 60
$numTimeout.Maximum = 3600
$numTimeout.Value = 600
$grpConfig.Controls.Add($numTimeout)

$cbSkipInstall = New-Check '-SkipInstall  (Google Drive déjà installé et connecté)' 15 115 510
$grpConfig.Controls.Add($cbSkipInstall)

# === Section 3 : OneDrive ===
$grpOneDrive = New-GroupBox 'OneDrive' 10 370 540 100

$cbDisableOD = New-Check '-DisableOneDrive  (migrer puis désinstaller OneDrive)' 15 25 510 `
    "Désinstalle OneDrive et bloque sa réinstallation via policy HKLM. Impact: désactive aussi OneDrive for Business."

$btnRestore = New-Object System.Windows.Forms.Button
$btnRestore.Text = 'Restaurer les policies OneDrive (annule -DisableOneDrive)'
$btnRestore.Location = New-Object System.Drawing.Point(15, 55)
$btnRestore.Size = New-Object System.Drawing.Size(510, 30)
$tooltip.SetToolTip($btnRestore, "Lance le script en mode -RestoreOneDrive sans toucher aux redirections de dossiers.")

$grpOneDrive.Controls.AddRange(@($cbDisableOD, $btnRestore))

# === Section 4 : Accès rapide Windows ===
$grpQA = New-GroupBox 'Accès rapide Windows' 10 480 540 195

$lblQA = New-Label "Épingle les dossiers Drive (Scripts, 3D Objects, etc.) à l'Accès rapide Explorer." 15 22 510
$lblQA.AutoSize = $false
$lblQA.Size = New-Object System.Drawing.Size(510, 20)

$btnRefreshQA = New-Object System.Windows.Forms.Button
$btnRefreshQA.Text = "Réépingler maintenant (one-shot)"
$btnRefreshQA.Location = New-Object System.Drawing.Point(15, 50)
$btnRefreshQA.Size = New-Object System.Drawing.Size(510, 30)
$tooltip.SetToolTip($btnRefreshQA, "Lance -RefreshQuickAccess : pas besoin d'admin, n'épingle que les dossiers extras (skip Documents/Pictures/etc. déjà visibles).")

# Sous-section tâche planifiée
$lblSched = New-Label "Tâche planifiée auto-refresh (toutes les N minutes)" 15 95 510
$lblSched.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)

$lblInterval = New-Label 'Intervalle (minutes) :' 15 120 130
$numInterval = New-Object System.Windows.Forms.NumericUpDown
$numInterval.Location = New-Object System.Drawing.Point(145, 118)
$numInterval.Size = New-Object System.Drawing.Size(70, 22)
$numInterval.Minimum = 5
$numInterval.Maximum = 1440
$numInterval.Value = 15

$btnInstallSched = New-Object System.Windows.Forms.Button
$btnInstallSched.Text = 'Activer'
$btnInstallSched.Location = New-Object System.Drawing.Point(225, 116)
$btnInstallSched.Size = New-Object System.Drawing.Size(140, 26)
$tooltip.SetToolTip($btnInstallSched, "Crée une tâche planifiée Windows qui run -RefreshQuickAccess à la connexion + toutes les N min.")

$btnUninstallSched = New-Object System.Windows.Forms.Button
$btnUninstallSched.Text = 'Désactiver'
$btnUninstallSched.Location = New-Object System.Drawing.Point(370, 116)
$btnUninstallSched.Size = New-Object System.Drawing.Size(155, 26)
$tooltip.SetToolTip($btnUninstallSched, "Supprime la tâche planifiée.")

$lblSchedState = New-Label '' 15 150 510
$lblSchedState.ForeColor = [System.Drawing.Color]::DarkGreen

function Update-ScheduledState {
    $task = Get-ScheduledTask -TaskName 'GoogleDriveSync-RefreshQuickAccess' -ErrorAction SilentlyContinue
    if ($task) {
        $info = Get-ScheduledTaskInfo -InputObject $task -ErrorAction SilentlyContinue
        $last = if ($info -and $info.LastRunTime.Year -gt 1) { $info.LastRunTime.ToString('yyyy-MM-dd HH:mm') } else { 'jamais' }
        $next = if ($info -and $info.NextRunTime.Year -gt 1) { $info.NextRunTime.ToString('HH:mm') } else { '?' }
        $lblSchedState.Text = "État : Activée (dernier run : $last  |  prochain : $next)"
        $lblSchedState.ForeColor = [System.Drawing.Color]::DarkGreen
    } else {
        $lblSchedState.Text = 'État : Désactivée (aucune tâche planifiée)'
        $lblSchedState.ForeColor = [System.Drawing.Color]::DarkRed
    }
}
Update-ScheduledState

$grpQA.Controls.AddRange(@($lblQA, $btnRefreshQA, $lblSched, $lblInterval, $numInterval, $btnInstallSched, $btnUninstallSched, $lblSchedState))

# === Section 5 : Options globales ===
$grpOpts = New-GroupBox 'Options' 10 685 540 60
$cbForce = New-Check '-Force  (aucune confirmation interactive — mode automatique)' 15 25 510 `
    "ATTENTION : -Force autorise aussi la sync de clés SSH non chiffrées."
$grpOpts.Controls.Add($cbForce)

# === Aperçu commande ===
$grpPreview = New-GroupBox 'Aperçu de la commande' 10 755 540 100
$txtPreview = New-Object System.Windows.Forms.TextBox
$txtPreview.Location = New-Object System.Drawing.Point(10, 22)
$txtPreview.Size = New-Object System.Drawing.Size(520, 70)
$txtPreview.Multiline = $true
$txtPreview.ReadOnly = $true
$txtPreview.ScrollBars = 'Vertical'
$txtPreview.BackColor = [System.Drawing.Color]::White
$txtPreview.Font = New-Object System.Drawing.Font('Consolas', 9)
$grpPreview.Controls.Add($txtPreview)

# Ajouter toutes les sections au panneau scrollable
$scrollPanel.Controls.AddRange(@(
    $grpFolders, $grpConfig, $grpOneDrive, $grpQA, $grpOpts, $grpPreview
))

# === Panneau du bas (fixe, contient les boutons) ===
$bottomPanel = New-Object System.Windows.Forms.Panel
$bottomPanel.Dock        = 'Bottom'
$bottomPanel.Height      = 55
$bottomPanel.BorderStyle = 'FixedSingle'

$btnLaunch = New-Object System.Windows.Forms.Button
$btnLaunch.Text = 'Lancer'
$btnLaunch.Location = New-Object System.Drawing.Point(355, 12)
$btnLaunch.Size = New-Object System.Drawing.Size(105, 32)
$btnLaunch.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$btnLaunch.Anchor = 'Top, Right'

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = 'Annuler'
$btnCancel.Location = New-Object System.Drawing.Point(465, 12)
$btnCancel.Size = New-Object System.Drawing.Size(105, 32)
$btnCancel.DialogResult = 'Cancel'
$btnCancel.Anchor = 'Top, Right'

$bottomPanel.Controls.AddRange(@($btnLaunch, $btnCancel))

# Ordre d'ajout : Bottom d'abord, Fill ensuite (pour que Fill prenne l'espace restant)
$form.Controls.Add($bottomPanel)
$form.Controls.Add($scrollPanel)
$form.AcceptButton = $btnLaunch
$form.CancelButton = $btnCancel

# --- Génération des args ---
function Get-ArgsForRestore { return @('-RestoreOneDrive') }
function Get-ArgsForRefreshQA { return @('-RefreshQuickAccess') }
function Get-ArgsForInstallSched { param([int]$Interval) return @('-InstallScheduledRefresh','-IntervalMinutes',$Interval.ToString()) }
function Get-ArgsForUninstallSched { return @('-UninstallScheduledRefresh') }

function Get-ArgsFromForm {
    $a = @()
    if ($txtRoot.Text -and $txtRoot.Text -ne 'WindowsLibraries') {
        $a += @('-RootName', $txtRoot.Text)
    }
    if ($cbDesktop.Checked)     { $a += '-IncludeDesktop' }
    if ($cbDesktopShortcuts.Checked) { $a += '-IncludeDesktopShortcuts' }
    if ($cb3D.Checked)          { $a += '-Force3DObjects' }
    if ($cbScripts.Checked)     { $a += '-IncludeScripts' }
    if ($cbDev.Checked)         { $a += '-IncludeDevConfig' }
    if ($cbOffice.Checked)      { $a += '-IncludeOfficeTemplates' }
    if ($cbSkipInstall.Checked) { $a += '-SkipInstall' }
    if ($cbDisableOD.Checked)   { $a += '-DisableOneDrive' }
    if ($cbForce.Checked)       { $a += '-Force' }
    if ($txtLetter.Text -match '^[A-Za-z]$') { $a += @('-DriveLetter', $txtLetter.Text) }
    if ($numTimeout.Value -ne 600) { $a += @('-MountTimeoutSeconds', $numTimeout.Value.ToString()) }
    return $a
}

function Format-Preview {
    param([string[]]$Args)
    $line = ".\Install-GoogleDriveSync.ps1"
    foreach ($x in $Args) {
        if ($x -match '\s') { $line += ' "' + $x + '"' } else { $line += ' ' + $x }
    }
    return $line
}

function Update-Preview { $txtPreview.Text = Format-Preview (Get-ArgsFromForm) }
Update-Preview

# Hook tous les contrôles pour rafraîchir l'aperçu
foreach ($c in @($cbDesktop, $cbDesktopShortcuts, $cb3D, $cbScripts, $cbDev, $cbOffice,
                 $cbSkipInstall, $cbDisableOD, $cbForce)) {
    $c.Add_CheckedChanged({ Update-Preview })
}
$txtRoot.Add_TextChanged({ Update-Preview })
$txtLetter.Add_TextChanged({ Update-Preview })
$numTimeout.Add_ValueChanged({ Update-Preview })

# --- Lancement du script ---
function Invoke-MainScript {
    param([string[]]$ScriptArgs)
    $psArgs = @('-NoExit','-NoProfile','-ExecutionPolicy','Bypass','-File',"`"$mainScript`"") + $ScriptArgs
    $argString = ($psArgs | ForEach-Object {
        if ($_ -match '\s' -and -not $_.StartsWith('"')) { '"' + $_ + '"' } else { $_ }
    }) -join ' '
    Start-Process -FilePath 'powershell.exe' -ArgumentList $argString
}

$btnLaunch.Add_Click({
    $a = Get-ArgsFromForm
    if ($cbDisableOD.Checked -and -not $cbForce.Checked) {
        $r = [System.Windows.Forms.MessageBox]::Show(
            "Tu as coché -DisableOneDrive. OneDrive sera DÉSINSTALLÉ et bloqué (impacte aussi OneDrive Business).`n`nContinuer ?",
            'Confirmation', 'YesNo', 'Warning')
        if ($r -ne 'Yes') { return }
    }
    Invoke-MainScript -ScriptArgs $a
    $form.Close()
})

$btnRestore.Add_Click({
    $r = [System.Windows.Forms.MessageBox]::Show(
        "Lancer le script en mode -RestoreOneDrive ?`n`nAnnule les policies de blocage OneDrive (ne réinstalle pas OneDrive, ne touche pas aux redirections de dossiers).",
        'Confirmation', 'YesNo', 'Question')
    if ($r -ne 'Yes') { return }
    Invoke-MainScript -ScriptArgs (Get-ArgsForRestore)
    $form.Close()
})

$btnRefreshQA.Add_Click({
    Invoke-MainScript -ScriptArgs (Get-ArgsForRefreshQA)
    $form.Close()
})

$btnInstallSched.Add_Click({
    $interval = [int]$numInterval.Value
    $r = [System.Windows.Forms.MessageBox]::Show(
        "Activer la tâche planifiée auto-refresh ?`n`nElle s'exécutera à chaque connexion + toutes les $interval minutes.`nUAC sera demandé.",
        'Confirmation', 'YesNo', 'Question')
    if ($r -ne 'Yes') { return }
    Invoke-MainScript -ScriptArgs (Get-ArgsForInstallSched -Interval $interval)
    $form.Close()
})

$btnUninstallSched.Add_Click({
    $r = [System.Windows.Forms.MessageBox]::Show(
        "Désactiver la tâche planifiée auto-refresh ?`n`nUAC sera demandé.",
        'Confirmation', 'YesNo', 'Question')
    if ($r -ne 'Yes') { return }
    Invoke-MainScript -ScriptArgs (Get-ArgsForUninstallSched)
    $form.Close()
})

[void]$form.ShowDialog()
