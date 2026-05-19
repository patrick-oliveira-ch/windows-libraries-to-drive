# windows-libraries-to-drive

Synchronise les **dossiers utilisateur Windows** (Documents, Images, Vidéos, Musique, Objets 3D, Bureau) avec **Google Drive** pour retrouver les mêmes fichiers sur plusieurs PC, sans dépendre de OneDrive.

Le script utilise les APIs Windows officielles (`SHSetKnownFolderPath`) pour rediriger les Known Folders, comme le ferait Windows lui-même quand tu changes l'emplacement via clic-droit > Propriétés.

## Ce qui est synchronisé

### Known Folders (par défaut)
- Documents
- Images (Pictures)
- Vidéos (Videos)
- Musique (Music)
- Objets 3D (3D Objects) — skippé si absent (Windows 11 22H2+), créé via `-Force3DObjects`
- Bureau (Desktop) — opt-in via `-IncludeDesktop`

### Extras (opt-in via symlinks)
| Flag | Dossiers |
|---|---|
| `-IncludeScripts` | `%USERPROFILE%\Scripts\` |
| `-IncludeDevConfig` | `~\.ssh\` + `~\.gitconfig` |
| `-IncludeOfficeTemplates` | Templates Word/Excel + Signatures Outlook |

## Prérequis

- Windows 10 / 11
- PowerShell 5.1+ (inclus par défaut)
- Droits administrateur
- Un compte Google

## Utilisation

### Installation rapide

Deux modes au choix :

**Mode GUI** (recommandé pour débuter) :
1. Double-clic sur `RunGUI.bat` → interface graphique de configuration
2. Coche les options voulues, ajuste, clique "Lancer"
3. Le script principal s'auto-élève via UAC et exécute

**Mode CLI** (avec arguments) :
1. Double-clic sur `Run.bat` → lance avec les défauts
2. Ou en ligne de commande : `.\Install-GoogleDriveSync.ps1 -IncludeScripts ...`

Dans les deux cas : le script installe Google Drive for Desktop si nécessaire (connecte ton compte Google quand demandé) et lance la sync.

### Sur un nouveau PC

Re-cloner le repo + relancer `Run.bat`. Une fois Google Drive synchronisé, le script crée juste les redirections vers les dossiers déjà présents dans le cloud — tes fichiers sont immédiatement accessibles.

### Options principales

```powershell
# Installation standard
.\Install-GoogleDriveSync.ps1

# Avec Bureau + Scripts + Office
.\Install-GoogleDriveSync.ps1 -IncludeDesktop -IncludeScripts -IncludeOfficeTemplates

# Forcer la creation du dossier Objets 3D (Win11 22H2+ par defaut sans)
.\Install-GoogleDriveSync.ps1 -Force3DObjects

# Synchroniser config dev (SSH + .gitconfig)
.\Install-GoogleDriveSync.ps1 -IncludeDevConfig

# Migrer vers Drive ET désinstaller OneDrive
.\Install-GoogleDriveSync.ps1 -DisableOneDrive -Force

# Annuler les policies OneDrive si on change d'avis
.\Install-GoogleDriveSync.ps1 -RestoreOneDrive

# Nom de dossier racine personnalisé sur Drive
.\Install-GoogleDriveSync.ps1 -RootName "SyncWindows"
```

Toutes les options : `Get-Help .\Install-GoogleDriveSync.ps1 -Full`

## Arborescence créée sur Google Drive

```
Mon Drive/
└── WindowsLibraries/           (configurable via -RootName)
    ├── Documents/
    ├── Pictures/
    ├── Videos/
    ├── Music/
    ├── 3D Objects/
    ├── Desktop/                (si -IncludeDesktop)
    ├── Scripts/                (si -IncludeScripts)
    ├── DevConfig/              (si -IncludeDevConfig)
    │   ├── .ssh/
    │   └── .gitconfig
    └── Office/                 (si -IncludeOfficeTemplates)
        ├── Templates/
        └── Signatures/
```

## Sécurité

### Vérifications automatiques
- TLS 1.2/1.3 forcé pour le téléchargement
- Signature **Authenticode** vérifiée sur `GoogleDriveSetup.exe` (signataire = Google LLC)
- `$RootName` validé via regex (anti path traversal)
- **Bloque** la sync si une clé SSH privée **sans passphrase** est détectée (sauf `-Force`)
- Symlinks supprimés via `[System.IO.Directory]::Delete` pour éviter de suivre la cible (bug PS 5.1)
- Robocopy avec retries `/R:5 /W:5`, throw sur code d'erreur ≥ 8

### Risques connus à comprendre
- **Clés SSH** : `-IncludeDevConfig` met tes clés privées sur Google Drive. Toujours utiliser une passphrase.
- **`.gitconfig`** : si tu as un `[credential]` helper avec un token GitHub en clair, il sera publié sur Drive.
- **OneDrive disable** : `-DisableOneDrive` modifie `HKLM\SOFTWARE\Policies\...\OneDrive` (DisableFileSyncNGSC=1), ce qui désactive **aussi** OneDrive for Business. Utiliser `-RestoreOneDrive` pour annuler.

## Comment ça marche techniquement

1. **Known Folders** : `SHSetKnownFolderPath` (API officielle) + mise à jour de `HKCU\...\User Shell Folders`. Notification de l'Explorateur via `SHChangeNotify`.
2. **Migration** : `robocopy /MOVE /E` pour préserver attributs et structure.
3. **Extras** : `New-Item -ItemType SymbolicLink` (requiert admin), avec migration + suppression source.
4. **Détection Drive** : recherche d'un dossier `Mon Drive` ou `My Drive` sur toutes les lettres montées.

## Désinstallation / Rollback

Le script n'inclut pas (encore) de mode `-Uninstall` complet. Pour annuler manuellement :

```powershell
# Restaurer les chemins par défaut (via clic-droit > Propriétés > Emplacement > Restaurer)
# OU via registre, restaurer chaque valeur dans :
# HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders

# Supprimer les symlinks créés (exemples)
Remove-Item C:\Users\<user>\Scripts -Force  # si c'est un symlink

# Annuler les policies OneDrive
.\Install-GoogleDriveSync.ps1 -RestoreOneDrive
```

## Logs

Chaque run produit un log UTF-8 dans `%TEMP%\GoogleDriveSync_<datetime>.log`. Les logs >30 jours sont purgés automatiquement au démarrage.

## Licence

MIT — voir [LICENSE](LICENSE).

## Disclaimer

Script fourni "as-is". Teste sur une machine non critique avant déploiement large. Auteur non responsable de pertes de données.
