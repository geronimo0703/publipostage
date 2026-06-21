 # Crée un raccourci Bureau pour Publipostage.
# À lancer une seule fois après extraction du zip, depuis le même dossier
# que l'exécutable publipostage-windows.exe.
#
# Si Windows refuse l'exécution (script bloqué car téléchargé depuis
# Internet), faire un clic droit sur ce fichier > Propriétés >
# cocher "Débloquer", ou lancer dans PowerShell :
#   Unblock-File .\creer_raccourci_windows.ps1

$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ExePath = Join-Path $ScriptDir "publipostage-windows.exe"

if (-not (Test-Path $ExePath)) {
    Write-Host "Erreur : executable introuvable a $ExePath"
    Write-Host "Assurez-vous que ce script reste dans le meme dossier que publipostage-windows.exe."
    exit 1
}

$DesktopDir = [Environment]::GetFolderPath("Desktop")
$ShortcutPath = Join-Path $DesktopDir "Publipostage.lnk"

$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = $ExePath
$Shortcut.WorkingDirectory = $ScriptDir
$Shortcut.Description = "Outil de publipostage - Banque Alimentaire 22"
$Shortcut.Save()

Write-Host "Raccourci cree sur le Bureau : $ShortcutPath"
Write-Host "Vous pouvez maintenant double-cliquer dessus pour lancer Publipostage."
