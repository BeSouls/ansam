<#
.SYNOPSIS
    Crée la tâche planifiée Windows « Synthèse mails ANSAM » : chaque jour à
    l'heure de la configuration (7h10 par défaut), pour une synthèse prête à
    7h30.

.DESCRIPTION
    La tâche se déclenche tous les jours ; c'est Lancer-SyntheseMails.ps1 qui
    décide, selon la configuration, s'il faut produire une synthèse (jours
    d'envoi, absences). Elle tourne sous ta session (Outlook doit être
    accessible) et uniquement quand tu es connecté·e. Si le PC est éteint à
    l'heure prévue, elle démarre dès que possible. Relancer le script
    remplace la tâche.

    La tâche se déclenche aussi à l'ouverture de session et au
    déverrouillage, pour rattraper un PC en veille à l'heure prévue ; elle
    tourne sur batterie et retente 3 fois en cas d'échec.

    Crée aussi « <TaskName> - page » : la page de suivi et de configuration,
    lancée à l'ouverture de session et gardée active (adresse fixe, lien dans
    chaque synthèse), démarrée tout de suite, et un raccourci « Synthèse
    mails – suivi » dans le menu Démarrer.

.EXAMPLE
    .\Installer-TacheSyntheseMails.ps1
    .\Installer-TacheSyntheseMails.ps1 -At '06:45'
    Unregister-ScheduledTask -TaskName 'Synthese mails ANSAM'          # désinstaller
    Unregister-ScheduledTask -TaskName 'Synthese mails ANSAM - page'
#>
[CmdletBinding()]
param(
    # Par défaut : l'heure de la configuration.
    [string]$At,
    [string]$TaskName = 'Synthese mails ANSAM'
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'SyntheseMails.Commun.ps1')
if (-not $At) { $At = (Get-SmConfig).heure }
$script = Join-Path $PSScriptRoot 'Lancer-SyntheseMails.ps1'

$user = "$env:USERDOMAIN\$env:USERNAME"
# conhost --headless : console sans fenêtre. Sans lui, Windows 11 confie la console à Windows Terminal
# et le script lancé par le planificateur peut être tué au démarrage (0xC000013A, 0x8007042B).
$action = New-ScheduledTaskAction -Execute "$env:WINDIR\System32\conhost.exe" `
    -Argument "--headless powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$script`""
# Un portable en veille à l'heure prévue rate le déclenchement, et au réveil PowerShell peut échouer
# avant même de démarrer (bureau verrouillé). D'où trois déclencheurs : l'heure prévue, l'ouverture de
# session et le déverrouillage. Le lanceur ne produit qu'une synthèse par jour, jamais avant l'heure.
$unlockClass = Get-CimClass -Namespace 'Root/Microsoft/Windows/TaskScheduler' -ClassName 'MSFT_TaskSessionStateChangeTrigger'
$unlock = New-CimInstance -CimClass $unlockClass -ClientOnly
$unlock.StateChange = 8   # TASK_SESSION_UNLOCK
$unlock.UserId = $user
$unlock.Delay = 'PT1M'
$triggers = @(
    (New-ScheduledTaskTrigger -Daily -At $At),
    (New-ScheduledTaskTrigger -AtLogOn -User $user),
    $unlock
)
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 45) `
    -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 5)
$principal = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $triggers `
    -Settings $settings -Principal $principal -Force | Out-Null
Write-Host "Tâche « $TaskName » créée : tous les jours à $At, rattrapée à l'ouverture de session et au déverrouillage."

# Page de suivi permanente : démarrée à l'ouverture de session, sans limite de durée.
$pageTask = "$TaskName - page"
$pageScript = Join-Path $PSScriptRoot 'Configurer-SyntheseMails.ps1'
$pageAction = New-ScheduledTaskAction -Execute "$env:WINDIR\System32\conhost.exe" `
    -Argument "--headless powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$pageScript`" -Service"
$pageTrigger = New-ScheduledTaskTrigger -AtLogOn -User "$env:USERDOMAIN\$env:USERNAME"
$pageSettings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit ([TimeSpan]::Zero) -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)
Register-ScheduledTask -TaskName $pageTask -Action $pageAction -Trigger $pageTrigger `
    -Settings $pageSettings -Principal $principal -Force | Out-Null
# Arrêter la tâche ne tue pas le PowerShell lancé par conhost : on arrête l'ancien service
# pour que la page redémarre avec la version installée.
Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" |
    Where-Object { $_.CommandLine -match 'Configurer-SyntheseMails\.ps1' -and $_.CommandLine -match '-Service' -and $_.ProcessId -ne $PID } |
    ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
Start-Sleep -Seconds 1
Start-ScheduledTask -TaskName $pageTask
Write-Host "Tâche « $pageTask » créée : page permanente (redémarrée)."

# Raccourci du menu Démarrer vers la page, avec son jeton (le lien sans jeton est refusé).
$startMenu = Join-Path ([Environment]::GetFolderPath('Programs')) 'Synthèse mails – suivi.url'
[IO.File]::WriteAllText($startMenu, "[InternetShortcut]`r`nURL=$(Get-SmPageUrl)`r`n", [Text.Encoding]::ASCII)
Write-Host "Raccourci créé dans le menu Démarrer : « Synthèse mails – suivi »."
