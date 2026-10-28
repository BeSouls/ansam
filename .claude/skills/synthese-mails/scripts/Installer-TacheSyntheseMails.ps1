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

    Crée aussi « <TaskName> - page » : la page de suivi et de configuration,
    lancée à l'ouverture de session et gardée active (adresse fixe, lien dans
    chaque synthèse). Elle est démarrée tout de suite.

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

$action = New-ScheduledTaskAction -Execute 'powershell.exe' `
    -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$script`""
$trigger = New-ScheduledTaskTrigger -Daily -At $At
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 45)
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger `
    -Settings $settings -Principal $principal -Force | Out-Null
Write-Host "Tâche « $TaskName » créée : tous les jours à $At."

# Page de suivi permanente : démarrée à l'ouverture de session, sans limite de durée.
$pageTask = "$TaskName - page"
$pageScript = Join-Path $PSScriptRoot 'Configurer-SyntheseMails.ps1'
$pageAction = New-ScheduledTaskAction -Execute 'powershell.exe' `
    -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$pageScript`" -Service"
$pageTrigger = New-ScheduledTaskTrigger -AtLogOn -User "$env:USERDOMAIN\$env:USERNAME"
$pageSettings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit ([TimeSpan]::Zero) -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)
Register-ScheduledTask -TaskName $pageTask -Action $pageAction -Trigger $pageTrigger `
    -Settings $pageSettings -Principal $principal -Force | Out-Null
if ((Get-ScheduledTask -TaskName $pageTask).State -ne 'Running') { Start-ScheduledTask -TaskName $pageTask }
Write-Host "Tâche « $pageTask » créée : page permanente sur $(Get-SmPageUrl)"
