<#
.SYNOPSIS
    Crée la tâche planifiée Windows « Synthèse mails ANSAM » : chaque jour à
    7h10, pour une synthèse prête à 7h30.

.DESCRIPTION
    La tâche tourne sous ta session (Outlook doit être accessible) et
    uniquement quand tu es connecté·e. Si le PC est éteint à l'heure prévue,
    elle démarre dès que possible. Relancer le script remplace la tâche.

.EXAMPLE
    .\Installer-TacheSyntheseMails.ps1
    .\Installer-TacheSyntheseMails.ps1 -At '06:45'
    Unregister-ScheduledTask -TaskName 'Synthese mails ANSAM'   # désinstaller
#>
[CmdletBinding()]
param(
    [string]$At = '07:10',
    [string]$TaskName = 'Synthese mails ANSAM'
)

$ErrorActionPreference = 'Stop'
$script = Join-Path $PSScriptRoot 'Lancer-SyntheseMails.ps1'

$action = New-ScheduledTaskAction -Execute 'powershell.exe' `
    -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$script`""
$trigger = New-ScheduledTaskTrigger -Daily -At $At
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 30)
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger `
    -Settings $settings -Principal $principal -Force | Out-Null
Write-Host "Tâche « $TaskName » créée : tous les jours à $At."
