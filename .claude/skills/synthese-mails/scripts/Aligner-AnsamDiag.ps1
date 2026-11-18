<#
.SYNOPSIS
    Aligne le statut d'un ticket AnsamDiag local sur la réalité (Jira, Panoramix), sur clic du TAM.

.DESCRIPTION
    Lancé par la page (bouton « Aligner AnsamDiag ») en arrière-plan. Passe par les fonctions du
    module AnsamDiag (Start-AnsamTicket, Set-AnsamTicket -Statut, Add-AnsamNote) : le dossier du
    ticket garde son format et sa note datée. Écrit seulement dans le dossier local du ticket ;
    rien dans Jira. État : <WorkDir>\travaux\ansam-<CLE>.json.

.EXAMPLE
    .\Aligner-AnsamDiag.ps1 -Cle ASD-27334 -Statut Resolu -Raison 'Jira : Résolu'
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [ValidatePattern('^[A-Za-z0-9-]+$')] [string]$Cle,
    [Parameter(Mandatory)] [ValidateSet('Ouvert', 'EnAttente', 'Resolu', 'Clos')] [string]$Statut,
    [string]$Raison = '',
    [string]$WorkDir = (Join-Path $env:USERPROFILE 'SyntheseMails')
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'SyntheseMails.Commun.ps1')
$statusFile = Join-Path $WorkDir "travaux\ansam-$Cle.json"
$started = (Get-Date).ToString('o')
function Set-Status([string]$State, [string]$Message) {
    Write-SmJson $statusFile ([ordered]@{ travail = 'ansam'; id = $Cle; etat = $State; message = $Message; le = $started; pid = $PID })
}
Set-Status 'en_cours' 'Ouverture du ticket AnsamDiag'

try {
    $data = [string](Get-SmConfig $WorkDir).dossierAnsamDiag
    if (-not (Test-Path -LiteralPath (Join-Path $data "Tickets\$Cle\ticket.json"))) { throw "Pas de dossier AnsamDiag pour $Cle." }
    $module = Join-Path (Split-Path $data) 'scripts\AnsamDiag\AnsamDiag.psd1'
    if (-not (Test-Path -LiteralPath $module)) { throw "Module AnsamDiag introuvable : $module" }
    Import-Module $module -Force -WarningAction SilentlyContinue | Out-Null
    Start-AnsamTicket $Cle | Out-Null
    $avant = [string](Get-AnsamTicket).Statut
    if ($avant -eq $Statut) { Set-Status 'ok' "$Cle est déjà $Statut dans AnsamDiag."; exit 0 }
    Set-AnsamTicket -Statut $Statut | Out-Null
    Add-AnsamNote ("Statut aligne depuis Panoramix : $avant -> $Statut" + $(if ($Raison) { " ($Raison)" } else { '' }) + '.') | Out-Null
    Set-Status 'ok' "$Cle : $avant -> $Statut dans AnsamDiag."
} catch {
    Set-Status 'echec' $_.Exception.Message
    exit 1
}
