<#
.SYNOPSIS
    Report des calculs mensuels de variabilité dans le registre des relevés, lancé
    depuis Panoramix (ou à la main).

.DESCRIPTION
    controle   Montre, sans rien écrire, ce qui serait reporté pour chaque client :
               compteurs, conflits avec un relevé déjà validé, conditions de validation.
    reporter   Écrit les relevés (statut « a_verifier », ou « valide » avec -Valider).

    Les calculs viennent de l'outil AnsamDiag Variabilite.Metriques (3.1 ou plus),
    lancé pour le mois. Rien n'est saisi dans Kiosk ni écrit dans Jira.
    Résultat : <dossierAnsamDiag>\variabilite\releves\report-<AAAA-MM>.json ;
    avancement : <WorkDir>\travaux\releves-<etape>.json.

.EXAMPLE
    .\Releves-Etape.ps1 -Etape controle -Mois 2026-11 -Tous
    .\Releves-Etape.ps1 -Etape reporter -Mois 2026-11 -Client STJAC -Valider
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [ValidateSet('controle', 'reporter')] [string]$Etape,
    [Parameter(Mandatory)] [ValidatePattern('^\d{4}-(0[1-9]|1[0-2])$')] [string]$Mois,
    [string[]]$Client,
    [switch]$Tous,
    [switch]$Valider,
    [switch]$Remplacer,
    [string]$WorkDir = (Join-Path $env:USERPROFILE 'SyntheseMails')
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'SyntheseMails.Commun.ps1')
. (Join-Path $PSScriptRoot 'SyntheseMails.PlusValue.ps1')
. (Join-Path $PSScriptRoot 'SyntheseMails.Releves.ps1')
$data = [string](Get-SmConfig $WorkDir).dossierAnsamDiag
$statusFile = Join-Path $WorkDir "travaux\releves-$Etape.json"
function Set-Status([string]$State, [string]$Message) {
    Write-SmJson $statusFile ([ordered]@{ etape = $Etape; mois = $Mois; etat = $State; message = $Message; le = (Get-Date).ToString('o'); pid = $PID })
}
Set-Status 'en_cours' 'Démarrage'

try {
    $codes = if ($Tous) {
        @(Get-ChildItem -LiteralPath (Join-Path $data 'variabilite\_etat') -Directory | Where-Object { Test-Path (Join-Path $_.FullName "Mois-$Mois.json") } | ForEach-Object { $_.Name })
    } elseif ($Client) { @($Client | ForEach-Object { $_.ToUpper() }) } else { throw 'Indiquer -Client ou -Tous.' }

    $resultats = @()
    foreach ($c in $codes) {
        Set-Status 'en_cours' "Relevés $c"
        $r = Get-SmRvReport -Code $c -Mois $Mois -Data $data
        $r.Ecrit = $null
        if ($Etape -eq 'reporter' -and -not $r.Erreurs.Count) {
            try { $r.Ecrit = Save-SmRvReport -Report $r -Data $data -Valider:$Valider -Remplacer:$Remplacer }
            catch { $r.Erreurs += $_.Exception.Message }
        }
        $resultats += [pscustomobject]$r
    }
    $out = [ordered]@{ mois = $Mois; etape = $Etape; valider = [bool]$Valider; le = (Get-Date).ToString('o'); clients = $resultats }
    [IO.File]::WriteAllText((Join-Path $data "variabilite\releves\report-$Mois.json"), (ConvertTo-Json -InputObject $out -Depth 10), (New-Object Text.UTF8Encoding($false)))

    Set-Status 'ok' ("{0} client(s) sans erreur sur {1}" -f @($resultats | Where-Object { -not $_.Erreurs.Count }).Count, $resultats.Count)
    $resultats | ForEach-Object {
        $etat = if ($_.Erreurs.Count) { 'REFUSE' } elseif ($_.Ecrit) { $_.Ecrit.Statut } else { 'controle' }
        $lignes = ($_.Ecritures | ForEach-Object { '{0}={1}' -f $_.Cle, $_.Valeur }) -join ' '
        "{0,-6} {1,-10} {2}" -f $_.Code, $etat, $(if ($_.Erreurs.Count) { $_.Erreurs -join ' | ' } else { $lignes })
        if ($_.Conflits.Count) { "       conflits : " + (($_.Conflits | ForEach-Object { '{0} {1}->{2}' -f $_.Cle, $_.Avant, $_.Valeur }) -join ', ') }
        if ($_.Bloquants.Count) { "       validation impossible : " + ($_.Bloquants -join ' ; ') }
    }
} catch {
    Set-Status 'echec' $_.Exception.Message
    throw
}
