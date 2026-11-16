<#
.SYNOPSIS
    Tests du report des calculs mensuels dans le registre des relevés. Travaille dans
    une copie temporaire (registre et contrat STJAC, calculs fabriqués) : le vrai
    registre n'est jamais touché. Code de sortie 0 si tout passe.
#>
[CmdletBinding()]
param([string]$WorkDir = (Join-Path $env:USERPROFILE 'SyntheseMails'))
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'SyntheseMails.Commun.ps1')
. (Join-Path $PSScriptRoot 'SyntheseMails.PlusValue.ps1')
. (Join-Path $PSScriptRoot 'SyntheseMails.Releves.ps1')
$vrai = [string](Get-SmConfig $WorkDir).dossierAnsamDiag
$data = Join-Path ([IO.Path]::GetTempPath()) ("test-releves-{0}" -f (Get-Date).ToString('yyyyMMdd-HHmmss'))
New-Item -ItemType Directory -Force -Path "$data\Facturation\contrats", "$data\variabilite\releves", "$data\variabilite\_etat\STJAC" | Out-Null
Copy-Item (Join-Path $vrai 'Facturation\contrats\STJAC.json') "$data\Facturation\contrats\"
Copy-Item (Join-Path $vrai 'variabilite\releves\STJAC.json') "$data\variabilite\releves\"
$echecs = 0
function Test-Egal($Nom, $Attendu, $Obtenu) {
    if ("$Attendu" -eq "$Obtenu") { "  ok    $Nom = $Obtenu" } else { "  ECHEC $Nom : attendu $Attendu, obtenu $Obtenu"; $script:echecs++ }
}
function New-Calcul([string]$Mois, [hashtable]$Valeurs, [string]$Confirme, [string]$DateM365, [int]$ATrancher = 0, [switch]$Ancien) {
    $tab = foreach ($k in $SmRvMetriques.Keys) {
        $src = if ($k -in '#Users') { "M365 du $DateM365" } else { "Manuel (confirme le $Confirme)" }
        [ordered]@{ Client = 'STJAC'; Metrique = $k; MoisCourant = $(if ($Valeurs.ContainsKey($k)) { $Valeurs[$k] } else { 0 }); Source = $src }
    }
    $o = [ordered]@{ Client = 'STJAC'; Mois = $Mois; Calcule = "$DateM365 10:00"; ValeursManuellesVerifiees = $Confirme
                     Tableau = @($tab); ATrancher = @(1..$ATrancher | Where-Object { $ATrancher } | ForEach-Object { @{ Compte = "x$_" } }) }
    if (-not $Ancien) { $o.Inventaires = [ordered]@{ M365 = $DateM365; AD = '' } }
    [IO.File]::WriteAllText("$data\variabilite\_etat\STJAC\Mois-$Mois.json", (ConvertTo-Json $o -Depth 6), (New-Object Text.UTF8Encoding($false)))
}
function Get-Releve($Mois, $Cle) { (Read-SmPvJson "$data\variabilite\releves\STJAC.json").Releves.$Mois.$Cle }

'Report a_verifier puis validation (2026-11)'
New-Calcul '2026-11' @{ '#Users' = 46; '#VM' = 4; '#Site Actuel' = 2 } '2026-11-20' '2026-11-28'
$r = Get-SmRvReport -Code STJAC -Mois '2026-11' -Data $data
Test-Egal 'erreurs' '' ($r.Erreurs -join ' | ')
Test-Egal 'bloquants' '' ($r.Bloquants -join ' | ')
Test-Egal 'compteurs reportés' 'Site Users VM' ((@($r.Ecritures | ForEach-Object { $_.Cle }) | Sort-Object) -join ' ')
Test-Egal 'zéros hors contrat ignorés' 'NAS SAN ServeurPhysique UsersSansSupport VMRepliquee' ((@($r.Ignores) | Sort-Object) -join ' ')
$null = Save-SmRvReport -Report $r -Data $data
Test-Egal 'Users écrit' '46 a_verifier 2026-11-28' ((Get-Releve '2026-11' 'Users') | ForEach-Object { "$($_.Valeur) $($_.Statut) $($_.DateReleve)" })
Test-Egal 'VM daté de la confirmation' '2026-11-20' (Get-Releve '2026-11' 'VM').DateReleve
Test-Egal 'janvier intact' 'valide' (Get-Releve '2026-01' 'Users').Statut
$null = Save-SmRvReport -Report (Get-SmRvReport -Code STJAC -Mois '2026-11' -Data $data) -Data $data -Valider
Test-Egal 'validé' 'valide' (Get-Releve '2026-11' 'Users').Statut
$moisReg = @((Read-SmPvJson "$data\variabilite\releves\STJAC.json").Releves.PSObject.Properties.Name)
Test-Egal 'mois dans l ordre' (($moisReg | Sort-Object) -join ',') ($moisReg -join ',')
Test-Egal 'sauvegarde avant écriture' 2 @(Get-ChildItem "$data\variabilite\releves\sauvegardes").Count

'Conflit avec un relevé validé'
New-Calcul '2026-11' @{ '#Users' = 48; '#VM' = 4; '#Site Actuel' = 2 } '2026-11-20' '2026-11-29'
$r = Get-SmRvReport -Code STJAC -Mois '2026-11' -Data $data
Test-Egal 'conflit détecté' 'Users 46->48' (($r.Conflits | ForEach-Object { "$($_.Cle) $($_.Avant)->$($_.Valeur)" }) -join ',')
Test-Egal 'inchangés' 'Site VM' ((@($r.Inchanges) | Sort-Object) -join ' ')
$refus = try { Save-SmRvReport -Report $r -Data $data | Out-Null; '' } catch { $_.Exception.Message }
Test-Egal 'écriture refusée sans -Remplacer' $true ([bool]($refus -match 'déjà validé'))
Test-Egal 'valeur validée conservée' 46 (Get-Releve '2026-11' 'Users').Valeur

'Validation refusée'
New-Calcul '2026-12' @{ '#Users' = 46; '#VM' = 4; '#Site Actuel' = 2 } '2026-11-20' '2026-12-28' -ATrancher 2
$r = Get-SmRvReport -Code STJAC -Mois '2026-12' -Data $data
Test-Egal 'bloquants' '2 compte(s) à trancher | valeurs manuelles non confirmées pour 2026-12 (dernière : 2026-11-20)' ($r.Bloquants -join ' | ')
$refus = try { Save-SmRvReport -Report $r -Data $data -Valider | Out-Null; '' } catch { $_.Exception.Message }
Test-Egal 'validation refusée' $true ([bool]($refus -match 'Validation refusée'))
Test-Egal 'rien écrit pour 2026-12' '' (Get-Releve '2026-12' 'Users')
New-Calcul '2026-10' @{ '#Users' = 46 } '2026-10-20' '2026-09-01'
Test-Egal 'inventaire hors du mois bloquant' $true ([bool]((Get-SmRvReport -Code STJAC -Mois '2026-10' -Data $data).Bloquants -match 'hors de 2026-10'))
New-Calcul '2026-09' @{ '#Users' = 47 } '' '2026-09-01' -Ancien
Test-Egal 'calcul antérieur à 3.1 refusé' $true ([bool]((Get-SmRvReport -Code STJAC -Mois '2026-09' -Data $data).Erreurs -match 'antérieur'))

if ($echecs) { "`n$echecs échec(s). Copie : $data"; exit 1 } else { Remove-Item -Recurse -Force $data; "`nTout passe."; exit 0 }
