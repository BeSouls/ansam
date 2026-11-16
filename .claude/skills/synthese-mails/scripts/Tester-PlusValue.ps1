<#
.SYNOPSIS
    Tests de référence de la plus-value : le 1er semestre 2026, contrôlé et envoyé à
    Finances le 16.11.2026, doit être reproduit au centime. À relancer après toute
    modification du générateur, du modèle ou des règles.

    Écrit seulement dans un dossier temporaire. Code de sortie 0 si tout passe.
#>
[CmdletBinding()]
param([string]$WorkDir = (Join-Path $env:USERPROFILE 'SyntheseMails'))
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'SyntheseMails.Commun.ps1')
. (Join-Path $PSScriptRoot 'SyntheseMails.PlusValue.ps1')
$data = [string](Get-SmConfig $WorkDir).dossierAnsamDiag
$modele = Get-SmPvModeleChemin $data
$tarifs = Get-SmPvTarifs $modele
$sortie = Join-Path ([IO.Path]::GetTempPath()) ("test-plus-value-{0}" -f (Get-Date).ToString('yyyyMMdd-HHmmss'))
$echecs = 0
function Test-Egal($Nom, $Attendu, $Obtenu) {
    if ("$Attendu" -eq "$Obtenu") { "  ok    $Nom = $Obtenu" } else { "  ECHEC $Nom : attendu $Attendu, obtenu $Obtenu"; $script:echecs++ }
}

$references = @(
    @{ Code = 'STJAC'; Total = 1369;    Fichier = 'Plus value -STJAC-202411-7752-H1-2026.xlsx'; Lignes = @{ Users = 589; VM = 780; Site = 0 } }
    @{ Code = 'SWPI';  Total = -93;     Fichier = 'Plus value -SWPI-202411-7754-H1-2026.xlsx';  Lignes = @{ Users = -93; Site = 0 } }
    @{ Code = 'PREST'; Total = -219.78; Fichier = 'Plus value -PREST-202106-2176-H1-2026.xlsx'; Lignes = @{ VM = 0; UsersSansSupport = -219.78; ServeurPhysique = 0; Site = 0 } }
)
foreach ($ref in $references) {
    "$($ref.Code) 2026-H1"
    $d = Get-SmPvDecompte -Code $ref.Code -Semestre '2026-H1' -Data $data -Tarifs $tarifs
    Test-Egal 'erreurs' '' ($d.Erreurs -join ' | ')
    Test-Egal 'total calculé' $ref.Total $d.Total
    foreach ($k in $ref.Lignes.Keys) { Test-Egal "montant $k" $ref.Lignes[$k] (($d.Lignes | Where-Object { $_.Cle -eq $k }).Montant) }
    if (-not $d.Erreurs.Count) {
        $f = Join-Path $sortie (Get-SmPvNomFichier $d)
        Test-Egal 'nom du fichier' $ref.Fichier (Split-Path $f -Leaf)
        $te = New-SmPvFichier -Decompte $d -Modele $modele -Destination $f
        Test-Egal 'total Excel' $ref.Total $te
    }
}

'Refus attendus'
$z = Get-SmPvDecompte -Code 'ZENIT' -Semestre '2026-H1' -Data $data -Tarifs $tarifs
Test-Egal 'ZENIT refusé (base absente)' $true ([bool]($z.Erreurs -match 'Base du contrat inconnue'))
$g = Get-SmPvDecompte -Code 'GIPPA' -Semestre '2026-H1' -Data $data -Tarifs $tarifs
Test-Egal 'GIPPA refusé (BWS)' $true ([bool]($g.Erreurs -match 'hors du modèle'))
$h2 = Get-SmPvDecompte -Code 'STJAC' -Semestre '2026-H2' -Data $data -Tarifs $tarifs
Test-Egal 'STJAC 2026-H2 refusé (relevés manquants)' $true ([bool]($h2.Erreurs -match 'relevé manquant'))

if ($echecs) { "`n$echecs échec(s). Fichiers : $sortie"; exit 1 } else { "`nTout passe. Fichiers : $sortie"; exit 0 }
