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

# SWPI est passé en ASW au 01.11.2026 : son contrat actuel est hors du modèle, le générateur
# prend l'ancien contrat SOM archivé (Facturation\contrats-archives, FinEffective 2026-10-31).
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

'SWPI 2026-H2 (contrat SOM jusqu''au 31.10.2026, relevés fictifs juillet-octobre)'
$dataH2 = Join-Path $sortie 'donnees-swpi-h2'
New-Item -ItemType Directory -Force -Path (Join-Path $dataH2 'Facturation\contrats'), (Join-Path $dataH2 'Facturation\contrats-archives'), (Join-Path $dataH2 'variabilite\releves') | Out-Null
Copy-Item (Join-Path $data 'Facturation\contrats\SWPI.json') (Join-Path $dataH2 'Facturation\contrats\SWPI.json')
Copy-Item (Join-Path $data 'Facturation\contrats-archives\SWPI-202411-7754.json') (Join-Path $dataH2 'Facturation\contrats-archives\SWPI-202411-7754.json')
$fictifs = [ordered]@{}
foreach ($m in @{ '2026-07' = 15; '2026-08' = 15; '2026-09' = 15; '2026-10' = 16 }.GetEnumerator()) {
    $fictifs[$m.Key] = [ordered]@{ Users = [ordered]@{ Valeur = $m.Value; Source = 'test'; DateReleve = "$($m.Key)-28"; Statut = 'valide' }
                                   Site  = [ordered]@{ Valeur = 1; Source = 'test'; DateReleve = "$($m.Key)-28"; Statut = 'valide' } }
}
# novembre : un relevé existe (sous ASW) mais doit être ignoré
$fictifs['2026-11'] = [ordered]@{ Users = [ordered]@{ Valeur = 99; Source = 'test'; DateReleve = '2026-11-28'; Statut = 'valide' } }
[IO.File]::WriteAllText((Join-Path $dataH2 'variabilite\releves\SWPI.json'), (([ordered]@{ Code = 'SWPI'; Releves = $fictifs }) | ConvertTo-Json -Depth 6), (New-Object Text.UTF8Encoding($false)))
$s2 = Get-SmPvDecompte -Code 'SWPI' -Semestre '2026-H2' -Data $dataH2 -Tarifs $tarifs
$attendu = [math]::Round(1 * [double]$tarifs['Users']['GOLD'], 2)
Test-Egal 'erreurs' '' ($s2.Erreurs -join ' | ')
Test-Egal 'contrat archivé utilisé' 'SWPI-202411-7754.json' $s2.ContratArchive
Test-Egal 'mois neutres' '2026-11,2026-12' ($s2.MoisHorsContrat -join ',')
Test-Egal 'montant Users (1 utilisateur-mois)' $attendu (($s2.Lignes | Where-Object { $_.Cle -eq 'Users' }).Montant)
Test-Egal 'total calculé' $attendu $s2.Total
if (-not $s2.Erreurs.Count) {
    $f = Join-Path $sortie (Get-SmPvNomFichier $s2)
    Test-Egal 'nom du fichier' 'Plus value -SWPI-202411-7754-H2-2026.xlsx' (Split-Path $f -Leaf)
    Test-Egal 'total Excel' $attendu (New-SmPvFichier -Decompte $s2 -Modele $modele -Destination $f)
}

'Refus attendus'
$z = Get-SmPvDecompte -Code 'ZENIT' -Semestre '2026-H1' -Data $data -Tarifs $tarifs
Test-Egal 'ZENIT refusé (base absente)' $true ([bool]($z.Erreurs -match 'Base du contrat inconnue'))
$g = Get-SmPvDecompte -Code 'GIPPA' -Semestre '2026-H1' -Data $data -Tarifs $tarifs
Test-Egal 'GIPPA refusé (BWS)' $true ([bool]($g.Erreurs -match 'hors du modèle'))
$w = Get-SmPvDecompte -Code 'SWPI' -Semestre '2027-H1' -Data $data -Tarifs $tarifs
Test-Egal 'SWPI 2027-H1 refusé (ASW, aucun contrat archivé ne couvre le semestre)' $true ([bool]($w.Erreurs -match 'hors du modèle'))
$h2 = Get-SmPvDecompte -Code 'STJAC' -Semestre '2026-H2' -Data $data -Tarifs $tarifs
Test-Egal 'STJAC 2026-H2 refusé (relevés manquants)' $true ([bool]($h2.Erreurs -match 'relevé manquant'))

if ($echecs) { "`n$echecs échec(s). Fichiers : $sortie"; exit 1 } else { "`nTout passe. Fichiers : $sortie"; exit 0 }
