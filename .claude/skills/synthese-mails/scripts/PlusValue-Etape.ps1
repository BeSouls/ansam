<#
.SYNOPSIS
    Plus-value semestrielle lancée depuis Panoramix (ou à la main) : décompte au
    format du modèle de M. Vaamonde pour un ou plusieurs clients.

.DESCRIPTION
    controle   Calcule les décomptes sans rien écrire, sauf le résultat :
               <dossierAnsamDiag>\variabilite\plus-value\<AAAA-Hn>.json
    generer    Idem, puis produit pour chaque client sans erreur le fichier
               « Plus value -<contrat>-Hn-AAAA.xlsx » dans <dossierClients>\<CODE>\plus value\
               (ou dans -Sortie), après double calcul (ici et par Excel).

    Rien n'est envoyé : le TAM relit, puis transmet à Finances (Mario en copie).
    L'avancement est dans <WorkDir>\travaux\plusvalue-<etape>.json.

.EXAMPLE
    .\PlusValue-Etape.ps1 -Etape controle -Semestre 2026-H2 -Tous
    .\PlusValue-Etape.ps1 -Etape generer -Semestre 2026-H1 -Client STJAC
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [ValidateSet('controle', 'generer')] [string]$Etape,
    [Parameter(Mandatory)] [ValidatePattern('^\d{4}-H[12]$')] [string]$Semestre,
    [string[]]$Client,
    [switch]$Tous,
    [string]$Sortie,
    [switch]$Remplacer,
    [string]$WorkDir = (Join-Path $env:USERPROFILE 'SyntheseMails')
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'SyntheseMails.Commun.ps1')
. (Join-Path $PSScriptRoot 'SyntheseMails.PlusValue.ps1')
$config = Get-SmConfig $WorkDir
$data = [string]$config.dossierAnsamDiag
$clients = [string]$config.dossierClients
$statusFile = Join-Path $WorkDir "travaux\plusvalue-$Etape.json"
function Set-Status([string]$State, [string]$Message) {
    Write-SmJson $statusFile ([ordered]@{ etape = $Etape; semestre = $Semestre; etat = $State; message = $Message; le = (Get-Date).ToString('o'); pid = $PID })
}
Set-Status 'en_cours' 'Démarrage'

try {
    $codes = if ($Tous) {
        @(Get-ChildItem -LiteralPath (Join-Path $data 'Facturation\contrats') -Filter '*.json' -File | ForEach-Object { $_.BaseName })
    } elseif ($Client) { @($Client | ForEach-Object { $_.ToUpper() }) } else { throw 'Indiquer -Client ou -Tous.' }

    $modele = Get-SmPvModeleChemin $data
    $tarifs = Get-SmPvTarifs $modele
    $resultats = @()
    foreach ($c in $codes) {
        Set-Status 'en_cours' "Décompte $c"
        $d = Get-SmPvDecompte -Code $c -Semestre $Semestre -Data $data -Tarifs $tarifs
        $d.Fichier = $null; $d.TotalExcel = $null
        if ($Etape -eq 'generer' -and -not $d.Erreurs.Count) {
            $dest = if ($Sortie) { Join-Path $Sortie (Get-SmPvNomFichier $d) } else { Join-Path $clients "$c\plus value\$(Get-SmPvNomFichier $d)" }
            try {
                $d.TotalExcel = New-SmPvFichier -Decompte $d -Modele $modele -Destination $dest -Remplacer:$Remplacer
                $d.Fichier = $dest
            } catch { $d.Erreurs += "Génération : $($_.Exception.Message)" }
        }
        $resultats += [pscustomobject]$d
    }

    $dirOut = if ($Sortie) { $Sortie } else { Join-Path $data 'variabilite\plus-value' }
    New-Item -ItemType Directory -Force -Path $dirOut | Out-Null
    $out = [ordered]@{ semestre = $Semestre; etape = $Etape; le = (Get-Date).ToString('o'); modele = $modele; clients = $resultats }
    [IO.File]::WriteAllText((Join-Path $dirOut "$Semestre.json"), (ConvertTo-Json -InputObject $out -Depth 10), (New-Object Text.UTF8Encoding($false)))

    $ok = @($resultats | Where-Object { -not $_.Erreurs.Count }).Count
    Set-Status 'ok' ("{0} client(s) sans erreur sur {1}" -f $ok, $resultats.Count)
    $resultats | ForEach-Object {
        $etat = if ($_.Erreurs.Count) { 'REFUSE' } elseif ($_.Provisoire) { 'PROVISOIRE' } else { 'OK' }
        "{0,-6} {1,-10} total {2,10}  {3}" -f $_.Code, $etat, $_.Total, $(if ($_.Erreurs.Count) { $_.Erreurs -join ' | ' } else { $_.Fichier })
    }
} catch {
    Set-Status 'echec' $_.Exception.Message
    throw
}
