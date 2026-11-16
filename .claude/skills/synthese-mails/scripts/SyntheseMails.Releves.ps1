<#
.SYNOPSIS
    Report d'un calcul mensuel de variabilité (outil AnsamDiag Variabilite.Metriques,
    v3.1 ou plus) dans le registre des relevés, base de la plus-value semestrielle.

.DESCRIPTION
    Lit   <dossierAnsamDiag>\variabilite\_etat\<CODE>\Mois-<AAAA-MM>.json
    Écrit <dossierAnsamDiag>\variabilite\releves\<CODE>.json
          Releves.<AAAA-MM>.<Cle> = { Valeur, Source, DateReleve, Statut }

    Règles :
      - un compteur est reporté s'il est non nul, si la base du contrat le mentionne
        (> 0) ou s'il a déjà des relevés pour ce client ; sinon il resterait un 0 qui
        ferait « suivre » une ligne absente du contrat ;
      - la date d'un relevé est celle de l'inventaire (M365, AD) ou de la dernière
        confirmation des valeurs manuelles ;
      - statut « a_verifier » par défaut ; « valide » seulement avec -Valider, refusé
        tant qu'un compte reste à trancher, que les valeurs manuelles ne sont pas
        confirmées pour le mois ou qu'un inventaire n'est pas daté du mois ;
      - un relevé déjà « valide » de valeur différente n'est jamais remplacé sans
        -Remplacer : c'est un conflit, signalé ;
      - le registre est sauvegardé avant chaque écriture (releves\sauvegardes\).

    Nécessite SyntheseMails.PlusValue.ps1 (correspondance compteur ↔ base du contrat).
#>

$SmRvMetriques = [ordered]@{
    '#Users' = 'Users'; '#USERnoSUPPORT' = 'UsersSansSupport'; '#Serveur Physique' = 'ServeurPhysique'; '#VM' = 'VM'
    '#VM Repliquee' = 'VMRepliquee'; '#SAN' = 'SAN'; '#NAS' = 'NAS'; '#Site Actuel' = 'Site'
}

# Objet JSON (PowerShell 5.1, sans -AsHashtable) → tables ordonnées modifiables.
function ConvertTo-SmRvTable($o) {
    if ($null -eq $o) { return $null }
    if ($o -is [System.Management.Automation.PSCustomObject]) {
        $t = [ordered]@{}
        foreach ($p in $o.PSObject.Properties) { $t[$p.Name] = ConvertTo-SmRvTable $p.Value }
        return $t
    }
    if ($o -is [System.Collections.IList] -and $o -isnot [string]) { return , @($o | ForEach-Object { ConvertTo-SmRvTable $_ }) }
    $o
}

function Get-SmRvReport {
    param([Parameter(Mandatory)][string]$Code, [Parameter(Mandatory)][string]$Mois, [Parameter(Mandatory)][string]$Data)
    $Code = $Code.ToUpper()
    $r = [ordered]@{ Code = $Code; Mois = $Mois; Erreurs = @(); Bloquants = @(); Ecritures = @(); Inchanges = @(); Conflits = @(); Ignores = @() }
    if ($Mois -notmatch '^\d{4}-(0[1-9]|1[0-2])$') { $r.Erreurs += "Mois attendu au format AAAA-MM : $Mois"; return $r }
    $calc = Read-SmPvJson (Join-Path $Data "variabilite\_etat\$Code\Mois-$Mois.json")
    if (-not $calc) { $r.Erreurs += "Aucun calcul de variabilité pour $Mois (lancer Variabilite.Metriques)"; return $r }
    if (-not $calc.PSObject.Properties['Inventaires']) { $r.Erreurs += "Calcul de $Mois antérieur à Variabilite.Metriques 3.1 (date des inventaires inconnue) : le relancer"; return $r }

    # Conditions de validation
    $nTrancher = @($calc.ATrancher | Where-Object { $_ }).Count
    if ($nTrancher) { $r.Bloquants += "$nTrancher compte(s) à trancher" }
    $conf = [string]$calc.ValeursManuellesVerifiees
    if (-not $conf -or $conf.Substring(0, [math]::Min(7, $conf.Length)) -lt $Mois) { $r.Bloquants += "valeurs manuelles non confirmées pour $Mois (dernière : $(if ($conf) { $conf } else { 'aucune' }))" }
    foreach ($t in 'M365', 'AD') {
        $d = [string]$calc.Inventaires.$t
        if ($d -and -not $d.StartsWith($Mois)) { $r.Bloquants += "inventaire $t daté du $d, hors de $Mois" }
    }

    $contrat = Read-SmPvJson (Join-Path $Data "Facturation\contrats\$Code.json")
    $base = if ($contrat) { $contrat.PerimetreContrat } else { $null }
    $reg = Read-SmPvJson (Join-Path $Data "variabilite\releves\$Code.json")
    $connues = @{}
    if ($reg -and $reg.Releves) { foreach ($m in $reg.Releves.PSObject.Properties) { foreach ($k in $m.Value.PSObject.Properties) { $connues[$k.Name] = $true } } }
    $existant = if ($reg -and $reg.Releves -and $reg.Releves.PSObject.Properties[$Mois]) { $reg.Releves.$Mois } else { $null }

    foreach ($row in @($calc.Tableau)) {
        $cle = $SmRvMetriques[[string]$row.Metrique]
        if (-not $cle) { continue }
        $val = [int]$row.MoisCourant
        $nomBase = ($SmPvLignes | Where-Object { $_.Cle -eq $cle } | Select-Object -First 1).Base
        $vBase = if ($base -and $nomBase -and $base.PSObject.Properties[$nomBase]) { [double]$base.$nomBase } else { 0 }
        if ($val -eq 0 -and $vBase -le 0 -and -not $connues[$cle]) { $r.Ignores += $cle; continue }
        $src = [string]$row.Source
        $date = if ($src -like 'Manuel*') { $conf } elseif ($src -like 'AD*') { [string]$calc.Inventaires.AD } else { [string]$calc.Inventaires.M365 }
        $e = [ordered]@{ Cle = $cle; Valeur = $val; Source = "Variabilite.Metriques, calcul du $($calc.Calcule) : $src"; DateReleve = $date }
        $avant = if ($existant -and $existant.PSObject.Properties[$cle]) { $existant.$cle } else { $null }
        if ($avant -and [string]$avant.Statut -eq 'valide') {
            if ([int]$avant.Valeur -eq $val) { $r.Inchanges += $cle; continue }
            $e.Avant = [int]$avant.Valeur; $e.SourceAvant = [string]$avant.Source
            $r.Conflits += $e; continue
        }
        if ($avant) { $e.Avant = [int]$avant.Valeur }
        $r.Ecritures += $e
    }
    $r
}

function Save-SmRvReport {
    param([Parameter(Mandatory)]$Report, [Parameter(Mandatory)][string]$Data, [switch]$Valider, [switch]$Remplacer)
    if ($Report.Erreurs.Count) { throw ($Report.Erreurs -join ' | ') }
    if ($Valider -and $Report.Bloquants.Count) { throw "Validation refusée : $($Report.Bloquants -join ' ; ')" }
    if ($Report.Conflits.Count -and -not $Remplacer) {
        throw ("Relevé(s) déjà validé(s) avec une autre valeur : " + (($Report.Conflits | ForEach-Object { "$($_.Cle) $($_.Avant) -> $($_.Valeur)" }) -join ', ') + ' (-Remplacer pour les remplacer)')
    }
    $aEcrire = @($Report.Ecritures) + $(if ($Remplacer) { @($Report.Conflits) } else { @() })
    $statut = if ($Valider) { 'valide' } else { 'a_verifier' }

    $dir = Join-Path $Data 'variabilite\releves'
    $f = Join-Path $dir "$($Report.Code).json"
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    $reg = ConvertTo-SmRvTable (Read-SmPvJson $f)
    if (-not $reg) { $reg = [ordered]@{ Code = $Report.Code; MisAJour = $null; Releves = [ordered]@{} } }
    if (-not $reg.Releves) { $reg.Releves = [ordered]@{} }
    if (Test-Path -LiteralPath $f) {
        $sv = Join-Path $dir 'sauvegardes'
        New-Item -ItemType Directory -Force -Path $sv | Out-Null
        Copy-Item -LiteralPath $f -Destination (Join-Path $sv ("{0}-{1}.json" -f $Report.Code, (Get-Date).ToString('yyyyMMdd-HHmmss-fff'))) -Force
    }
    if (-not $reg.Releves.Contains($Report.Mois)) {
        # Mois rangés dans l'ordre chronologique.
        $tri = [ordered]@{}
        foreach ($m in (@($reg.Releves.Keys) + $Report.Mois | Sort-Object)) { $tri[$m] = if ($m -eq $Report.Mois) { [ordered]@{} } else { $reg.Releves[$m] } }
        $reg.Releves = $tri
    }
    $mois = $reg.Releves[$Report.Mois]
    foreach ($e in $aEcrire) {
        $mois[$e.Cle] = [ordered]@{ Valeur = $e.Valeur; Source = $e.Source; DateReleve = $e.DateReleve; Statut = $statut }
    }
    $reg.MisAJour = (Get-Date).ToString('s')
    [IO.File]::WriteAllText($f, (ConvertTo-Json -InputObject $reg -Depth 10), (New-Object Text.UTF8Encoding($false)))
    [ordered]@{ Fichier = $f; Statut = $statut; Ecrits = @($aEcrire | ForEach-Object { $_.Cle }) }
}
