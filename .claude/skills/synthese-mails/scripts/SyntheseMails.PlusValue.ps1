<#
.SYNOPSIS
    Plus-value semestrielle : décompte au format du modèle de M. Vaamonde, à partir
    des relevés mensuels validés et de la base du contrat.

.DESCRIPTION
    Sources, toutes dans <dossierAnsamDiag> :
      Facturation\contrats\<CODE>.json   Contrat.Reference, Contrat.Modele (Silver, Gold,
                                         Platinum) et PerimetreContrat (base du contrat)
      variabilite\releves\<CODE>.json    relevés mensuels : Releves.<AAAA-MM>.<Cle> =
                                         { Valeur, Source, DateReleve, Statut }
      variabilite\modeles\modele-plus-value-*.xlsx   modèle (le plus récent)

    Le décompte est calculé deux fois : ici, et par Excel dans le fichier produit. Les
    deux totaux doivent être égaux au centime, sinon rien n'est gardé.

    Règles :
      - une ligne est suivie si la base du contrat la mentionne avec une valeur > 0 ou
        si un relevé existe pour elle ; elle exige alors un relevé pour chacun des
        6 mois, sinon le décompte est refusé ;
      - une ligne non suivie vaut 0 (base 0, relevés 0) ;
      - un relevé « a_verifier » n'empêche pas le décompte mais le rend provisoire ;
      - un contrat BWS ou de type inconnu n'entre pas dans ce modèle : refusé ;
      - si le contrat actuel est hors du modèle (ASW, BWS…), un contrat du modèle archivé
        dans Facturation\contrats-archives\<CODE>-*.json et couvrant une partie du semestre
        est utilisé à sa place (cas SWPI : SOM jusqu'au 31.10.2026, puis ASW) ;
      - les mois hors de Contrat.DebutEffectif / Contrat.FinEffective (champs explicites,
        posés seulement quand un contrat commence ou finit en cours de semestre) sont
        neutres : valeur = base, écart nul, aucun relevé exigé. Le modèle Excel garde ses
        6 colonnes.
#>

$SmPvLignes = @(
    [ordered]@{ Ligne = 10; Cle = 'NAS';              Base = 'NAS';                     Libelle = 'NAS';                          Tarif = 8 }
    [ordered]@{ Ligne = 11; Cle = 'SAN';              Base = 'SAN';                     Libelle = 'SAN';                          Tarif = 9 }
    [ordered]@{ Ligne = 12; Cle = 'ServeurPhysique';  Base = 'ServeursPhysiques';       Libelle = 'Serveur physique';             Tarif = 10 }
    [ordered]@{ Ligne = 13; Cle = 'VM';               Base = 'ServeursVirtuels';        Libelle = 'Serveur virtuel';              Tarif = 11 }
    [ordered]@{ Ligne = 14; Cle = 'Site';             Base = 'Sites';                   Libelle = 'Site';                         Tarif = 13 }
    [ordered]@{ Ligne = 15; Cle = 'Users';            Base = 'Utilisateurs';            Libelle = 'Utilisateur incl. support L1'; Tarif = 14 }
    [ordered]@{ Ligne = 16; Cle = 'UsersSansSupport'; Base = 'UtilisateursSansSupport'; Libelle = 'Utilisateur sans support L1';  Tarif = 15 }
)
$SmPvColonneTarif = @{ SILVER = 2; GOLD = 3; PLATINUM = 4 }
$SmPvMoisCourts = @('Janv.', 'Févr.', 'Mars', 'Avr.', 'Mai', 'Juin', 'Juil.', 'Août', 'Sept.', 'Oct.', 'Nov.', 'Déc.')
$SmPvMoisLongs = @('janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre')

function Get-SmPvMois([string]$Semestre) {
    if ($Semestre -notmatch '^(\d{4})-H([12])$') { throw "Semestre attendu au format AAAA-H1 ou AAAA-H2 : $Semestre" }
    $an = [int]$Matches[1]; $debut = if ($Matches[2] -eq '1') { 1 } else { 7 }
    @($debut..($debut + 5) | ForEach-Object { '{0}-{1:D2}' -f $an, $_ })
}

function Read-SmPvJson([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8) | ConvertFrom-Json
}

# Date AAAA-MM-JJ (ou AAAA-MM) -> AAAA-MM ; toute autre valeur (vide, « reconduction… ») -> $null.
function Get-SmPvMoisDate($Valeur) {
    $s = [string]$Valeur
    if ($s -match '^(\d{4}-\d{2})(-\d{2})?') { $Matches[1] } else { $null }
}

# Contrat à appliquer au semestre : le contrat actuel s'il est du modèle (Silver, Gold,
# Platinum) ; sinon le plus récent contrat archivé du modèle qui couvre au moins un mois.
function Get-SmPvContrat([string]$Data, [string]$Code, [string[]]$Mois) {
    $courant = Read-SmPvJson (Join-Path $Data "Facturation\contrats\$Code.json")
    if (-not $courant -or $SmPvColonneTarif.ContainsKey(([string]$courant.Contrat.Modele).ToUpper())) {
        return [ordered]@{ Contrat = $courant; Archive = $null }
    }
    $dir = Join-Path $Data 'Facturation\contrats-archives'
    foreach ($f in @(Get-ChildItem -LiteralPath $dir -Filter "$Code-*.json" -File -ErrorAction SilentlyContinue | Sort-Object Name -Descending)) {
        $c = Read-SmPvJson $f.FullName
        if (-not $c -or -not $SmPvColonneTarif.ContainsKey(([string]$c.Contrat.Modele).ToUpper())) { continue }
        $d = Get-SmPvMoisDate $(if ($c.Contrat.DebutEffectif) { $c.Contrat.DebutEffectif } else { $c.Contrat.Debut })
        $fi = Get-SmPvMoisDate $(if ($c.Contrat.FinEffective) { $c.Contrat.FinEffective } else { $c.Contrat.Fin })
        if (@($Mois | Where-Object { (-not $d -or $_ -ge $d) -and (-not $fi -or $_ -le $fi) }).Count) {
            return [ordered]@{ Contrat = $c; Archive = $f.Name }
        }
    }
    [ordered]@{ Contrat = $courant; Archive = $null }
}

function Get-SmPvModeleChemin([string]$Data) {
    $f = Get-ChildItem -LiteralPath (Join-Path $Data 'variabilite\modeles') -Filter 'modele-plus-value-*.xlsx' -File -ErrorAction SilentlyContinue |
        Sort-Object Name -Descending | Select-Object -First 1
    if (-not $f) { throw "Aucun modèle dans $(Join-Path $Data 'variabilite\modeles')" }
    $f.FullName
}

# Tarifs lus dans le modèle lui-même (Feuil2), jamais recopiés dans le code.
function Get-SmPvTarifs([string]$Modele) {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false; $xl.DisplayAlerts = $false
    try {
        $wb = $xl.Workbooks.Open($Modele, 0, $true)
        $ws = $wb.Worksheets.Item('Feuil2')
        $t = @{}
        foreach ($l in $SmPvLignes) {
            $t[$l.Cle] = @{}
            foreach ($k in $SmPvColonneTarif.Keys) { $t[$l.Cle][$k] = [double]$ws.Cells.Item([int]$l.Tarif, [int]$SmPvColonneTarif[$k]).Value2 }
        }
        $wb.Close($false)
        $t
    } finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl); [GC]::Collect() }
}

<#
    Rassemble et contrôle les données d'un client pour un semestre. Ne touche à aucun
    fichier. Renvoie { Code, Semestre, Mois, Reference, Nom, Type, Lignes, Total,
    Erreurs, Avertissements, Provisoire }.
#>
function Get-SmPvDecompte {
    param([Parameter(Mandatory)][string]$Code, [Parameter(Mandatory)][string]$Semestre,
          [Parameter(Mandatory)][string]$Data, [Parameter(Mandatory)][hashtable]$Tarifs)
    $Code = $Code.ToUpper()
    $mois = Get-SmPvMois $Semestre
    $r = [ordered]@{ Code = $Code; Semestre = $Semestre; Mois = $mois; Reference = $null; Nom = $null; Type = $null
                     Lignes = @(); Total = $null; Erreurs = @(); Avertissements = @(); Provisoire = $false
                     ContratArchive = $null; MoisHorsContrat = @() }
    $choix = Get-SmPvContrat $Data $Code $mois
    $contrat = $choix.Contrat
    if (-not $contrat) { $r.Erreurs += "Fiche de contrat absente : Facturation\contrats\$Code.json"; return $r }
    if ($choix.Archive) {
        $r.ContratArchive = $choix.Archive
        $r.Avertissements += "Contrat archivé utilisé : Facturation\contrats-archives\$($choix.Archive) (le contrat actuel est hors du modèle)"
    }
    $dEff = Get-SmPvMoisDate $contrat.Contrat.DebutEffectif
    $fEff = Get-SmPvMoisDate $contrat.Contrat.FinEffective
    $r.MoisHorsContrat = @($mois | Where-Object { ($dEff -and $_ -lt $dEff) -or ($fEff -and $_ -gt $fEff) })
    $requis = @($mois | Where-Object { $r.MoisHorsContrat -notcontains $_ })
    if (-not $requis.Count) { $r.Erreurs += "Aucun mois du semestre $Semestre n'est couvert par le contrat"; return $r }
    if ($r.MoisHorsContrat.Count) { $r.Avertissements += "Mois hors contrat, neutres (valeur = base) : $($r.MoisHorsContrat -join ', ')" }
    $r.Nom = [string]$contrat.Nom
    $ref = [string]$contrat.Contrat.Reference
    if ($ref -match '^([A-Z0-9]+-\d{6}-\d+)') { $r.Reference = $Matches[1] } else { $r.Erreurs += "Référence de contrat illisible : '$ref'" }
    $type = ([string]$contrat.Contrat.Modele).ToUpper()
    if (-not $SmPvColonneTarif.ContainsKey($type)) { $r.Erreurs += "Type de contrat « $($contrat.Contrat.Modele) » hors du modèle (Silver, Gold ou Platinum seulement)" } else { $r.Type = $type }
    $base = $contrat.PerimetreContrat
    if (-not $base) { $r.Erreurs += 'Base du contrat inconnue (PerimetreContrat absent de la fiche)' }
    $releves = Read-SmPvJson (Join-Path $Data "variabilite\releves\$Code.json")
    if (-not $releves) { $r.Erreurs += "Aucun relevé mensuel : variabilite\releves\$Code.json" }
    if ($r.Erreurs.Count) { return $r }

    $total = 0.0
    foreach ($l in $SmPvLignes) {
        $p = $base.PSObject.Properties[$l.Base]
        $vBase = if ($p -and $null -ne $p.Value) { [double]$p.Value } else { $null }
        $vals = @(); $sources = @(); $presents = 0
        foreach ($m in $mois) {
            if ($r.MoisHorsContrat -contains $m) { $vals += $(if ($null -ne $vBase) { $vBase } else { 0.0 }); continue }
            $rm = $releves.Releves.PSObject.Properties[$m]
            $x = if ($rm) { $rm.Value.PSObject.Properties[$l.Cle] } else { $null }
            if ($x) {
                $presents++
                $vals += [double]$x.Value.Valeur
                $sources += [ordered]@{ Mois = $m; Valeur = $x.Value.Valeur; Statut = $x.Value.Statut; Source = $x.Value.Source; DateReleve = $x.Value.DateReleve }
                if ($x.Value.Statut -ne 'valide') { $r.Avertissements += "$($l.Libelle), $m : relevé « $($x.Value.Statut) » ($($x.Value.Source))"; $r.Provisoire = $true }
            } else {
                $vals += $null
                if ($rm -and $rm.Value.PSObject.Properties['UsersTotal'] -and $l.Cle -in 'Users', 'UsersSansSupport') {
                    $r.Erreurs += "$($l.Libelle), $m : seul un total utilisateurs existe (sans séparation avec / sans support)"
                }
            }
        }
        $suivie = ($presents -gt 0) -or ($vBase -gt 0)
        if ($suivie -and $null -eq $vBase) { $r.Erreurs += "$($l.Libelle) : relevés présents mais base du contrat absente ($($l.Base))"; continue }
        if ($suivie -and $presents -lt $requis.Count) {
            $manque = @(for ($i = 0; $i -lt 6; $i++) { if ($null -eq $vals[$i]) { $mois[$i] } })
            $r.Erreurs += "$($l.Libelle) : relevé manquant pour $($manque -join ', ')"; continue
        }
        if (-not $suivie) { $vBase = 0; $vals = @(0, 0, 0, 0, 0, 0) }
        $tarif = [double]$Tarifs[$l.Cle][$r.Type]
        $ecart = 0.0; foreach ($v in $vals) { $ecart += $v }; $ecart -= 6 * $vBase
        $montant = [math]::Round($ecart * $tarif, 2)
        $total += $montant
        $r.Lignes += [ordered]@{ Cle = $l.Cle; Libelle = $l.Libelle; Ligne = $l.Ligne; Suivie = $suivie; Base = $vBase
                                 Valeurs = $vals; EcartMoisUnites = $ecart; Tarif = $tarif; Montant = $montant; Sources = $sources }
    }
    $r.Total = [math]::Round($total, 2)
    $r
}

function Get-SmPvNomFichier($Decompte) {
    $s = $Decompte.Semestre -split '-'
    "Plus value -$($Decompte.Reference)-$($s[1])-$($s[0]).xlsx"
}

<#
    Remplit une copie du modèle avec un décompte sans erreur, laisse Excel recalculer
    et vérifie que son total égale celui du décompte. Le fichier n'est gardé que si
    les deux concordent. -Remplacer met l'ancien fichier de côté (sauvegardes\).
#>
function New-SmPvFichier {
    param([Parameter(Mandatory)]$Decompte, [Parameter(Mandatory)][string]$Modele,
          [Parameter(Mandatory)][string]$Destination, [switch]$Remplacer)
    if ($Decompte.Erreurs.Count) { throw "Décompte $($Decompte.Code) refusé : $($Decompte.Erreurs -join ' ; ')" }
    $dir = Split-Path $Destination -Parent
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    if (Test-Path -LiteralPath $Destination) {
        if (-not $Remplacer) { throw "Le fichier existe déjà : $Destination" }
        $sv = Join-Path $dir 'sauvegardes'; New-Item -ItemType Directory -Force -Path $sv | Out-Null
        Move-Item -LiteralPath $Destination -Destination (Join-Path $sv ("{0}-{1}.xlsx" -f [IO.Path]::GetFileNameWithoutExtension($Destination), (Get-Date).ToString('yyyyMMdd-HHmmss')))
    }
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ("pv-{0}.xlsx" -f [guid]::NewGuid().ToString('N'))
    Copy-Item -LiteralPath $Modele -Destination $tmp
    $mois = $Decompte.Mois
    $an = $mois[0].Substring(0, 4)
    $idx = @($mois | ForEach-Object { [int]$_.Substring(5, 2) - 1 })
    $lettres = 'C', 'D', 'E', 'F', 'G', 'H'
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false; $xl.DisplayAlerts = $false
    $totalExcel = $null
    try {
        $wb = $xl.Workbooks.Open($tmp)
        $ws = $wb.Worksheets.Item('Feuil1')
        foreach ($l in $SmPvLignes) {
            $lu = [string]$ws.Range("A$($l.Ligne)").Value2
            if ($lu.Trim() -ne $l.Libelle) { throw "Modèle inattendu : A$($l.Ligne) = « $lu », attendu « $($l.Libelle) »" }
        }
        $ws.Range('A2').Value2 = "$($Decompte.Nom) - Evolution du contrat"
        $ws.Range('B3').Value2 = [string]$Decompte.Reference
        $ws.Range('A7').Value2 = "Période : {0} {1} à {2} {1}" -f ((Get-Culture).TextInfo.ToTitleCase($SmPvMoisLongs[$idx[0]])), $an, $SmPvMoisLongs[$idx[5]]
        for ($i = 0; $i -lt 6; $i++) { $ws.Range("$($lettres[$i])9").Value2 = [string]("{0} {1}" -f $SmPvMoisCourts[$idx[$i]], $an) }
        foreach ($l in $Decompte.Lignes) {
            $ws.Range("B$($l.Ligne)").Value2 = [double]$l.Base
            for ($i = 0; $i -lt 6; $i++) { $ws.Range("$($lettres[$i])$($l.Ligne)").Value2 = [double]$l.Valeurs[$i] }
        }
        $ws.Range('A19').Value2 = "Surcoût par rapport au contrat sur le {0} semestre {1} :" -f $(if ($idx[0] -eq 0) { 'premier' } else { 'second' }), $an
        $ws.Range('B24').Value2 = [string]$Decompte.Type
        $xl.CalculateFull()
        $totalExcel = [math]::Round([double]$ws.Range('F19').Value2, 2)
        foreach ($l in $Decompte.Lignes) {
            $j = [math]::Round([double]$ws.Range("J$($l.Ligne)").Value2, 2)
            if ($j -ne $l.Montant) { throw "Écart Excel ligne $($l.Libelle) : Excel $j, calcul $($l.Montant)" }
        }
        if ($totalExcel -ne $Decompte.Total) { throw "Écart de total : Excel $totalExcel, calcul $($Decompte.Total)" }
        $wb.Save(); $wb.Close($true)
    } catch {
        if ($wb) { try { $wb.Close($false) } catch { } }
        Remove-Item -LiteralPath $tmp -ErrorAction SilentlyContinue
        throw
    } finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl); [GC]::Collect() }
    Move-Item -LiteralPath $tmp -Destination $Destination
    $totalExcel
}
