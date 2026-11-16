<#
.SYNOPSIS
    Onglet « Mois » de Panoramix : état du cycle mensuel (facturation, variabilité,
    plus-value), actions du TAM et rappels de la synthèse du matin.

.DESCRIPTION
    Lit, sans rien calculer lui-même :
      <dossierAnsamDiag>\variabilite\perimetre.json         clients suivis
      <dossierClients>\<CODE>.json                           fiche (sources, valeurs manuelles)
      <dossierClients>\<CODE>\Inventaires\M365|AD\<date>\    inventaires (et l'ancien dossier de la fiche)
      <dossierAnsamDiag>\variabilite\_etat\<CODE>\Mois-<AAAA-MM>.json   calcul Variabilite.Metriques
      <dossierAnsamDiag>\variabilite\releves\<CODE>.json     registre des relevés
      <dossierAnsamDiag>\variabilite\plus-value\<AAAA-Hn>.json
      <dossierAnsamDiag>\Facturation\<AAAA-MM>\              état de la facturation
      <WorkDir>\mois\<AAAA-MM>.json                          ce que le TAM a marqué : Kiosk saisi,
                                                             Jira commenté, comptes tranchés

    Écrit seulement : les marques du TAM, et dans la fiche du client les valeurs
    manuelles confirmées ou un compte exclu (fiche sauvegardée avant).
    Rien n'est saisi dans Kiosk ni écrit dans Jira.
#>

$SmMoisClientsDefaut = @('STJAC', 'SWPI', 'PREST', 'SOLLE', 'VETBE', 'ZENIT', 'SLV', 'AVE', 'GIPPA', 'SBVME')
$SmMoisManuels = @('VM', 'ServeurPhysique', 'SiteActuel', 'NAS', 'SAN', 'VMRepliquee', 'USERnoSUPPORT')

function Assert-SmMois([string]$Mois) { if ($Mois -notmatch '^\d{4}-(0[1-9]|1[0-2])$') { throw "Mois invalide : $Mois" } }
function Get-SmMoisDecale([string]$Mois, [int]$Delta) {
    ([datetime]::ParseExact("$Mois-01", 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)).AddMonths($Delta).ToString('yyyy-MM')
}
function Read-SmMoisJson([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    try { [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8) | ConvertFrom-Json } catch { $null }
}
function Write-SmMoisFiche([string]$Path, $Fiche, [string]$Data) {
    $sv = Join-Path $Data 'variabilite\sauvegardes-fiches'
    New-Item -ItemType Directory -Force -Path $sv | Out-Null
    Copy-Item -LiteralPath $Path -Destination (Join-Path $sv ('{0}-{1}.json' -f [IO.Path]::GetFileNameWithoutExtension($Path), (Get-Date).ToString('yyyyMMdd-HHmmss-fff'))) -Force
    [IO.File]::WriteAllText($Path, (ConvertTo-Json -InputObject $Fiche -Depth 10), (New-Object Text.UTF8Encoding($false)))
}
function Set-SmMoisPropriete($Objet, [string]$Nom, $Valeur) {
    if ($Objet.PSObject.Properties[$Nom]) { $Objet.$Nom = $Valeur } else { $Objet | Add-Member -NotePropertyName $Nom -NotePropertyValue $Valeur }
}

function Get-SmMoisPerimetre($Config) {
    $f = Join-Path ([string]$Config.dossierAnsamDiag) 'variabilite\perimetre.json'
    $p = Read-SmMoisJson $f
    if (-not $p) {
        $p = [ordered]@{ Clients = $SmMoisClientsDefaut; Note = 'Clients dont la variabilité est relevée chaque mois (décision du TAM du 16.11.2026). Liste modifiable ici.'; MisAJour = (Get-Date).ToString('yyyy-MM-dd') }
        New-Item -ItemType Directory -Force -Path (Split-Path $f) | Out-Null
        [IO.File]::WriteAllText($f, (ConvertTo-Json $p), (New-Object Text.UTF8Encoding($false)))
    }
    @($p.Clients | ForEach-Object { ([string]$_).ToUpper() } | Where-Object { $_ })
}

function Get-SmMoisMarques([string]$WorkDir, [string]$Mois) {
    $m = Read-SmMoisJson (Join-Path $WorkDir "mois\$Mois.json")
    if (-not $m) { $m = [pscustomobject]@{ mois = $Mois; kiosk = [pscustomobject]@{}; jira = [pscustomobject]@{}; tranches = [pscustomobject]@{} } }
    foreach ($k in 'kiosk', 'jira', 'tranches') { if (-not $m.PSObject.Properties[$k]) { Set-SmMoisPropriete $m $k ([pscustomobject]@{}) } }
    $m
}
function Save-SmMoisMarques([string]$WorkDir, $Marques) {
    Write-SmJson (Join-Path $WorkDir "mois\$($Marques.mois).json") $Marques
}

# Inventaires datés (AAAA-MM-JJ) d'un client : stockage AnsamDiag et ancien dossier de la fiche.
function Get-SmMoisInventaires([string]$Clients, [string]$Code, [string]$Type, [string]$Ancien) {
    $racines = @((Join-Path $Clients "$Code\Inventaires\$Type"))
    if ($Ancien) { $racines += $Ancien }
    $dates = foreach ($r in $racines) {
        if (-not $r -or -not (Test-Path -LiteralPath $r)) { continue }
        $d = @(Get-ChildItem -LiteralPath $r -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^\d{4}-\d{2}-\d{2}$' })
        if (-not $d.Count) {
            $b = @(Get-ChildItem -LiteralPath $r -Directory -ErrorAction SilentlyContinue | Where-Object { Get-ChildItem -LiteralPath $_.FullName -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^\d{4}-\d{2}-\d{2}$' } })
            if ($b.Count -eq 1) { $d = @(Get-ChildItem -LiteralPath $b[0].FullName -Directory | Where-Object { $_.Name -match '^\d{4}-\d{2}-\d{2}$' }) }
        }
        $d | ForEach-Object { $_.Name }
    }
    @($dates | Sort-Object -Unique -Descending)
}

function Get-SmMoisEtape([string]$WorkDir, [string]$Nom, [string]$Mois) {
    $s = Read-SmJson (Join-Path $WorkDir "travaux\$Nom.json")
    if ($s -and $Mois -and $s.mois -and $s.mois -ne $Mois) { return $null }     # étape d'un autre mois
    if ($s -and $s.etat -eq 'en_cours') {
        $since = ((Get-Date) - (ConvertTo-SmDateTime $s.le)).TotalSeconds
        if (($s.pid -and -not (Test-SmProcessus $s.pid $s.le)) -or (-not $s.pid -and $since -gt 30)) { $s.etat = 'echec'; $s.message = 'Étape interrompue avant la fin.' }
    }
    $s
}

# Semestre de plus-value dû : janvier → 2e semestre de l'an passé ; juillet → 1er semestre.
function Get-SmMoisSemestre([string]$Mois) {
    $an = [int]$Mois.Substring(0, 4); $m = [int]$Mois.Substring(5, 2)
    if ($m -ge 7) { "$an-H1" } else { "$($an - 1)-H2" }
}

function Get-SmMoisEtat($Config, [string]$WorkDir, [string]$Mois) {
    Assert-SmMois $Mois
    $data = [string]$Config.dossierAnsamDiag
    $clientsDir = [string]$Config.dossierClients
    $marques = Get-SmMoisMarques $WorkDir $Mois
    $precedent = Get-SmMoisDecale $Mois -1
    $clients = @()
    foreach ($code in (Get-SmMoisPerimetre $Config)) {
        $fiche = Read-SmMoisJson (Join-Path $clientsDir "$code.json")
        $contrat = Read-SmMoisJson (Join-Path $data "Facturation\contrats\$code.json")
        $c = [ordered]@{ code = $code; nom = $null; modele = $null; contrat = $null; base = $null; erreurs = @() }
        if (-not $fiche) { $c.erreurs += "Fiche client absente : $code.json"; $clients += $c; continue }
        $c.nom = if ($contrat -and $contrat.Nom) { [string]$contrat.Nom } else { [string]$fiche.Nom }
        $c.modele = [string]$fiche.Modele
        if ($contrat) { $c.contrat = [string]$contrat.Contrat.Modele; $c.base = $contrat.PerimetreContrat }
        $src = [ordered]@{ Users = $(if ($fiche.Sources.Users) { [string]$fiche.Sources.Users } else { 'M365' }); VM = [string]$fiche.Sources.VM; ServeurPhysique = [string]$fiche.Sources.ServeurPhysique }
        $c.sources = $src
        $c.manuel = [ordered]@{}; foreach ($k in $SmMoisManuels) { $c.manuel[$k] = $(if ($fiche.ValeursManuelles -and $null -ne $fiche.ValeursManuelles.$k) { [int]$fiche.ValeursManuelles.$k } else { 0 }) }
        $c.manuelConfirme = [string]$fiche.ValeursManuellesVerifiees
        $c.manuelOk = [bool]($c.manuelConfirme -and $c.manuelConfirme.Substring(0, [math]::Min(7, $c.manuelConfirme.Length)) -ge $Mois)
        $c.reglesSansSupport = (@(@($fiche.SansSupport) + @($fiche.SansSupportMotifs) | Where-Object { $_ -and ([string]$_).Trim() }).Count -gt 0)

        $besoinM365 = $src.Users -ne 'AD'
        $besoinAD = ($src.Users -ne 'M365') -or $src.VM -eq 'AD' -or $src.ServeurPhysique -eq 'AD'
        $c.inventaires = [ordered]@{}
        foreach ($t in 'M365', 'AD') {
            $ancien = if ($t -eq 'M365') { [string]$fiche.DossierM365 } else { [string]$fiche.DossierAD }
            $dates = Get-SmMoisInventaires $clientsDir $code $t $ancien
            $c.inventaires[$t] = [ordered]@{ requis = $(if ($t -eq 'M365') { $besoinM365 } else { $besoinAD }); duMois = @($dates | Where-Object { $_.StartsWith($Mois) } | Select-Object -First 1)[0]; dernier = @($dates)[0] }
        }

        $calc = Read-SmMoisJson (Join-Path $data "variabilite\_etat\$code\Mois-$Mois.json")
        $c.calcul = $null
        if ($calc) {
            $tr = @($calc.ATrancher | Where-Object { $_ } | ForEach-Object {
                $id = if ($_.Email) { [string]$_.Email } else { [string]$_.Compte }
                $dec = if ($marques.tranches.PSObject.Properties[$code] -and $marques.tranches.$code.PSObject.Properties[$id]) { $marques.tranches.$code.$id } else { $null }
                [ordered]@{ id = $id; compte = [string]$_.Compte; email = [string]$_.Email; licences = [string]$_.Licences; decision = $(if ($dec) { [string]$dec.decision } else { '' }) }
            })
            $c.calcul = [ordered]@{
                calcule = [string]$calc.Calcule
                version31 = [bool]$calc.PSObject.Properties['Inventaires']
                inventaires = $calc.Inventaires
                tableau = @($calc.Tableau | ForEach-Object { [ordered]@{ metrique = $_.Metrique; valeur = $_.MoisCourant; avant = $_.MoisAvant; ecart = $_.Ecart; source = $_.Source } })
                mouvements = @($calc.Mouvements | Where-Object { $_ } | ForEach-Object { [ordered]@{ mouvement = $_.Mouvement; compte = $_.Compte; licences = $_.Licences } })
                aTrancher = $tr
                calculStats = $calc.Calcul
            }
        }
        $reg = Read-SmMoisJson (Join-Path $data "variabilite\releves\$code.json")
        $c.releves = [ordered]@{ mois = $null; precedent = $null }
        foreach ($p in @(@('mois', $Mois), @('precedent', $precedent))) {
            if ($reg -and $reg.Releves -and $reg.Releves.PSObject.Properties[$p[1]]) {
                $t = [ordered]@{}; foreach ($k in $reg.Releves.($p[1]).PSObject.Properties) { $t[$k.Name] = [ordered]@{ valeur = $k.Value.Valeur; statut = $k.Value.Statut; date = $k.Value.DateReleve } }
                $c.releves[$p[0]] = $t
            }
        }
        $c.kiosk = if ($marques.kiosk.PSObject.Properties[$code]) { [string]$marques.kiosk.$code } else { '' }
        $c.jira = if ($marques.jira.PSObject.Properties[$code]) { [string]$marques.jira.$code } else { '' }
        $c.etat = Get-SmMoisEtatClient $c
        $clients += $c
    }

    # Facturation : elle porte sur le mois précédent.
    $fDir = Join-Path $data "Facturation\$precedent"
    $fac = [ordered]@{ mois = $precedent; etat = 'à faire' }
    if (Test-Path -LiteralPath (Join-Path $fDir 'cloture.json')) { $fac.etat = 'clos' }
    elseif ((Read-SmMoisJson (Join-Path $fDir 'journal-application.json')).ok) { $fac.etat = 'appliqué, à clore' }
    elseif (Test-Path -LiteralPath (Join-Path $fDir 'resultat.json')) { $fac.etat = 'calculé, décisions en cours' }
    elseif (Test-Path -LiteralPath $fDir) { $fac.etat = 'commencé' }

    $sem = Get-SmMoisSemestre $Mois
    $pv = Read-SmMoisJson (Join-Path $data "variabilite\plus-value\$sem.json")
    $etapes = [ordered]@{}
    foreach ($e in 'variabilite-collecte', 'variabilite-calcul', 'releves-reporter', 'plusvalue-controle', 'plusvalue-generer') { $etapes[$e] = Get-SmMoisEtape $WorkDir $e $Mois }
    [ordered]@{
        mois = $Mois; precedent = $precedent; facturation = $fac; clients = $clients; etapes = $etapes
        plusValue = [ordered]@{ semestre = $sem; du = ([int]$Mois.Substring(5, 2) -in 1, 7); resultat = $pv }
        rappels = @(Get-SmMoisRappels $Config $WorkDir (Get-Date))
    }
}

# Où en est un client pour le mois, dans l'ordre du cycle.
function Get-SmMoisEtatClient($c) {
    if ($c.erreurs.Count) { return 'fiche absente' }
    foreach ($t in 'M365', 'AD') { if ($c.inventaires[$t].requis -and -not $c.inventaires[$t].duMois) { return "inventaire $t à faire" } }
    if (-not $c.calcul) { return 'à calculer' }
    if (-not $c.calcul.version31) { return 'à recalculer' }
    foreach ($t in 'M365', 'AD') { if ($c.inventaires[$t].duMois -and $c.calcul.inventaires.$t -ne $c.inventaires[$t].duMois) { return 'à recalculer (inventaire plus récent)' } }
    if (@($c.calcul.aTrancher | Where-Object { -not $_.decision }).Count) { return 'comptes à trancher' }
    if (@($c.calcul.aTrancher | Where-Object { $_.decision -eq 'exclure' }).Count) { return 'à recalculer (exclusion)' }
    if (-not $c.manuelOk) { return 'valeurs manuelles à confirmer' }
    $r = $c.releves.mois
    if (-not $r -or @($r.Values | Where-Object { $_.statut -ne 'valide' }).Count) { return 'à valider' }
    if (-not $c.kiosk) { return 'Kiosk à saisir' }
    if (-not $c.jira) { return 'Jira à commenter' }
    'fait'
}

function Set-SmMoisManuel($Config, [string]$Code, $Valeurs) {
    $Code = $Code.ToUpper()
    $f = Join-Path ([string]$Config.dossierClients) "$Code.json"
    $fiche = Read-SmMoisJson $f
    if (-not $fiche) { throw "Fiche client absente : $Code" }
    if (-not $fiche.ValeursManuelles) { Set-SmMoisPropriete $fiche 'ValeursManuelles' ([pscustomobject]@{}) }
    $avant = @()
    foreach ($k in $SmMoisManuels) {
        if ($null -eq $Valeurs.$k -or "$($Valeurs.$k)" -eq '') { continue }
        $v = [int]$Valeurs.$k
        if ($v -lt 0 -or $v -gt 10000) { throw "Valeur invalide pour $k : $v" }
        $ancien = $fiche.ValeursManuelles.$k
        if ("$ancien" -ne "$v") { $avant += "$k $ancien -> $v" }
        Set-SmMoisPropriete $fiche.ValeursManuelles $k $v
    }
    Set-SmMoisPropriete $fiche 'ValeursManuellesVerifiees' (Get-Date).ToString('yyyy-MM-dd')
    Write-SmMoisFiche $f $fiche ([string]$Config.dossierAnsamDiag)
    if ($avant.Count) { "Valeurs manuelles de $Code confirmées, changées : $($avant -join ', '). Recalcule le mois." } else { "Valeurs manuelles de $Code confirmées sans changement." }
}

function Set-SmMoisMarque([string]$WorkDir, [string]$Mois, [string]$Code, [string]$Quoi, [bool]$Fait) {
    Assert-SmMois $Mois
    if ($Quoi -notin 'kiosk', 'jira') { throw "Marque inconnue : $Quoi" }
    $m = Get-SmMoisMarques $WorkDir $Mois
    Set-SmMoisPropriete $m.$Quoi $Code.ToUpper() $(if ($Fait) { (Get-Date).ToString('yyyy-MM-dd') } else { $null })
    Save-SmMoisMarques $WorkDir $m
}

# Compte nouveau du mois : « compter » (rien à changer, le compte est déjà compté) ou
# « exclure » (ajouté aux ExclusionsNominatives de la fiche ; il faut recalculer).
function Set-SmMoisTranche($Config, [string]$WorkDir, [string]$Mois, [string]$Code, [string]$Id, [string]$Compte, [string]$Decision, [string]$Note) {
    Assert-SmMois $Mois
    $Code = $Code.ToUpper()
    if ($Decision -notin 'compter', 'exclure', '') { throw "Décision inconnue : $Decision" }
    if (-not $Id) { throw 'Compte non précisé.' }
    $m = Get-SmMoisMarques $WorkDir $Mois
    if (-not $m.tranches.PSObject.Properties[$Code]) { Set-SmMoisPropriete $m.tranches $Code ([pscustomobject]@{}) }
    $msg = "Décision retirée pour $Compte."
    if ($Decision -eq 'exclure') {
        if (-not $Note) { throw 'Une exclusion demande une raison (elle reste dans la fiche).' }
        $f = Join-Path ([string]$Config.dossierClients) "$Code.json"
        $fiche = Read-SmMoisJson $f
        $liste = @($fiche.ExclusionsNominatives | Where-Object { $_ })
        if (-not ($liste | Where-Object { ([string]$_.Nom).ToLower() -eq $Id.ToLower() })) {
            $liste += [pscustomobject]@{ Nom = $Id; Categorie = 'Exclu'; Decision = "Exclu par le TAM le $((Get-Date).ToString('dd.MM.yyyy'))"; Note = $Note; MoisObserves = $Mois }
            Set-SmMoisPropriete $fiche 'ExclusionsNominatives' $liste
            Write-SmMoisFiche $f $fiche ([string]$Config.dossierAnsamDiag)
        }
        $msg = "$Compte exclu dans la fiche $Code : recalcule le mois."
    } elseif ($Decision -eq 'compter') { $msg = "$Compte compté." }
    Set-SmMoisPropriete $m.tranches.$Code $Id $(if ($Decision) { [pscustomobject]@{ decision = $Decision; compte = $Compte; note = $Note; le = (Get-Date).ToString('yyyy-MM-dd') } } else { $null })
    Save-SmMoisMarques $WorkDir $m
    $msg
}

# Comptes tranchés « compter » : ils ne bloquent plus la validation.
function Get-SmMoisTranches([string]$WorkDir, [string]$Mois, [string]$Code) {
    $m = Get-SmMoisMarques $WorkDir $Mois
    if (-not $m.tranches.PSObject.Properties[$Code]) { return @() }
    @($m.tranches.$Code.PSObject.Properties | Where-Object { $_.Value -and $_.Value.decision -eq 'compter' } | ForEach-Object { $_.Name })
}

# Rappels du cycle mensuel, pour la page et la synthèse du matin. Seulement ce qui reste à faire.
function Get-SmMoisRappels($Config, [string]$WorkDir, [datetime]$Jour) {
    $data = [string]$Config.dossierAnsamDiag
    $mois = $Jour.ToString('yyyy-MM')
    $prec = $Jour.AddMonths(-1).ToString('yyyy-MM')
    $finMois = (New-Object DateTime($Jour.Year, $Jour.Month, 1)).AddMonths(1).AddDays(-1)
    $r = @()
    $clientsDir = [string]$Config.dossierClients
    $perimetre = @(Get-SmMoisPerimetre $Config)
    if (($finMois - $Jour.Date).TotalDays -le 6) {
        $manque = @($perimetre | Where-Object {
            $f = Read-SmMoisJson (Join-Path $clientsDir "$_.json")
            $f -and ([string]$f.Sources.Users -ne 'AD') -and -not (@(Get-SmMoisInventaires $clientsDir $_ 'M365' ([string]$f.DossierM365)) | Where-Object { $_.StartsWith($mois) })
        })
        if ($manque.Count) { $r += "Variabilité de $mois : inventaires M365 à collecter avant la fin du mois ($($manque -join ', ')), onglet Mois de Panoramix." }
    }
    if ($Jour.Day -le 10) {
        $fDir = Join-Path $data "Facturation\$prec"
        if (-not (Test-Path -LiteralPath (Join-Path $fDir 'cloture.json'))) { $r += "Facturation de $prec : pas encore close (onglet Facturation)." }
        $marques = Get-SmMoisMarques $WorkDir $prec
        $ouverts = @($perimetre | Where-Object {
            $reg = Read-SmMoisJson (Join-Path $data "variabilite\releves\$_.json")
            $rm = if ($reg -and $reg.Releves -and $reg.Releves.PSObject.Properties[$prec]) { $reg.Releves.$prec } else { $null }
            -not $rm -or @($rm.PSObject.Properties | Where-Object { $_.Value.Statut -ne 'valide' }).Count -or -not $marques.kiosk.PSObject.Properties[$_] -or -not $marques.jira.PSObject.Properties[$_]
        })
        if ($ouverts.Count) { $r += "Variabilité de $prec : à valider, saisir dans Kiosk ou commenter dans Jira pour $($ouverts -join ', ') (onglet Mois)." }
    }
    if ($Jour.Month -in 1, 7 -and $Jour.Day -le 20) {
        $sem = Get-SmMoisSemestre $mois
        $pv = Read-SmMoisJson (Join-Path $data "variabilite\plus-value\$sem.json")
        if (-not $pv -or $pv.etape -ne 'generer') { $r += "Plus-value $sem : décomptes à générer au format de Mario (onglet Mois)." }
    }
    $r
}
