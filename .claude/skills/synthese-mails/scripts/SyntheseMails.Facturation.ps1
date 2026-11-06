<#
.SYNOPSIS
    Facturation mensuelle dans la page de suivi : état, décisions, contrôle des
    journaux du fichier commun. À charger par dot-sourcing depuis
    Configurer-SyntheseMails.ps1.

.DESCRIPTION
    Le moteur et les règles restent dans <dossierAnsamDiag>\Facturation. Ici :
    - decisions.json n'est écrit que sur un clic dans la page, avec une sauvegarde
      avant chaque écriture ; une décision d'un mois passé est conservée ;
    - le fichier commun n'est jamais touché : le TAM lance le script dans Excel et
      colle son journal, contrôlé ligne à ligne contre les données du mois.
#>

function Get-SmFactureRacine($Config) { Join-Path ([string]$Config.dossierAnsamDiag) 'Facturation' }
function Assert-SmFactureMois([string]$Mois) { if ($Mois -notmatch '^\d{4}-\d{2}$') { throw "Mois invalide : $Mois" } }

# Fichier JSON tel quel, pour le renvoyer à la page sans le relire ni le réécrire.
function Get-SmJsonBrut([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return 'null' }
    ([IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8)).TrimStart([char]0xFEFF)
}

function Get-SmFactureDecisions($Racine) {
    [IO.File]::ReadAllText((Join-Path $Racine 'decisions.json'), [Text.Encoding]::UTF8) | ConvertFrom-Json
}

# La décision qui vaut pour ce mois, ou $null.
function Find-SmFactureDecision($Decisions, [string]$Cle, [string]$Mois) {
    $property = $Decisions.Tickets.PSObject.Properties[$Cle]
    if (-not $property) { return $null }
    foreach ($d in @($property.Value) + @($property.Value.Anciennes | Where-Object { $_ })) {
        if ($d.Mois -eq $Mois -or @($d.Reconduite) -contains $Mois) { return $d }
    }
    $null
}

function Get-SmFactureEtat($Config, [string]$WorkDir, [string]$Mois) {
    Assert-SmFactureMois $Mois
    $racine = Get-SmFactureRacine $Config
    $dir = Join-Path $racine $Mois
    $downloads = Join-Path $env:USERPROFILE 'Downloads'
    $candidats = @(Get-ChildItem -LiteralPath $downloads -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '(?i)^Jira.*\.(xlsx|csv)$' } |
        Sort-Object LastWriteTime -Descending | Select-Object -First 5 | ForEach-Object { [ordered]@{ nom = $_.Name; le = $_.LastWriteTime.ToString('o'); taille = $_.Length } })
    $sources = @(Get-ChildItem -LiteralPath $dir -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '(?i)^Jira.*\.(xlsx|csv)$' } | Sort-Object LastWriteTime -Descending | ForEach-Object { $_.Name })
    $disponibles = @(Get-ChildItem -LiteralPath $racine -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^\d{4}-\d{2}$' } | Sort-Object Name -Descending | ForEach-Object { $_.Name })
    $etapes = [ordered]@{}
    foreach ($e in 'moteur', 'controles', 'analyse') {
        $state = Read-SmJson (Join-Path $WorkDir "travaux\facturation-$e.json")
        if ($state -and $state.mois -ne $Mois) { $state = $null }
        if ($state -and $state.etat -eq 'en_cours') {
            $since = ((Get-Date) - (ConvertTo-SmDateTime $state.le)).TotalSeconds
            if (($state.pid -and -not (Get-Process -Id $state.pid -ErrorAction SilentlyContinue)) -or (-not $state.pid -and $since -gt 30)) { $state.etat = 'echec'; $state.message = 'Étape interrompue avant la fin.' }
        }
        $etapes[$e] = $state
    }
    # Décisions déjà prises pour ce mois, sur les tickets du mois.
    $result = Get-SmJsonBrut (Join-Path $dir 'resultat.json')
    $valides = [ordered]@{}
    if ($result -ne 'null') {
        $decisions = Get-SmFactureDecisions $racine
        foreach ($key in @(($result | ConvertFrom-Json).lignes | ForEach-Object { [string]$_.Ticket } | Sort-Object -Unique)) {
            $d = Find-SmFactureDecision $decisions $key $Mois
            if ($d) { $valides[$key] = [ordered]@{ statut = $d.Statut; motif = $d.Motif; source = $d.Source; valide = [bool]$d.Valide; reconduite = (@($d.Reconduite) -contains $Mois) } }
        }
    }
    $ts = Join-Path $dir "colorer-fichier-commun-$Mois.ts"
    $journaux = [ordered]@{}
    foreach ($m in 'simulation', 'application') { $journaux[$m] = Read-SmJson (Join-Path $dir "journal-$m.json") }
    $script = $(if (Test-Path -LiteralPath $ts) { (Get-Item -LiteralPath $ts).LastWriteTime.ToString('o') } else { $null })
    $decisionsLe = (Get-Item -LiteralPath (Join-Path $racine 'decisions.json')).LastWriteTime.ToString('o')
    $meta = [ordered]@{ mois = $Mois; disponibles = $disponibles; candidats = $candidats; sources = $sources; etapes = $etapes; decisions = $valides
                        journaux = $journaux; scriptLe = $script; decisionsLe = $decisionsLe; clos = (Test-Path -LiteralPath (Join-Path $dir 'cloture.json')) } | ConvertTo-Json -Depth 8 -Compress
    '{"meta":' + $meta + ',"resultat":' + $result + ',"controles":' + (Get-SmJsonBrut (Join-Path $dir 'controles.json')) +
        ',"propositions":' + (Get-SmJsonBrut (Join-Path $dir 'propositions.json')) + '}'
}

# Copie le fichier téléchargé dans le dossier du mois. Le nom vient de la liste des Téléchargements.
function Copy-SmFactureSource($Config, [string]$Mois, [string]$Nom) {
    Assert-SmFactureMois $Mois
    if ($Nom -notmatch '(?i)^Jira[^\\/]*\.(xlsx|csv)$') { throw "Fichier refusé : $Nom" }
    $from = Join-Path (Join-Path $env:USERPROFILE 'Downloads') $Nom
    if (-not (Test-Path -LiteralPath $from)) { throw "Fichier introuvable : $Nom" }
    $dir = Join-Path (Get-SmFactureRacine $Config) $Mois
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    $to = Join-Path $dir ("Jira-source-{0:yyyyMMdd-HHmm}{1}" -f (Get-Date), [IO.Path]::GetExtension($Nom))
    # Excel peut garder le fichier ouvert : lecture partagée.
    $in = [IO.File]::Open($from, 'Open', 'Read', 'ReadWrite')
    try { $out = [IO.File]::Create($to); try { $in.CopyTo($out) } finally { $out.Close() } } finally { $in.Close() }
    Split-Path $to -Leaf
}

function Save-SmFactureDecisions($Decisions, [string]$Racine, [string]$Mois) {
    $path = Join-Path $Racine 'decisions.json'
    $backups = Join-Path $Racine "$Mois\sauvegardes"
    New-Item -ItemType Directory -Force -Path $backups | Out-Null
    Copy-Item -LiteralPath $path -Destination (Join-Path $backups ("decisions-{0:yyyyMMdd-HHmmssfff}.json" -f (Get-Date)))
    $temp = "$path.tmp"
    [IO.File]::WriteAllText($temp, (ConvertTo-Json -InputObject $Decisions -Depth 12), (New-Object Text.UTF8Encoding($false)))
    [void](ConvertFrom-Json ([IO.File]::ReadAllText($temp, [Text.Encoding]::UTF8)))
    Move-Item -LiteralPath $temp -Destination $path -Force
}

# Une décision de ticket, sur mon clic. -Posted : mois, cle, client, reconduire, ou statut, motif, source,
# confiance, deplacements, motifDeplacements, offre, heuresOffre, lignes.
function Set-SmFactureDecision($Config, $Posted) {
    $mois = [string]$Posted.mois
    Assert-SmFactureMois $mois
    $key = [string]$Posted.cle
    if ($key -notmatch '^[A-Z][A-Z0-9]+-\d+$') { throw "Ticket invalide : $key" }
    $racine = Get-SmFactureRacine $Config
    if (Test-Path -LiteralPath (Join-Path $racine "$mois\cloture.json")) { throw "Le mois $mois est clos." }
    $decisions = Get-SmFactureDecisions $racine
    $existing = $(if ($decisions.Tickets.PSObject.Properties[$key]) { $decisions.Tickets.PSObject.Properties[$key].Value })
    $today = (Get-Date).ToString('dd.MM.yyyy')

    if ($Posted.reconduire) {
        if (-not $existing) { throw "Aucune décision à reconduire pour $key." }
        $existing | Add-Member -NotePropertyName Reconduite -NotePropertyValue @(@($existing.Reconduite | Where-Object { $_ }) + $mois | Sort-Object -Unique) -Force
    } else {
        $statut = [string]$Posted.statut
        if (@('N', 'I', 'Q', 'QX', 'F', 'P') -notcontains $statut) { throw "Statut invalide : $statut" }
        $motif = ([string]$Posted.motif).Trim()
        if (-not $motif) { throw 'Motif obligatoire.' }
        if ($statut -eq 'P' -and -not $Posted.offre) { throw 'Offre obligatoire pour un projet.' }
        $source = ([string]$Posted.source).Trim()
        $new = [ordered]@{ Mois = $mois; Client = [string]$Posted.client; Statut = $statut; Confiance = $(if ($Posted.confiance) { [string]$Posted.confiance } else { 'forte' }); Valide = $true
                           Motif = $motif; Source = "Décision du TAM du $today" + $(if ($source) { " ; $source" } else { '' }) }
        if ($Posted.deplacements) { $new.Deplacements = [string]$Posted.deplacements; $new.MotifDeplacements = [string]$Posted.motifDeplacements }
        if ($statut -eq 'P') { $new.Offre = [string]$Posted.offre; $new.HeuresOffre = $Posted.heuresOffre }
        if ($existing -and $existing.Mois -and $existing.Mois -ne $mois) {
            # Une décision d'un mois passé reste : recalculer ce mois doit redonner le même résultat.
            $old = [ordered]@{}
            foreach ($p in $existing.PSObject.Properties) { if ($p.Name -ne 'Anciennes') { $old[$p.Name] = $p.Value } }
            $new.Anciennes = @([pscustomobject]$old) + @($existing.Anciennes | Where-Object { $_ })
        } elseif ($existing -and $existing.Anciennes) {
            $new.Anciennes = @($existing.Anciennes)
        }
        $decisions.Tickets | Add-Member -NotePropertyName $key -NotePropertyValue ([pscustomobject]$new) -Force
        foreach ($l in @($Posted.lignes | Where-Object { $_ })) {
            $entry = [pscustomobject][ordered]@{ Date = [string]$l.date; Minutes = [int]$l.minutes; Statut = [string]$l.statut; Motif = [string]$l.motif; Confiance = 'forte'; Valide = $true; Mois = $mois }
            if (-not $decisions.Lignes.PSObject.Properties[$key]) { $decisions.Lignes | Add-Member -NotePropertyName $key -NotePropertyValue @() }
            $others = @($decisions.Lignes.$key | Where-Object { $_ -and -not ($_.Date -eq $entry.Date -and [int]$_.Minutes -eq $entry.Minutes) })
            $decisions.Lignes.$key = @($others + $entry)
        }
    }
    Save-SmFactureDecisions $decisions $racine $mois
}

function Get-SmFactureScript($Config, [string]$Mois, [bool]$Appliquer) {
    Assert-SmFactureMois $Mois
    $path = Join-Path (Get-SmFactureRacine $Config) "$Mois\colorer-fichier-commun-$Mois.ts"
    if (-not (Test-Path -LiteralPath $path)) { throw 'Script du mois introuvable : lance le calcul.' }
    $text = ([IO.File]::ReadAllText($path, [Text.Encoding]::UTF8)).TrimStart([char]0xFEFF)
    if ($Appliquer) {
        if ($text -notmatch 'const APPLIQUER = false;') { throw 'Ligne APPLIQUER introuvable dans le script.' }
        $text = $text -replace 'const APPLIQUER = false;', 'const APPLIQUER = true;'
    }
    $text
}

# Contrôle d'un journal collé depuis Excel, ligne à ligne, contre les données du script du mois.
function Test-SmFactureJournal($Config, [string]$Mois, [string]$Mode, [string]$Texte) {
    Assert-SmFactureMois $Mois
    if (@('simulation', 'application') -notcontains $Mode) { throw "Mode inconnu : $Mode" }
    $racine = Get-SmFactureRacine $Config
    $dir = Join-Path $racine $Mois
    $tsText = Get-SmFactureScript $Config $Mois $false
    $m = [regex]::Match($tsText, '(?m)^const DONNEES: Donnees = (.*);\s*$')
    if (-not $m.Success) { throw 'Données du script illisibles.' }
    $data = $m.Groups[1].Value | ConvertFrom-Json
    $attendues = @($data.lignes).Count
    $points = New-Object System.Collections.ArrayList
    function Add-Point([bool]$Ok, [string]$Texte) { [void]$points.Add([ordered]@{ ok = $Ok; texte = $Texte }) }

    $titre = [regex]::Match($Texte, '=== Coloration (\S+) - (APPLICATION|SIMULATION)')
    Add-Point ($titre.Success -and $titre.Groups[1].Value -eq $Mois) "Journal du mois $Mois"
    $bonMode = $(if ($Mode -eq 'simulation') { 'SIMULATION' } else { 'APPLICATION' })
    Add-Point ($titre.Groups[2].Value -eq $bonMode) "Mode $bonMode"
    Add-Point ($Texte -notmatch '(?m)^ARRET') 'Aucun arrêt du script'
    $appro = [regex]::Match($Texte, 'Rapprochees : (\d+) ; deja conformes : (\d+) ; couleur a changer : (\d+) ; debordement a retirer : (\d+) ; anomalies : (\d+)')
    Add-Point ($appro.Success -and [int]$appro.Groups[1].Value -eq $attendues) "$($appro.Groups[1].Value) lignes rapprochées sur $attendues"
    Add-Point ($appro.Success -and [int]$appro.Groups[5].Value -eq 0 -and $Texte -notmatch '(?m)^ANOMALIE') 'Aucune anomalie'
    $vert = [regex]::Match($Texte, 'Vert utilise : (#[0-9A-Fa-f]{6})')
    Add-Point ($vert.Groups[1].Value.ToUpperInvariant() -eq '#92D050') "Vert $($vert.Groups[1].Value)"
    $rest = [regex]::Match($Texte, 'absentes des donnees \(non touchees\) : (\d+)')
    Add-Point ($rest.Success -and [int]$rest.Groups[1].Value -eq 0) "Lignes de tes clients non traitées : $($rest.Groups[1].Value)"

    $index = @{}
    foreach ($l in @($data.lignes)) { $index["$($l.k)|$($l.d)|$($l.a)|$($l.m)"] = $l }
    $changes = New-Object System.Collections.ArrayList
    $faux = 0
    foreach ($c in [regex]::Matches($Texte, '(?m)^ligne (\d+) \| (\S+) \| (\S+) \| (.+?) \| (\d+) min \| (.+?) \| (.*)$')) {
        $g = $c.Groups
        $line = $index["$($g[2].Value)|$($g[3].Value)|$($g[4].Value)|$($g[5].Value)"]
        $quoi = $g[6].Value
        $target = [regex]::Match($quoi, '->\s*(\S+)')
        $conforme = [bool]$line
        if ($line -and $target.Success) {
            $versVert = $target.Groups[1].Value -match '(?i)vert|92D050'
            $conforme = ($versVert -eq [bool]$line.vert)
        }
        if (-not $conforme) { $faux++ }
        [void]$changes.Add([ordered]@{ ligne = [int]$g[1].Value; ticket = $g[2].Value; date = $g[3].Value; auteur = $g[4].Value; minutes = [int]$g[5].Value; quoi = $quoi; decision = $g[7].Value.Trim(); conforme = $conforme })
    }
    if ($appro.Success) {
        $annonces = [int]$appro.Groups[3].Value + [int]$appro.Groups[4].Value
        Add-Point ($changes.Count -ge [int]$appro.Groups[3].Value -and $changes.Count -le $annonces) "$($changes.Count) changements lus (couleur : $($appro.Groups[3].Value), débordement : $($appro.Groups[4].Value))"
    }
    Add-Point ($faux -eq 0) "Chaque changement correspond à une décision du mois ($faux écart)"
    if ($Mode -eq 'application') { Add-Point ($Texte -match '(?m)^Termine : ') 'Script terminé' }

    $ok = @($points | Where-Object { -not $_.ok }).Count -eq 0
    $verdict = [ordered]@{ mode = $Mode; ok = $ok; le = (Get-Date).ToString('o'); points = @($points); changements = @($changes); journal = $Texte }
    Write-SmJson (Join-Path $dir "journal-$Mode.json") $verdict
    if ($Mode -eq 'application' -and $ok) {
        $md = "# Coloration du fichier commun - $Mois`r`n`r`nContrôlée par Panoramix le $((Get-Date).ToString('dd.MM.yyyy HH:mm')) : $($changes.Count) changements conformes.`r`n`r`n## Journal du script`r`n`r`n``````text`r`n$Texte`r`n```````r`n"
        [IO.File]::WriteAllText((Join-Path $dir 'journal-application-fichier-commun.md'), $md, (New-Object Text.UTF8Encoding($true)))
    }
    $verdict
}

# Clôture : le fichier commun est à jour, les décisions du mois sont figées.
function Close-SmFacture($Config, [string]$Mois) {
    Assert-SmFactureMois $Mois
    $racine = Get-SmFactureRacine $Config
    $dir = Join-Path $racine $Mois
    $journal = Read-SmJson (Join-Path $dir 'journal-application.json')
    if (-not $journal -or -not $journal.ok) { throw 'Le journal d''application contrôlé est requis avant la clôture.' }
    $decisions = Get-SmFactureDecisions $racine
    $n = 0
    foreach ($p in $decisions.Tickets.PSObject.Properties) {
        if ($p.Value.Mois -eq $Mois -and -not $p.Value.Valide) { $p.Value.Valide = $true; $n++ }
    }
    if ($n) { Save-SmFactureDecisions $decisions $racine $Mois }
    Write-SmJson (Join-Path $dir 'cloture.json') ([ordered]@{ le = (Get-Date).ToString('o'); decisionsValidees = $n })
}
