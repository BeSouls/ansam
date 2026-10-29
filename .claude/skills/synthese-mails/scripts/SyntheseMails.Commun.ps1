<#
.SYNOPSIS
    Fonctions communes de la synthèse mails : configuration, état, calendrier
    (jours d'envoi, absences) et mise en forme HTML. À charger par
    dot-sourcing : . "$PSScriptRoot\SyntheseMails.Commun.ps1"
#>

$SmWorkDir = Join-Path $env:USERPROFILE 'SyntheseMails'
$SmInvariant = [Globalization.CultureInfo]::InvariantCulture
$SmJours = @{ 1 = 'lundi'; 2 = 'mardi'; 3 = 'mercredi'; 4 = 'jeudi'; 5 = 'vendredi'; 6 = 'samedi'; 7 = 'dimanche' }

function Get-SmDefaultConfig {
    [ordered]@{
        actif               = $true
        heure               = '07:10'
        joursEnvoi          = @(1, 2, 3, 4, 5)   # 1 = lundi … 7 = dimanche
        fenetreJours        = 14
        envoiMail           = $true
        dossierAnsamDiag    = Join-Path $env:USERPROFILE 'OneDrive - ANSAM group\Documents\AnsamDiag-Donnees'
        dossierPortefeuille = 'Boîte de réception\MES CLIENTS'
        absences            = @()                # [{ debut: 'AAAA-MM-JJ'; fin: 'AAAA-MM-JJ'; libelle: '…' }]
        portPage            = 8765               # page locale : http://127.0.0.1:<port>/
    }
}

# Jeton de la page locale, créé une fois et gardé dans <WorkDir>\page.token :
# l'adresse reste la même d'un démarrage à l'autre (favori, lien du mail).
function Get-SmPageToken([string]$WorkDir = $SmWorkDir) {
    $path = Join-Path $WorkDir 'page.token'
    if (Test-Path $path) {
        $token = ([IO.File]::ReadAllText($path)).Trim()
        if ($token -match '^[0-9a-f]{32}$') { return $token }
    }
    New-Item -ItemType Directory -Force -Path $WorkDir | Out-Null
    $token = [guid]::NewGuid().ToString('N')
    [IO.File]::WriteAllText($path, $token)
    $token
}

function Get-SmPageUrl([string]$WorkDir = $SmWorkDir) {
    "http://127.0.0.1:$((Get-SmConfig $WorkDir).portPage)/?t=$(Get-SmPageToken $WorkDir)"
}

function Read-SmJson([string]$Path) {
    if (-not (Test-Path $Path)) { return $null }
    [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8) | ConvertFrom-Json
}

function Write-SmJson([string]$Path, $Value) {
    New-Item -ItemType Directory -Force -Path (Split-Path $Path) | Out-Null
    [IO.File]::WriteAllText($Path, ($Value | ConvertTo-Json -Depth 6), (New-Object Text.UTF8Encoding($false)))
}

# Configuration du fichier, complétée par les valeurs par défaut.
function Get-SmConfig([string]$WorkDir = $SmWorkDir) {
    $config = Get-SmDefaultConfig
    $saved = Read-SmJson (Join-Path $WorkDir 'config.json')
    if ($saved) {
        foreach ($property in $saved.PSObject.Properties) {
            if ($config.Contains($property.Name)) { $config[$property.Name] = $property.Value }
        }
    }
    $config.joursEnvoi = @($config.joursEnvoi | ForEach-Object { [int]$_ })
    $config.absences = @($config.absences | Where-Object { $_ } | ForEach-Object {
        [ordered]@{ debut = [string]$_.debut; fin = [string]$_.fin; libelle = [string]$_.libelle }
    })
    $config
}

# Contrôle une configuration ; renvoie la liste des erreurs (vide si tout va bien).
function Test-SmConfig($Config) {
    $errors = @()
    if ($Config.heure -notmatch '^([01]\d|2[0-3]):[0-5]\d$') { $errors += "Heure invalide : « $($Config.heure) » (format HH:mm)." }
    $days = @($Config.joursEnvoi)
    if (-not $days.Count) { $errors += 'Choisir au moins un jour d''envoi.' }
    foreach ($day in $days) { if ([int]$day -lt 1 -or [int]$day -gt 7) { $errors += "Jour d'envoi invalide : $day." } }
    if ([int]$Config.fenetreJours -lt 1 -or [int]$Config.fenetreJours -gt 60) { $errors += 'La fenêtre doit être comprise entre 1 et 60 jours.' }
    foreach ($absence in @($Config.absences)) {
        $start = ConvertTo-SmDate $absence.debut
        $end = ConvertTo-SmDate $absence.fin
        if (-not $start -or -not $end) { $errors += "Absence « $($absence.libelle) » : dates invalides (AAAA-MM-JJ)." }
        elseif ($end -lt $start) { $errors += "Absence « $($absence.libelle) » : la fin précède le début." }
    }
    , $errors
}

function Save-SmConfig($Config, [string]$WorkDir = $SmWorkDir) {
    Write-SmJson (Join-Path $WorkDir 'config.json') $Config
}

function Get-SmEtat([string]$WorkDir = $SmWorkDir) {
    $etat = Read-SmJson (Join-Path $WorkDir 'etat.json')
    if (-not $etat) { $etat = [pscustomobject]@{ derniereSynthese = $null; dernierMode = $null; dernierFichier = $null } }
    $etat
}

function Save-SmEtat($Etat, [string]$WorkDir = $SmWorkDir) {
    Write-SmJson (Join-Path $WorkDir 'etat.json') $Etat
}

function ConvertTo-SmDate([string]$Text) {
    $date = [datetime]::MinValue
    if ([datetime]::TryParseExact($Text, 'yyyy-MM-dd', $SmInvariant, 'None', [ref]$date)) { return $date }
    $null
}

function Get-SmJourSemaine([datetime]$Date) {
    $day = [int]$Date.DayOfWeek
    if ($day -eq 0) { 7 } else { $day }
}

function Get-SmAbsence($Config, [datetime]$Date) {
    foreach ($absence in @($Config.absences)) {
        $start = ConvertTo-SmDate $absence.debut
        $end = ConvertTo-SmDate $absence.fin
        if ($start -and $end -and $Date.Date -ge $start -and $Date.Date -le $end) { return $absence }
    }
    $null
}

# Décide si une synthèse doit partir à $Moment, et sur quelle période.
function Get-SmPlan($Config, $Etat, [datetime]$Moment = (Get-Date)) {
    $plan = [ordered]@{ lancer = $false; raison = $null; depuis = $null; jours = [int]$Config.fenetreJours; absences = @() }

    if (-not $Config.actif) { $plan.raison = 'Synthèse désactivée dans la configuration.'; return $plan }
    $absence = Get-SmAbsence $Config $Moment
    if ($absence) { $plan.raison = "Absence en cours : $($absence.libelle) ($($absence.debut) → $($absence.fin))."; return $plan }
    if (@($Config.joursEnvoi) -notcontains (Get-SmJourSemaine $Moment)) {
        $plan.raison = "Pas d'envoi le $($SmJours[(Get-SmJourSemaine $Moment)])."
        return $plan
    }

    $last = $null
    if ($Etat.derniereSynthese) { $last = [datetime]::Parse([string]$Etat.derniereSynthese, $SmInvariant) }
    if ($last -and $last.Date -eq $Moment.Date) { $plan.raison = "Synthèse déjà produite aujourd'hui à $($last.ToString('HH:mm'))."; return $plan }
    # La tâche se déclenche aussi à l'ouverture de session et au déverrouillage : pas avant l'heure prévue.
    $parts = $Config.heure -split ':'
    if ($Moment -lt $Moment.Date.AddHours([int]$parts[0]).AddMinutes([int]$parts[1])) {
        $plan.raison = "Avant l'heure prévue ($($Config.heure))."
        return $plan
    }

    if (-not $last) {
        # Première synthèse : depuis le jour d'envoi précédent, 7h00.
        $last = $Moment.Date.AddDays(-1)
        while (@($Config.joursEnvoi) -notcontains (Get-SmJourSemaine $last) -and ($Moment.Date - $last).Days -lt 7) { $last = $last.AddDays(-1) }
        $last = $last.AddHours(7)
    }
    $plan.depuis = $last
    $plan.absences = @(@($Config.absences) | Where-Object {
        $end = ConvertTo-SmDate $_.fin
        $end -and $end -ge $last.Date -and $end -lt $Moment.Date
    })
    # La fenêtre couvre au moins la période écoulée depuis la dernière synthèse, plus une semaine.
    $gap = [int][math]::Ceiling(($Moment - $last).TotalDays)
    $plan.jours = [math]::Max([int]$Config.fenetreJours, $gap + 7)
    $plan.lancer = $true
    $plan
}

# Prochaine synthèse prévue (pour la page de configuration).
function Get-SmProchaine($Config, $Etat, [datetime]$Moment = (Get-Date)) {
    $parts = $Config.heure -split ':'
    for ($i = 0; $i -le 90; $i++) {
        $candidate = $Moment.Date.AddDays($i).AddHours([int]$parts[0]).AddMinutes([int]$parts[1])
        if ($candidate -lt $Moment) { continue }
        $plan = Get-SmPlan $Config $Etat $candidate
        if ($plan.lancer) { return [ordered]@{ moment = $candidate; plan = $plan } }
        if (-not $Config.actif) { return $null }
    }
    $null
}

# Exécutable de Claude Code : « claude » du PATH, sinon l'installation native
# (%USERPROFILE%\.local\bin), que l'installeur n'ajoute pas toujours au PATH.
function Resolve-SmClaude([string]$ClaudeExe) {
    if ($ClaudeExe) { return $ClaudeExe }
    $command = Get-Command claude -ErrorAction SilentlyContinue
    $nativeExe = Join-Path $env:USERPROFILE '.local\bin\claude.exe'
    if ($command) { return $command.Source }
    if (Test-Path $nativeExe) { return $nativeExe }
    throw "Claude Code introuvable : installe-le (irm https://claude.ai/install.ps1 | iex) ou précise -ClaudeExe."
}

# --- Registre de suivi des points ------------------------------------------
# suivi.json garde chaque point d'une synthèse à l'autre. Statuts :
#   a_faire        une action m'attend
#   attente_tiers  mon action est faite, j'attends quelqu'un
#   traite         réglé
#   ignore         sans objet pour moi (posé par moi seulement)
# Un statut posé par moi (majPar = 'moi') n'est changé par la synthèse que si
# un message plus récent que mon marquage arrive sur le point.

$SmStatuts = @('a_faire', 'attente_tiers', 'traite', 'ignore')

function Get-SmSuivi([string]$WorkDir = $SmWorkDir) {
    $saved = Read-SmJson (Join-Path $WorkDir 'suivi.json')
    $points = New-Object System.Collections.ArrayList   # pas de List[object] : @() échoue dessus sous PS 5.1
    if ($saved) {
        foreach ($point in @($saved.points)) {
            if (-not $point) { continue }
            $copy = [ordered]@{}
            foreach ($property in $point.PSObject.Properties) { $copy[$property.Name] = $property.Value }
            $copy.historique = @($point.historique | Where-Object { $_ })
            [void]$points.Add($copy)
        }
    }
    [ordered]@{ majLe = $(if ($saved) { $saved.majLe }); points = $points }
}

function Save-SmSuivi($Suivi, [string]$WorkDir = $SmWorkDir) {
    $Suivi.majLe = (Get-Date).ToString('o')
    Write-SmJson (Join-Path $WorkDir 'suivi.json') ([ordered]@{ majLe = $Suivi.majLe; points = @($Suivi.points) })
}

function ConvertTo-SmDateTime($Value) {
    if (-not $Value) { return $null }
    $date = [datetime]::MinValue
    if ([datetime]::TryParse([string]$Value, $SmInvariant, [Globalization.DateTimeStyles]::None, [ref]$date)) { return $date }
    $null
}

function Add-SmHistorique($Point, [string]$Par, [string]$Statut, [string]$Texte) {
    $Point.historique = @(@($Point.historique) + [ordered]@{ le = (Get-Date).ToString('o'); par = $Par; statut = $Statut; texte = $Texte } |
        Select-Object -Last 20)
}

# Intègre les points proposés par la synthèse. Renvoie les points dont le
# statut a changé, pour le compte rendu.
function Merge-SmSuivi($Suivi, $Updates) {
    $changes = @()
    $today = (Get-Date).ToString('yyyyMMdd')
    $counter = @($Suivi.points | Where-Object { $_.id -like "P-$today-*" }).Count
    foreach ($update in @($Updates)) {
        if (-not $update -or -not $update.titre) { continue }
        $statut = [string]$update.statut
        if ($SmStatuts -notcontains $statut -or $statut -eq 'ignore') { $statut = 'a_faire' }

        $point = $null
        if ($update.id) { $point = $Suivi.points | Where-Object { $_.id -eq $update.id } | Select-Object -First 1 }
        if (-not $point -and $update.ticket) {
            $point = $Suivi.points | Where-Object { $_.ticket -eq $update.ticket -and $_.statut -ne 'traite' -and $_.statut -ne 'ignore' } | Select-Object -First 1
        }
        if (-not $point) {
            $id = [string]$update.id
            if (-not $id -or ($Suivi.points | Where-Object { $_.id -eq $id })) {
                if ($update.ticket) { $id = [string]$update.ticket } else { $counter++; $id = "P-$today-$counter" }
                if ($Suivi.points | Where-Object { $_.id -eq $id }) { $counter++; $id = "P-$today-$counter" }
            }
            $point = [ordered]@{ id = $id; creeLe = (Get-Date).ToString('o'); statut = $null; majPar = 'synthese'; majLe = $null; noteMoi = ''; historique = @() }
            [void]$Suivi.points.Add($point)
        }

        # Champs descriptifs : toujours mis à jour.
        foreach ($field in 'titre', 'client', 'ticket', 'priorite', 'action', 'echeance', 'source', 'dernierMessage') {
            if ($update.PSObject.Properties.Name -contains $field) { $point[$field] = $update.$field }
        }

        if ($point.statut -eq $statut) { continue }
        if ($point.majPar -eq 'moi' -and $point.statut) {
            # Mon marquage tient, sauf message plus récent que lui.
            $marked = ConvertTo-SmDateTime $point.majLe
            $lastMessage = ConvertTo-SmDateTime $update.dernierMessage
            if (-not ($marked -and $lastMessage -and $lastMessage -gt $marked)) { continue }
        }
        $previous = $point.statut
        $point.statut = $statut
        $point.majPar = 'synthese'
        $point.majLe = (Get-Date).ToString('o')
        Add-SmHistorique $point 'synthese' $statut ([string]$update.motif)
        $changes += [ordered]@{ id = $point.id; titre = $point.titre; avant = $previous; apres = $statut; motif = $update.motif }
    }

    # Les points réglés depuis plus de 60 jours quittent le registre.
    $limit = (Get-Date).AddDays(-60)
    $kept = New-Object System.Collections.ArrayList
    foreach ($point in $Suivi.points) {
        $updated = ConvertTo-SmDateTime $point.majLe
        if (($point.statut -eq 'traite' -or $point.statut -eq 'ignore') -and $updated -and $updated -lt $limit) { continue }
        [void]$kept.Add($point)
    }
    $Suivi.points = $kept
    , $changes
}

# Marquage d'un point par moi (page de suivi).
function Set-SmSuiviStatut($Suivi, [string]$Id, [string]$Statut, [string]$Note) {
    $point = $Suivi.points | Where-Object { $_.id -eq $Id } | Select-Object -First 1
    if (-not $point) { throw "Point « $Id » introuvable." }
    if ($SmStatuts -notcontains $Statut) { throw "Statut inconnu : $Statut." }
    $changed = $point.statut -ne $Statut
    $point.statut = $Statut
    $point.noteMoi = $Note
    $point.majPar = 'moi'
    $point.majLe = (Get-Date).ToString('o')
    if ($changed) { Add-SmHistorique $point 'moi' $Statut $Note }
    $point
}

# Tableau JSON d'un bloc de fin de synthèse (« <!-- SUIVI --> », « <!-- ARCHIVES --> ») :
# du marqueur jusqu'au marqueur suivant. $null si le bloc est absent.
function Get-SmBlocJson([string]$Text, [string]$Marker) {
    $at = $Text.IndexOf("<!-- $Marker -->")
    if ($at -lt 0) { return $null }
    $segment = $Text.Substring($at + $Marker.Length + 9)
    $next = $segment.IndexOf('<!-- ')
    if ($next -ge 0) { $segment = $segment.Substring(0, $next) }
    $start = $segment.IndexOf('[')
    $end = $segment.LastIndexOf(']')
    if ($start -lt 0 -or $end -le $start) { throw "bloc $Marker sans tableau JSON" }
    $parsed = $segment.Substring($start, $end - $start + 1) | ConvertFrom-Json
    , @($parsed | ForEach-Object { $_ })   # PS 5.1 : un tableau JSON arrive comme un seul objet
}

# --- Registre des éléments à garder -----------------------------------------
# archives.json garde ce qui sert plus tard : variabilité mensuelle, facturation,
# suivi TAM, points d'attention. Chaque élément a un statut :
#   a_traiter  à reporter (variabilité), à traiter (facturation), à voir (attention)
#   fait       reporté / traité / vu
#   ignore     écarté
# « certitude » : etabli (le mail l'affirme) ou a_verifier (indice à contrôler).

$SmArchiveTypes = @('variabilite', 'facturation', 'tam', 'attention')
$SmArchiveStatuts = @('a_traiter', 'fait', 'ignore')

function Get-SmArchives([string]$WorkDir = $SmWorkDir) {
    $saved = Read-SmJson (Join-Path $WorkDir 'archives.json')
    $items = New-Object System.Collections.ArrayList
    if ($saved) {
        foreach ($item in @($saved.elements)) {
            if (-not $item) { continue }
            $copy = [ordered]@{}
            foreach ($property in $item.PSObject.Properties) { $copy[$property.Name] = $property.Value }
            [void]$items.Add($copy)
        }
    }
    [ordered]@{ majLe = $(if ($saved) { $saved.majLe }); elements = $items }
}

function Save-SmArchives($Archives, [string]$WorkDir = $SmWorkDir) {
    $Archives.majLe = (Get-Date).ToString('o')
    Write-SmJson (Join-Path $WorkDir 'archives.json') ([ordered]@{ majLe = $Archives.majLe; elements = @($Archives.elements) })
}

# Intègre les éléments proposés par la synthèse ; renvoie le nombre de nouveaux.
function Merge-SmArchives($Archives, $Updates) {
    $added = 0
    $today = (Get-Date).ToString('yyyyMMdd')
    $prefixes = @{ variabilite = 'V'; facturation = 'F'; tam = 'T'; attention = 'A' }
    foreach ($update in @($Updates)) {
        if (-not $update -or -not $update.element) { continue }
        $type = [string]$update.type
        if ($SmArchiveTypes -notcontains $type) { $type = 'attention' }

        $item = $null
        if ($update.id) { $item = $Archives.elements | Where-Object { $_.id -eq $update.id } | Select-Object -First 1 }
        if (-not $item) {
            $prefix = "$($prefixes[$type])-$today-"
            $n = @($Archives.elements | Where-Object { $_.id -like "$prefix*" }).Count + 1
            $item = [ordered]@{ id = "$prefix$n"; type = $type; statut = 'a_traiter'; majPar = 'synthese'; majLe = (Get-Date).ToString('o');
                                creeLe = (Get-Date).ToString('o'); noteMoi = '' }
            [void]$Archives.elements.Add($item)
            $added++
        }
        foreach ($field in 'certitude', 'client', 'element', 'sens', 'quantite', 'montant', 'dateEffet', 'mois', 'echeance',
                           'source', 'indice', 'verifier', 'dossier', 'ticket') {
            if ($update.PSObject.Properties.Name -contains $field) { $item[$field] = $update.$field }
        }
        # Une suggestion confirmée devient établie ; jamais l'inverse sans le dire.
        if (-not $item.certitude) { $item.certitude = 'etabli' }
    }

    # Les éléments faits ou écartés depuis plus de 120 jours quittent le registre (variabilité : un mois et demi d'historique au moins).
    $limit = (Get-Date).AddDays(-120)
    $kept = New-Object System.Collections.ArrayList
    foreach ($item in $Archives.elements) {
        $updated = ConvertTo-SmDateTime $item.majLe
        if ($item.statut -ne 'a_traiter' -and $updated -and $updated -lt $limit) { continue }
        [void]$kept.Add($item)
    }
    $Archives.elements = $kept
    $added
}

function Set-SmArchiveStatut($Archives, [string]$Id, [string]$Statut, [string]$Note) {
    $item = $Archives.elements | Where-Object { $_.id -eq $Id } | Select-Object -First 1
    if (-not $item) { throw "Élément « $Id » introuvable." }
    if ($SmArchiveStatuts -notcontains $Statut) { throw "Statut inconnu : $Statut." }
    $item.statut = $Statut
    $item.noteMoi = $Note
    $item.majPar = 'moi'
    $item.majLe = (Get-Date).ToString('o')
    $item
}

# --- Markdown -> HTML (le sous-ensemble produit par la synthèse) -----------

function ConvertTo-SmInlineHtml([string]$Text) {
    $html = [Net.WebUtility]::HtmlEncode($Text)
    $html = [regex]::Replace($html, '`([^`]+)`', '<code>$1</code>')
    $html = [regex]::Replace($html, '\*\*(.+?)\*\*', '<strong>$1</strong>')
    $html = [regex]::Replace($html, '\[([^\]]+)\]\((https?://[^)\s]+)\)', '<a href="$2">$1</a>')
    $html
}

function ConvertTo-SmHtml([string]$Markdown) {
    $out = New-Object Text.StringBuilder
    $lines = $Markdown -replace "`r", '' -split "`n"
    $inList = $false
    $inTable = $false
    $headerDone = $false
    foreach ($line in $lines) {
        $isRow = $line -match '^\s*\|.*\|\s*$'
        $isItem = $line -match '^\s*[-*]\s+'
        if ($inList -and -not $isItem) { [void]$out.Append('</ul>'); $inList = $false }
        if ($inTable -and -not $isRow) { [void]$out.Append('</tbody></table>'); $inTable = $false }

        if ($isRow) {
            $cells = ($line.Trim().Trim('|') -split '\|') | ForEach-Object { $_.Trim() }
            if (($cells -join '') -match '^[-:\s]+$') { continue }   # ligne de séparation
            if (-not $inTable) {
                [void]$out.Append('<table><thead><tr>')
                foreach ($cell in $cells) { [void]$out.Append("<th>$(ConvertTo-SmInlineHtml $cell)</th>") }
                [void]$out.Append('</tr></thead><tbody>')
                $inTable = $true
                continue
            }
            [void]$out.Append('<tr>')
            foreach ($cell in $cells) { [void]$out.Append("<td>$(ConvertTo-SmInlineHtml $cell)</td>") }
            [void]$out.Append('</tr>')
        } elseif ($isItem) {
            if (-not $inList) { [void]$out.Append('<ul>'); $inList = $true }
            [void]$out.Append("<li>$(ConvertTo-SmInlineHtml ($line -replace '^\s*[-*]\s+', ''))</li>")
        } elseif ($line -match '^-{3,}\s*$') {
            [void]$out.Append('<hr>')
        } elseif ($line -match '^(#{1,4})\s+(.*)$') {
            $level = $Matches[1].Length
            [void]$out.Append("<h$level>$(ConvertTo-SmInlineHtml $Matches[2])</h$level>")
        } elseif ($line -match '^>\s?(.*)$') {
            [void]$out.Append("<blockquote>$(ConvertTo-SmInlineHtml $Matches[1])</blockquote>")
        } elseif ($line.Trim()) {
            [void]$out.Append("<p>$(ConvertTo-SmInlineHtml $line)</p>")
        }
    }
    if ($inList) { [void]$out.Append('</ul>') }
    if ($inTable) { [void]$out.Append('</tbody></table>') }

    $style = 'body{font-family:Segoe UI,Arial,sans-serif;font-size:10.5pt;color:#1f2328}' +
             'h1{font-size:15pt}h2{font-size:13pt;margin-top:18px}h3{font-size:11pt}' +
             'table{border-collapse:collapse;margin:6px 0}th,td{border:1px solid #d0d7de;padding:4px 6px;vertical-align:top;text-align:left}' +
             'th{background:#f3f4f6}code{font-family:Consolas,monospace;font-size:9.5pt;background:#f3f4f6}' +
             'blockquote{color:#57606a;border-left:3px solid #d0d7de;margin:6px 0;padding-left:8px}'
    "<html><head><meta charset=`"utf-8`"><style>$style</style></head><body>$($out.ToString())</body></html>"
}
