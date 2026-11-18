# --- File « À publier » -------------------------------------------------------------
# Ce qui est prêt à saisir dans Jira (les clôtures préparées, `cloture.md` ou `cloture-panoramix.md` du
# dossier AnsamDiag du ticket), et ce que Jira en montre. Je publie moi-même ; Panoramix vérifie, en
# lecture seule, que le commentaire client, la note interne, les worklogs et le statut sont bien là.
# État local : <WorkDir>\publications.json (une entrée par ticket). Rien n'est jamais écrit dans Jira.

$SmPublierJours = 45   # une clôture plus ancienne ne revient pas dans la file

function Get-SmPublicationsEtat([string]$WorkDir) {
    $map = @{}
    $saved = Read-SmJson (Join-Path $WorkDir 'publications.json')
    foreach ($entry in @($saved.entrees | ForEach-Object { $_ } | Where-Object { $_ -and $_.cle })) {
        $copy = [ordered]@{}
        foreach ($property in $entry.PSObject.Properties) { $copy[$property.Name] = $property.Value }
        $map[[string]$entry.cle] = $copy
    }
    $map
}

function Save-SmPublicationsEtat($Map, [string]$WorkDir) {
    $entries = @($Map.Keys | Sort-Object | ForEach-Object { $Map[$_] })
    Write-SmJson (Join-Path $WorkDir 'publications.json') ([ordered]@{ majLe = (Get-Date).ToString('o'); entrees = $entries })
}

# Type d'une section de clôture, d'après son titre.
function Get-SmTypeBloc([string]$Titre) {
    switch -Regex ($Titre) {
        '(?i)^titre'                    { return 'titre' }
        '(?i)^statut'                   { return 'statut' }
        '(?i)^cat'                      { return 'categorie' }
        '(?i)^résolution|^resolution' { return 'resolution' }
        '(?i)commentaire'               { return 'public' }
        '(?i)note interne'              { return 'interne' }
        '(?i)timesheet|worklog'         { return 'worklog' }
        '(?i)connaissance'              { return 'kb' }
    }
    'autre'
}

# Jours du timesheet (« JJ.MM » puis des puces). L'année est celle qui place le jour avant la clôture.
function ConvertFrom-SmTimesheet([string]$Texte, [datetime]$Fichier) {
    $days = New-Object System.Collections.ArrayList
    $current = $null
    $notes = New-Object System.Collections.ArrayList
    foreach ($line in ($Texte -split "`r?`n")) {
        if ($line -match '^\s*(\d{1,2})\.(\d{1,2})\.?\s*$') {
            $year = $Fichier.Year
            $date = $null
            try { $date = Get-Date -Year $year -Month ([int]$Matches[2]) -Day ([int]$Matches[1]) -Hour 0 -Minute 0 -Second 0 } catch { continue }
            if ($date -gt $Fichier.AddDays(2)) { $date = $date.AddYears(-1) }
            $current = [ordered]@{ jour = $date.ToString('yyyy-MM-dd'); libelle = $date.ToString('dd.MM'); lignes = New-Object System.Collections.ArrayList }
            [void]$days.Add($current)
        } elseif ($line -match '^\s*[-*]\s+(.+)$' -and $current) {
            [void]$current.lignes.Add($Matches[1].Trim())
        } elseif ($line.Trim()) {
            [void]$notes.Add($line.Trim())
        }
    }
    [ordered]@{
        jours = @($days | ForEach-Object { [ordered]@{ jour = $_.jour; libelle = $_.libelle; texte = (@($_.lignes | ForEach-Object { "- $_" }) -join "`n") } })
        note = ($notes -join ' ')
    }
}

function ConvertFrom-SmCloture([string]$Texte, [datetime]$Fichier) {
    $parts = [regex]::Split($Texte.TrimStart([char]0xFEFF), '(?m)^###\s+')
    $head = ($parts[0] -replace '(?m)^##\s+.*$', '').Trim()
    $blocks = @(for ($i = 1; $i -lt $parts.Count; $i++) {
        $at = $parts[$i].IndexOf("`n")
        $title = $(if ($at -lt 0) { $parts[$i] } else { $parts[$i].Substring(0, $at) }).Trim()
        $body = $(if ($at -lt 0) { '' } else { $parts[$i].Substring($at + 1) }).Trim()
        $type = Get-SmTypeBloc $title
        $block = [ordered]@{ titre = $title; type = $type; texte = $body }
        if ($type -eq 'worklog') {
            $sheet = ConvertFrom-SmTimesheet $body $Fichier
            $block.jours = $sheet.jours
            $block.note = $sheet.note
        }
        $block
    })
    [ordered]@{ avant = $head; blocs = $blocks }
}

# La file : une entrée par ticket dont la clôture a moins de $SmPublierJours jours (ou n'est pas encore vérifiée).
function Get-SmPublications($Config, [string]$WorkDir, $Suivi) {
    $root = Join-Path ([string]$Config.dossierAnsamDiag) 'Tickets'
    if (-not (Test-Path $root)) { return @() }
    $etat = Get-SmPublicationsEtat $WorkDir
    # Depuis le début de la période pas encore facturée : un mois clos dans l'onglet Facturation
    # (Facturation\<AAAA-MM>\cloture.json) ne revient plus ici. $SmPublierJours jours au plus.
    $limit = (Get-Date).AddDays(-$SmPublierJours)
    $month = (Get-Date -Day 1).Date
    $previous = $month.AddMonths(-1)
    $closed = Test-Path (Join-Path ([string]$Config.dossierAnsamDiag) ('Facturation\{0:yyyy-MM}\cloture.json' -f $previous))
    $start = $(if ($closed) { $month } else { $previous })
    if ($start -gt $limit) { $limit = $start }
    $out = New-Object System.Collections.ArrayList
    foreach ($folder in Get-ChildItem $root -Directory -ErrorAction SilentlyContinue) {
        $file = @('cloture.md', 'cloture-panoramix.md' | ForEach-Object { Get-Item (Join-Path $folder.FullName $_) -ErrorAction SilentlyContinue } |
            Where-Object { $_ }) | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if (-not $file) { continue }
        $cle = $folder.Name
        $state = $etat[$cle]
        $written = $file.LastWriteTime.ToString('o')
        $same = $state -and [string]$state.fichierLe -eq $written
        if ($file.LastWriteTime -lt $limit -and -not ($same -and $state.publieLe -and -not ($state.verification -and $state.verification.complet))) { continue }
        if ($same -and $state.ecarteLe) { continue }
        $parsed = ConvertFrom-SmCloture ([IO.File]::ReadAllText($file.FullName, [Text.Encoding]::UTF8)) $file.LastWriteTime
        $point = $Suivi.points | Where-Object { -not $_.estGroupe -and ($_.ticket -eq $cle -or $_.id -eq $cle) } | Sort-Object { [string]$_.majLe } -Descending | Select-Object -First 1
        $titleBlock = @($parsed.blocs | Where-Object { $_.type -eq 'titre' }) | Select-Object -First 1
        [void]$out.Add([ordered]@{
            cle = $cle; fichier = $file.Name; fichierLe = $written
            titre = $(if ($point) { $point.titre } elseif ($titleBlock) { $titleBlock.texte } else { $cle })
            client = $(if ($point) { $point.client }); pointId = $(if ($point) { $point.id }); pointStatut = $(if ($point) { $point.statut })
            jiraStatut = $(if ($point) { $point.jiraStatut })
            avant = $parsed.avant; blocs = @($parsed.blocs)
            publieLe = $(if ($same) { $state.publieLe }); verification = $(if ($same) { $state.verification })
        })
    }
    @($out | Sort-Object { [string]$_.fichierLe } -Descending)
}

function Set-SmPublication([string]$WorkDir, [string]$Cle, [string]$FichierLe, [string]$Action) {
    $etat = Get-SmPublicationsEtat $WorkDir
    $entry = $etat[$Cle]
    if (-not $entry -or [string]$entry.fichierLe -ne $FichierLe) { $entry = [ordered]@{ cle = $Cle; fichierLe = $FichierLe } }
    switch ($Action) {
        'publie'  { $entry.publieLe = (Get-Date).ToString('o') }
        'ecarter' { $entry.ecarteLe = (Get-Date).ToString('o') }
        'rouvrir' { $entry.Remove('publieLe'); $entry.Remove('ecarteLe'); $entry.Remove('verification') }
        default   { throw "Action inconnue : $Action" }
    }
    $etat[$Cle] = $entry
    Save-SmPublicationsEtat $etat $WorkDir
    $entry
}

# --- Vérification dans Jira (lecture seule) -----------------------------------------
function ConvertTo-SmMots([string]$Texte) {
    $plain = $Texte.ToLowerInvariant().Normalize([Text.NormalizationForm]::FormD) -replace '\p{Mn}', '' -replace '[^a-z0-9]+', ' '
    @($plain.Split(' ') | Where-Object { $_.Length -ge 4 } | Sort-Object -Unique)
}

# Part des mots du bloc retrouvés dans le commentaire : le TAM retouche souvent le texte avant de publier.
function Get-SmRessemblance([string]$Bloc, [string]$Commentaire) {
    $words = ConvertTo-SmMots $Bloc
    if (-not $words.Count) { return 0 }
    $set = @{}
    foreach ($w in (ConvertTo-SmMots $Commentaire)) { $set[$w] = $true }
    [math]::Round(@($words | Where-Object { $set.ContainsKey($_) }).Count / $words.Count, 2)
}

function Test-SmPublication($Acces, $Publication, [string]$MonCompte) {
    $cle = [string]$Publication.cle
    $since = (ConvertTo-SmDateTime $Publication.fichierLe).AddDays(-3)
    $issue = Invoke-SmJira $Acces "/rest/api/3/issue/${cle}?fields=status,resolution"
    $comments = @((Invoke-SmJira $Acces "/rest/api/3/issue/$cle/comment?orderBy=-created&maxResults=100").comments | Where-Object { $_ })
    # Chez ANSAM, le message de résolution au client est publié par « Automation for Jira » à la transition
    # (le texte que je saisis) ; mes notes internes sont courtes. On cherche donc le message au client parmi
    # mes commentaires et ceux de l'automatisation, et la note interne parmi les miens seulement.
    $recent = @($comments | Where-Object { (ConvertTo-SmDateTime $_.created) -ge $since } |
        ForEach-Object { [pscustomobject]@{ le = $_.created; moi = ($_.author.accountId -eq $MonCompte); auto = ([string]$_.author.displayName -eq 'Automation for Jira')
                                            public = [bool]$_.jsdPublic; texte = (Format-SmCommentaireJira $_.body) } })
    $quand = { param($c) ([string]$c.le).Substring(0, 16).Replace('T', ' ') }
    $checks = New-Object System.Collections.ArrayList
    foreach ($block in @($Publication.blocs)) {
        switch ($block.type) {
            'public' {
                # Un message au client après la préparation : le mien, ou celui de la transition (automatisation).
                $candidates = @($recent | Where-Object { ($_.moi -or $_.auto) -and ($_.texte -match '^(?i)(bonjour|hello|madame|monsieur|cher|chère)' -or (Get-SmRessemblance $block.texte $_.texte) -ge 0.2) } |
                    Where-Object { $_.texte -notmatch "(?i)^la requête est passée" })
                $best = $candidates | Sort-Object { Get-SmRessemblance $block.texte $_.texte } -Descending | Select-Object -First 1
                $detail = $(if (-not $best) { 'aucun message au client dans Jira depuis la préparation' }
                            else { "message publié le $(& $quand $best)$(if ($best.auto) { ' (à la résolution)' } else { '' }), ressemblance $([math]::Round((Get-SmRessemblance $block.texte $best.texte) * 100)) %" })
                [void]$checks.Add([ordered]@{ type = 'public'; titre = $block.titre; ok = [bool]$best; bloquant = $true; detail = $detail })
            }
            'interne' {
                $note = @($recent | Where-Object { $_.moi }) | Select-Object -First 1
                [void]$checks.Add([ordered]@{ type = 'interne'; titre = $block.titre; ok = [bool]$note; bloquant = $false
                                              detail = $(if ($note) { "note de moi le $(& $quand $note)" } else { 'aucune note de moi depuis la préparation (conseillé, pas obligatoire)' }) })
            }
            'worklog' {
                $days = @($block.jours | Where-Object { $_ })
                if (-not $days.Count) { break }
                $first = [datetime]::ParseExact(($days | Sort-Object { $_.jour } | Select-Object -First 1).jour, 'yyyy-MM-dd', $null).AddDays(-1)
                $ms = [long]([DateTimeOffset]$first).ToUnixTimeMilliseconds()
                $logs = @((Invoke-SmJira $Acces "/rest/api/3/issue/$cle/worklog?startedAfter=$ms&maxResults=500").worklogs | Where-Object { $_ -and $_.author.accountId -eq $MonCompte })
                $logged = @{}
                foreach ($log in $logs) { $logged[([string]$log.started).Substring(0, 10)] = [int]$logged[([string]$log.started).Substring(0, 10)] + [int]$log.timeSpentSeconds }
                foreach ($day in $days) {
                    $sec = [int]$logged[[string]$day.jour]
                    [void]$checks.Add([ordered]@{ type = 'worklog'; jour = $day.jour; titre = "Worklog du $($day.libelle)"; ok = $sec -gt 0; bloquant = $true
                                                  detail = $(if ($sec) { "saisi : $([math]::Round($sec / 60)) min" } else { "aucun worklog de moi ce jour-là$(if ($logged.Count) { ' (saisis : ' + ((@($logged.Keys) | Sort-Object | ForEach-Object { $_.Substring(8, 2) + '.' + $_.Substring(5, 2) }) -join ', ') + ')' })" }) })
                }
            }
            'statut' {
                # Le statut proposé est le début de la ligne (« Résolu — … », « Résolu, d'après … », « En attente du client (…) »).
                $wanted = (([string]$block.texte -split '[\r\n]')[0] -split '\s+[—–-]\s+|[,;:(]')[0].Trim().TrimEnd('.')
                $name = [string]$issue.fields.status.name
                $done = [string]$issue.fields.status.statusCategory.key -eq 'done'
                $n = { param($s) ($s.ToLowerInvariant().Normalize([Text.NormalizationForm]::FormD) -replace '\p{Mn}', '').Trim() }
                $ok = ((& $n $name).Contains((& $n $wanted))) -or ($done -and (& $n $wanted) -match 'resolu|clos|termine|ferme')
                [void]$checks.Add([ordered]@{ type = 'statut'; titre = 'Statut'; ok = [bool]$ok; bloquant = $true; detail = "Jira : $name$(if ($issue.fields.resolution) { " ($($issue.fields.resolution.name))" }) ; attendu : $wanted" })
            }
        }
    }
    [ordered]@{
        le = (Get-Date).ToString('o'); jiraStatut = [string]$issue.fields.status.name
        complet = [bool]($checks.Count -and -not @($checks | Where-Object { -not $_.ok -and $_.bloquant }).Count)
        manques = @($checks | Where-Object { -not $_.ok -and $_.bloquant }).Count
        controles = @($checks)
    }
}
