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
        reveilPC            = $true              # réveiller le PC de sa veille à l'heure prévue (sur secteur)
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
# Les groupes (estGroupe) rassemblent des sous-tâches (parent) : voir « Groupes de points ».

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
        if ($point -and $point.estGroupe) { continue }   # l'état d'un groupe se déduit de ses sous-tâches
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
        # Une prise de contact rédigée avec tout le dossier (redigeLe) tient tant qu'aucun message
        # n'est arrivé après elle ; sinon le brouillon de la synthèse la remplace.
        if ($update.PSObject.Properties.Name -contains 'contact') {
            $written = $(if ($point.contact -and $point.contact.redigeLe) { ConvertTo-SmDateTime $point.contact.redigeLe })
            $lastMessage = ConvertTo-SmDateTime $update.dernierMessage
            $keep = $written -and $update.contact -and -not ($lastMessage -and $lastMessage -gt $written)
            if (-not $keep) { $point.contact = $update.contact }
        }

        # Regroupement proposé par la synthèse ; un groupement fait ou défait par moi prime.
        if ($update.groupe -and $point.parentPar -ne 'moi') {
            $group = Get-SmGroupe $Suivi ([string]$update.groupe)
            if (-not $group) { $group = New-SmGroupe $Suivi ([string]$update.groupe) 'synthese' }
            $point.parent = $group.id
            $point.parentPar = 'synthese'
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
        if (-not $point.estGroupe -and ($point.statut -eq 'traite' -or $point.statut -eq 'ignore') -and $updated -and $updated -lt $limit) { continue }
        [void]$kept.Add($point)
    }
    $Suivi.points = $kept
    Update-SmGroupes $Suivi
    , $changes
}

# Marquage par moi (page de suivi). Sur un groupe, le statut s'applique à toutes ses sous-tâches.
function Set-SmSuiviStatut($Suivi, [string]$Id, [string]$Statut, [string]$Note) {
    $point = $Suivi.points | Where-Object { $_.id -eq $Id } | Select-Object -First 1
    if (-not $point) { throw "Point « $Id » introuvable." }
    if ($SmStatuts -notcontains $Statut) { throw "Statut inconnu : $Statut." }
    if ($point.estGroupe) {
        # Seules les sous-tâches encore ouvertes changent : « tout marquer traité » ne rouvre rien.
        foreach ($child in @($Suivi.points | Where-Object { $_.parent -eq $Id -and ($_.statut -eq 'a_faire' -or $_.statut -eq 'attente_tiers') })) {
            [void](Set-SmSuiviStatut $Suivi $child.id $Statut $child.noteMoi)
        }
        $point.noteMoi = $Note
        Update-SmGroupes $Suivi
        return $point
    }
    $changed = $point.statut -ne $Statut
    $point.statut = $Statut
    $point.noteMoi = $Note
    $point.majPar = 'moi'
    $point.majLe = (Get-Date).ToString('o')
    if ($changed) { Add-SmHistorique $point 'moi' $Statut $Note }
    Update-SmGroupes $Suivi
    $point
}

# --- Groupes de points --------------------------------------------------------
# Un groupe (estGroupe = $true, id G-AAAAMMJJ-n) rassemble des sous-tâches (parent = id du
# groupe). Son statut, sa priorité et son échéance se déduisent des sous-tâches ouvertes.
# parentPar dit qui a rangé le point : 'moi' (jamais déplacé par la synthèse) ou 'synthese'.

function ConvertTo-SmCle([string]$Text) {
    ($Text.ToLowerInvariant().Normalize([Text.NormalizationForm]::FormD) -replace '\p{Mn}', '' -replace '[^a-z0-9]+', ' ').Trim()
}

function Get-SmGroupe($Suivi, [string]$Titre) {
    $key = ConvertTo-SmCle $Titre
    $Suivi.points | Where-Object { $_.estGroupe -and (ConvertTo-SmCle $_.titre) -eq $key } | Select-Object -First 1
}

function New-SmGroupe($Suivi, [string]$Titre, [string]$Par) {
    $prefix = 'G-' + (Get-Date).ToString('yyyyMMdd') + '-'
    $n = @($Suivi.points | Where-Object { $_.id -like "$prefix*" }).Count + 1
    while ($Suivi.points | Where-Object { $_.id -eq "$prefix$n" }) { $n++ }
    $group = [ordered]@{
        id = "$prefix$n"; estGroupe = $true; titre = $Titre.Trim(); creePar = $Par; creeLe = (Get-Date).ToString('o')
        statut = 'a_faire'; majPar = $Par; majLe = (Get-Date).ToString('o'); noteMoi = ''; historique = @()
    }
    [void]$Suivi.points.Add($group)
    $group
}

# Range des points dans un groupe (ou les sort de tout groupe avec $GroupeId vide), par moi.
function Set-SmParent($Suivi, [string[]]$Ids, [string]$GroupeId) {
    foreach ($id in $Ids) {
        $point = $Suivi.points | Where-Object { $_.id -eq $id -and -not $_.estGroupe } | Select-Object -First 1
        if (-not $point) { continue }
        $point.parent = $(if ($GroupeId) { $GroupeId } else { $null })
        $point.parentPar = 'moi'
    }
    Update-SmGroupes $Suivi
}

# --- Synchronisation avec Jira -------------------------------------------------
# Déterministe (sans l'agent) : chaque ticket Jira ouvert qui m'est assigné a un point
# dans le registre, et un ticket encore ouvert dans Jira n'est pas « réglé » en silence.

function Get-SmStatutJira([string]$JiraStatus) {
    if ($JiraStatus -match 'R[ée]solu|Resolved|Clos|Closed|Termin|Done|Annul|Cancel') { return 'traite' }
    if ($JiraStatus -match 'attente|Waiting|Awaiting|Pending|Planifi|Scheduled|Suspend') { return 'attente_tiers' }
    'a_faire'
}

function Get-SmPrioriteJira([string]$JiraPriority) {
    if ($JiraPriority -match 'P1|P2|Highest|High|Haute|Critique|Urgent') { return 'rouge' }
    if ($JiraPriority -match 'P4|P5|Low|Basse|Mineure') { return 'vert' }
    'orange'
}

# Client d'un ticket : dossier AnsamDiag du ticket, domaine du rapporteur, puis projet Jira (clé ou
# nom), d'après les fiches clients. Sinon vide : la synthèse le complète à partir des mails.
function Resolve-SmClientJira($Ticket, $Context) {
    $local = @($Context.tickets | ForEach-Object { $_ } | Where-Object { $_ -and $_.cle -eq $Ticket.cle -and $_.client }) | Select-Object -First 1
    if ($local) { return $local.client }   # dossier AnsamDiag du ticket
    $clients = @($Context.clients | ForEach-Object { $_ } | Where-Object { $_ })
    # Organisation Service Desk du ticket : son nom ressemble au nom complet de la fiche client.
    foreach ($organisation in @($Ticket.organisations | ForEach-Object { $_ } | Where-Object { $_ })) {
        $key = ConvertTo-SmCle ([string]$organisation)
        $hit = $clients | Where-Object { $_.code -and ((ConvertTo-SmCle ([string]$_.code)) -eq $key -or
            ($_.nom -and ((ConvertTo-SmCle ([string]$_.nom)).StartsWith($key) -or $key.StartsWith((ConvertTo-SmCle ([string]$_.nom)))))) } | Select-Object -First 1
        if ($hit) { return $hit.code }
    }
    $reporter = $(if ($Ticket.rapporteurEmail) { [string]$Ticket.rapporteurEmail } else { [string]$Ticket.rapporteur })
    if ($reporter -match '@([A-Za-z0-9.-]+)$') {
        $domain = $Matches[1].ToLowerInvariant()
        $hit = $clients | Where-Object { @($_.domaines) -contains $domain } | Select-Object -First 1
        if ($hit) { return $hit.code }
    }
    $prefix = ([string]$Ticket.cle -split '-')[0]
    if ($prefix -ne 'ASD') {
        $hit = $clients | Where-Object { $_.code -and ($prefix -eq $_.code -or $prefix -like "*$($_.code)") } | Select-Object -First 1
        if ($hit) { return $hit.code }
        $project = ConvertTo-SmCle ([string]$Ticket.projet)
        $hit = $clients | Where-Object { $_.nom -and $project -and (ConvertTo-SmCle $_.nom).StartsWith(($project -split ' ')[0]) } | Select-Object -First 1
        if ($hit) { return $hit.code }
    }
    # Client sans fiche AnsamDiag : l'organisation Jira, sinon le domaine du rapporteur, plutôt que rien.
    $organisation = @($Ticket.organisations | ForEach-Object { $_ } | Where-Object { $_ }) | Select-Object -First 1
    if ($organisation) { return [string]$organisation }
    if ($reporter -match '@([A-Za-z0-9.-]+)$') { return $Matches[1].ToLowerInvariant() }
    $null
}

function Sync-SmJira($Suivi, $Jira, $Context) {
    $report = [ordered]@{ ajoutes = 0; regles = 0; rouverts = 0 }
    if (-not $Jira) { return $report }
    $tickets = @($Jira.tickets | ForEach-Object { $_ } | Where-Object { $_ -and $_.cle })
    $now = (Get-Date).ToString('o')
    $newByProject = @{}

    foreach ($ticket in $tickets) {
        $mapped = Get-SmStatutJira $ticket.statut
        $updated = ConvertTo-SmDateTime $ticket.misAJour
        $updatedText = $(if ($updated) { $updated.ToString('dd.MM') } else { '?' })
        $candidates = @($Suivi.points | Where-Object { -not $_.estGroupe -and ($_.ticket -eq $ticket.cle -or $_.id -eq $ticket.cle) })
        $point = @($candidates | Where-Object { $_.statut -eq 'a_faire' -or $_.statut -eq 'attente_tiers' }) + $candidates | Select-Object -First 1

        if ($point) {
            $point.jiraStatut = $ticket.statut
            $point.jiraMisAJour = $ticket.misAJour
            if (-not $point.client) { $found = Resolve-SmClientJira $ticket $Context; if ($found) { $point.client = $found } }
            $open = $point.statut -eq 'a_faire' -or $point.statut -eq 'attente_tiers'
            $point.jiraEncoreOuvert = (-not $open) -and $mapped -ne 'traite'
            if ($point.majPar -eq 'moi') { continue }   # mon marquage prime ; jiraEncoreOuvert le signale
            if ($mapped -eq 'traite' -and $open) {
                $point.statut = 'traite'; $point.majPar = 'synthese'; $point.majLe = $now
                Add-SmHistorique $point 'synthese' 'traite' "Jira « $($ticket.statut) »"
                $report.regles++
            } elseif ($mapped -ne 'traite' -and -not $open) {
                $point.statut = 'a_faire'; $point.majPar = 'synthese'; $point.majLe = $now; $point.jiraEncoreOuvert = $false
                $point.action = "Réglé d'après les mails, mais encore « $($ticket.statut) » dans Jira : clôturer le ticket ou le reprendre."
                Add-SmHistorique $point 'synthese' 'a_faire' "Encore « $($ticket.statut) » dans Jira"
                $report.rouverts++
            }
            continue
        }
        if ($mapped -eq 'traite') { continue }   # résolu, en attente de clôture : rien à faire

        $id = [string]$ticket.cle
        if ($Suivi.points | Where-Object { $_.id -eq $id }) { $id = "$id-J" }
        $action = $(if ($mapped -eq 'attente_tiers') { "En attente dans Jira (« $($ticket.statut) ») : relancer si rien ne bouge." }
                    else { "Faire avancer le ticket : aucun mail récent, dernière mise à jour Jira le $updatedText." })
        $point = [ordered]@{
            id = $id; ticket = $ticket.cle; titre = $ticket.titre; client = (Resolve-SmClientJira $ticket $Context)
            statut = $mapped; priorite = (Get-SmPrioriteJira $ticket.priorite); echeance = $ticket.echeance; action = $action
            source = "Jira · $($ticket.statut) · mis à jour le $updatedText"; dernierMessage = $null; origine = 'jira'
            jiraStatut = $ticket.statut; jiraMisAJour = $ticket.misAJour; creeLe = $now; majPar = 'synthese'; majLe = $now; noteMoi = ''
            historique = @([ordered]@{ le = $now; par = 'synthese'; statut = $mapped; texte = "Ajouté depuis Jira (« $($ticket.statut) »)" })
        }
        [void]$Suivi.points.Add($point)
        $report.ajoutes++
        if ($ticket.projet -and $ticket.projet -notmatch 'Service Desk') { $newByProject[[string]$ticket.projet] = $true }
    }

    # Tâches d'un même projet (hors Service Desk) : regroupées sous le nom du projet dès trois tickets ouverts.
    foreach ($project in @($newByProject.Keys)) {
        $keys = @($tickets | Where-Object { $_.projet -eq $project } | ForEach-Object { [string]$_.cle })
        $ids = @($Suivi.points | Where-Object {
            -not $_.estGroupe -and $keys -contains $_.ticket -and -not $_.parent -and $_.parentPar -ne 'moi' -and
            ($_.statut -eq 'a_faire' -or $_.statut -eq 'attente_tiers') } | ForEach-Object { $_.id })
        if ($ids.Count -lt 3) { continue }
        $group = Get-SmGroupe $Suivi $project
        if (-not $group) { $group = New-SmGroupe $Suivi $project 'synthese' }
        foreach ($id in $ids) {
            $p = $Suivi.points | Where-Object { $_.id -eq $id } | Select-Object -First 1
            $p.parent = $group.id; $p.parentPar = 'synthese'
        }
    }
    Update-SmGroupes $Suivi
    $report
}

# Recalcule chaque groupe depuis ses sous-tâches ; supprime les groupes vides.
function Update-SmGroupes($Suivi) {
    $groupIds = @($Suivi.points | Where-Object { $_.estGroupe } | ForEach-Object { $_.id })
    foreach ($point in @($Suivi.points | Where-Object { -not $_.estGroupe -and $_.parent })) {
        if ($groupIds -notcontains $point.parent) { $point.parent = $null }
    }
    $rank = @{ rouge = 0; orange = 1; vert = 2 }
    $kept = New-Object System.Collections.ArrayList
    foreach ($point in $Suivi.points) {
        if (-not $point.estGroupe) { [void]$kept.Add($point); continue }
        $children = @($Suivi.points | Where-Object { -not $_.estGroupe -and $_.parent -eq $point.id })
        if (-not $children.Count) { continue }
        $open = @($children | Where-Object { $_.statut -eq 'a_faire' -or $_.statut -eq 'attente_tiers' })
        $point.statut = $(if ($children | Where-Object { $_.statut -eq 'a_faire' }) { 'a_faire' }
                          elseif ($open.Count) { 'attente_tiers' }
                          elseif (-not ($children | Where-Object { $_.statut -ne 'ignore' })) { 'ignore' }
                          else { 'traite' })
        $best = $open | Where-Object { $_.priorite } | Sort-Object { $rank[[string]$_.priorite] } | Select-Object -First 1
        $point.priorite = $(if ($best) { $best.priorite } else { $null })
        $due = $open | Where-Object { $_.echeance } | ForEach-Object { [string]$_.echeance } | Sort-Object | Select-Object -First 1
        $point.echeance = $(if ($due) { $due } else { $null })
        $point.clients = @($children | Where-Object { $_.client } | ForEach-Object { [string]$_.client } | Sort-Object -Unique)
        $point.nbTotal = $children.Count
        $point.nbOuverts = $open.Count
        [void]$kept.Add($point)
    }
    $Suivi.points = $kept
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
