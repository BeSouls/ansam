<#
.SYNOPSIS
    Ouvre la page locale de suivi et de configuration de la synthèse mails.

.DESCRIPTION
    Petit serveur web sur 127.0.0.1 (jamais accessible depuis le réseau), au
    port de la configuration (8765 par défaut). Chaque requête doit porter le
    jeton gardé dans <WorkDir>\page.token : l'adresse reste la même, on peut
    la mettre en favori, et chaque synthèse envoyée contient le lien.

    -Service : mode permanent, lancé à l'ouverture de session par la tâche
    « Synthese mails ANSAM - page » (installée avec la tâche du matin) ; ni
    navigateur ni arrêt automatique.
    Sans -Service : si le service tourne déjà, ouvre simplement la page ;
    sinon démarre le serveur, ouvre la page, et s'arrête sur « Arrêter » ou
    après 30 minutes sans activité.

    Onglet « Suivi » : les points du registre (<WorkDir>\suivi.json), dont
    je change le statut (à faire, en attente d'un tiers, traité, sans objet)
    avec une note ; la synthèse suivante en tient compte.

    Vues « Variabilité », « Facturation » et « Attention » : les éléments à
    garder (<WorkDir>\archives.json), établis ou à vérifier, que je marque
    reportés / traités / vus ou écartés.

    Onglet « Réglages » : <WorkDir>\config.json : synthèse active, heure, jours
    d'envoi, fenêtre en jours, envoi par mail, absences, dossiers. Elle
    affiche la prochaine synthèse prévue, l'état de la tâche planifiée et les
    dernières synthèses, et peut installer ou mettre à jour la tâche.

.EXAMPLE
    .\Configurer-SyntheseMails.ps1      # ou double-clic sur Configurer-SyntheseMails.cmd
    .\Configurer-SyntheseMails.ps1 -Service
#>
[CmdletBinding()]
param(
    [string]$WorkDir = (Join-Path $env:USERPROFILE 'SyntheseMails'),
    [string]$TaskName = 'Synthese mails ANSAM',
    [int]$IdleMinutes = 30,
    # Mode permanent (tâche à l'ouverture de session).
    [switch]$Service,
    # Affiche l'adresse sans ouvrir le navigateur.
    [switch]$PasDeNavigateur
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'SyntheseMails.Commun.ps1')

$utf8 = New-Object Text.UTF8Encoding($false)
$pagePath = Join-Path $PSScriptRoot 'config.html'
$port = [int](Get-SmConfig $WorkDir).portPage
$token = Get-SmPageToken $WorkDir
$url = "http://127.0.0.1:$port/?t=$token"

function Test-PageRunning {
    $probe = New-Object Net.Sockets.TcpClient
    try { $probe.Connect([Net.IPAddress]::Loopback, $port); $true } catch { $false } finally { $probe.Close() }
}

# L'état des tâches planifiées coûte plusieurs secondes (CIM) : calculé au démarrage puis
# gardé 5 minutes, et recalculé après une installation.
$taskCache = $null
$taskCacheAt = [datetime]::MinValue
function Get-TaskState([switch]$Refresh) {
    if ($Refresh -or -not $script:taskCache -or ((Get-Date) - $script:taskCacheAt).TotalMinutes -ge 5) {
        $script:taskCache = Read-TaskState
        $script:taskCacheAt = Get-Date
    }
    $script:taskCache
}

function Read-TaskState {
    $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    if (-not $task) { return [ordered]@{ installee = $false } }
    $info = Get-ScheduledTaskInfo -TaskName $TaskName
    $trigger = @($task.Triggers)[0]
    [ordered]@{
        installee          = $true
        pagePermanente     = [bool](Get-ScheduledTask -TaskName "$TaskName - page" -ErrorAction SilentlyContinue)
        heure              = $(if ($trigger.StartBoundary) { ([datetime]$trigger.StartBoundary).ToString('HH:mm') })
        prochaineExecution = $(if ($info.NextRunTime) { $info.NextRunTime.ToString('yyyy-MM-ddTHH:mm') })
        derniereExecution  = $(if ($info.LastRunTime -and $info.LastRunTime.Year -gt 2000) { $info.LastRunTime.ToString('yyyy-MM-ddTHH:mm') })
        dernierResultat    = $info.LastTaskResult
    }
}

function Get-State {
    $config = Get-SmConfig $WorkDir
    $etat = Get-SmEtat $WorkDir
    $next = Get-SmProchaine $config $etat
    $nextInfo = $null
    if ($next) {
        $nextInfo = [ordered]@{
            moment   = $next.moment.ToString('yyyy-MM-ddTHH:mm')
            depuis   = $(if ($next.plan.depuis) { $next.plan.depuis.ToString('yyyy-MM-ddTHH:mm') })
            jours    = $next.plan.jours
            absences = @($next.plan.absences)
        }
    }
    $synthesisDir = Join-Path $WorkDir 'syntheses'
    $files = @()
    if (Test-Path $synthesisDir) {
        $files = @(Get-ChildItem $synthesisDir -Filter '*.md' | Sort-Object Name -Descending |
            Select-Object -First 10 | ForEach-Object { $_.BaseName })
    }
    [ordered]@{
        config     = $config
        etat       = $etat
        prochaine  = $nextInfo
        tache      = Get-TaskState
        syntheses  = $files
        dossier    = $WorkDir
        service    = $Service.IsPresent
    }
}

function Read-Request($stream) {
    $buffer = New-Object byte[] 8192
    $memory = New-Object IO.MemoryStream
    $headerEnd = -1
    while ($headerEnd -lt 0) {
        $read = $stream.Read($buffer, 0, $buffer.Length)
        if ($read -le 0) { return $null }
        $memory.Write($buffer, 0, $read)
        # Décodage ASCII : un octet = un caractère, les positions restent alignées.
        $headerEnd = [Text.Encoding]::ASCII.GetString($memory.ToArray()).IndexOf("`r`n`r`n")
        if ($headerEnd -lt 0 -and $memory.Length -gt 65536) { return $null }
    }
    $bytes = $memory.ToArray()
    $lines = [Text.Encoding]::ASCII.GetString($bytes, 0, $headerEnd) -split "`r`n"
    $first = $lines[0] -split ' '
    $headers = @{}
    foreach ($line in ($lines | Select-Object -Skip 1)) {
        $colon = $line.IndexOf(':')
        if ($colon -gt 0) { $headers[$line.Substring(0, $colon).Trim().ToLower()] = $line.Substring($colon + 1).Trim() }
    }
    $length = 0
    if ($headers['content-length']) { $length = [int]$headers['content-length'] }
    if ($length -gt 1MB) { return $null }
    $bodyStart = $headerEnd + 4
    while ($bytes.Length - $bodyStart -lt $length) {
        $read = $stream.Read($buffer, 0, $buffer.Length)
        if ($read -le 0) { break }
        $memory.Write($buffer, 0, $read)
        $bytes = $memory.ToArray()
    }
    $target = $first[1]
    $query = @{}
    $path = $target
    if ($target -match '^([^?]*)\?(.*)$') {
        $path = $Matches[1]
        foreach ($pair in ($Matches[2] -split '&')) {
            $kv = $pair -split '=', 2
            if ($kv.Count -eq 2) { $query[$kv[0]] = [Uri]::UnescapeDataString($kv[1]) }
        }
    }
    [pscustomobject]@{
        Method  = $first[0]
        Path    = $path
        Query   = $query
        Headers = $headers
        Body    = $utf8.GetString($bytes, $bodyStart, [math]::Min($length, $bytes.Length - $bodyStart))
    }
}

function Send-Response($stream, [int]$Status, [string]$ContentType, [string]$Body) {
    $reasons = @{ 200 = 'OK'; 400 = 'Bad Request'; 403 = 'Forbidden'; 404 = 'Not Found'; 500 = 'Internal Server Error' }
    $payload = $utf8.GetBytes($Body)
    $head = "HTTP/1.1 $Status $($reasons[$Status])`r`nContent-Type: $ContentType; charset=utf-8`r`n" +
            "Content-Length: $($payload.Length)`r`nCache-Control: no-store`r`nConnection: close`r`n`r`n"
    $headBytes = [Text.Encoding]::ASCII.GetBytes($head)
    $stream.Write($headBytes, 0, $headBytes.Length)
    $stream.Write($payload, 0, $payload.Length)
}

function Send-Json($stream, [int]$Status, $Value) {
    Send-Response $stream $Status 'application/json' ($Value | ConvertTo-Json -Depth 6)
}

# Construit la configuration à partir de ce que la page envoie (seulement les clés connues).
function ConvertFrom-PageConfig([string]$Json) {
    $posted = $Json | ConvertFrom-Json
    $config = Get-SmConfig $WorkDir
    foreach ($key in @($config.Keys)) {
        if ($posted.PSObject.Properties.Name -contains $key) { $config[$key] = $posted.$key }
    }
    $config.actif = [bool]$config.actif
    $config.envoiMail = [bool]$config.envoiMail
    $config.fenetreJours = [int]$config.fenetreJours
    $config.joursEnvoi = @($config.joursEnvoi | ForEach-Object { [int]$_ } | Sort-Object -Unique)
    $config.absences = @(@($config.absences) | Where-Object { $_ } | ForEach-Object {
        [ordered]@{ debut = [string]$_.debut; fin = [string]$_.fin; libelle = ([string]$_.libelle).Trim() }
    } | Sort-Object { $_.debut })
    $config
}

# Lancement à la demande (bouton de la page) : même lanceur que la tâche du matin, avec -Force.
function Get-RunState {
    $lock = Read-SmJson (Join-Path $WorkDir 'lancement.lock')
    $running = [bool]($lock -and (Get-Process -Id $lock.pid -ErrorAction SilentlyContinue))
    $step = $null
    $log = Join-Path $WorkDir ("logs\{0:yyyy-MM-dd}.log" -f (Get-Date))
    if ($running -and (Test-Path $log)) {
        # Le journal est ouvert en écriture par le lanceur : lecture partagée.
        $fs = [IO.File]::Open($log, 'Open', 'Read', 'ReadWrite')
        try { $text = (New-Object IO.StreamReader($fs, [Text.Encoding]::UTF8)).ReadToEnd() } finally { $fs.Close() }
        $lines = @($text -split "`r?`n" | Where-Object { $_ -match 'Étape :' })
        if ($lines.Count) { $step = ($lines[-1] -replace '^\d\d:\d\d:\d\d\s*Étape :\s*', '').Trim() }
    }
    [ordered]@{
        enCours = $running
        debut   = $(if ($running) { $lock.debut } else { $null })
        mail    = $(if ($running) { [bool]$lock.mail } else { $null })
        etape   = $step
        dernier = Read-SmJson (Join-Path $WorkDir 'dernier-lancement.json')
    }
}

# État d'une rédaction avec tout le dossier (Rediger-Contact.ps1) ; une rédaction dont le processus
# a disparu sans conclure est signalée en échec. Le démarrage (pid encore vide) a 30 s pour se lancer.
function Get-RedactionState([string]$Id) {
    $state = Read-SmJson (Join-Path $WorkDir "redaction-$Id.json")
    if (-not $state) { return [ordered]@{ id = $Id; etat = 'aucune' } }
    if ($state.etat -eq 'en_cours') {
        $since = ConvertTo-SmDateTime $state.le
        $dead = $(if ($state.pid) { -not (Get-Process -Id $state.pid -ErrorAction SilentlyContinue) } else { $since -and ((Get-Date) - $since).TotalSeconds -gt 30 })
        if ($dead) { $state.etat = 'echec'; $state.message = 'Rédaction interrompue.' }
    }
    $state
}

# Session Claude Code interactive sur un point, dans le dossier AnsamDiag (son CLAUDE.md et ses skills),
# ouverte dans un onglet Windows Terminal. Chaque session porte un identifiant fixé d'avance
# (--session-id) et gardé sur le point, pour la reprendre plus tard (--resume).
$SmClaudeCommandes = [ordered]@{
    ticket    = 'Analyser le ticket'
    timesheet = 'Timesheet'
    cloture   = 'Clôture'
    escalade  = 'Escalade'
    libre     = 'Session libre'
}

function Start-ClaudeSession($Point, [string]$Commande, [string]$Reprendre) {
    $config = Get-SmConfig $WorkDir
    $dir = $config.dossierAnsamDiag
    if (-not (Test-Path $dir)) { throw "Dossier AnsamDiag introuvable : $dir" }
    $claude = Resolve-SmClaude $null
    $key = [string]$Point.ticket
    $label = $(if ($key) { $key } else { [string]$Point.id })
    $quote = { param($s) "'" + ([string]$s -replace "'", "''") + "'" }

    if ($Reprendre) {
        if ($Reprendre -notmatch '^[0-9a-f-]{36}$') { throw 'Session invalide.' }
        $line = "& $(& $quote $claude) --resume $(& $quote $Reprendre)"
        $title = "Claude · $label (reprise)"
        $session = $null
    } else {
        if (-not $SmClaudeCommandes.Contains($Commande)) { throw "Commande inconnue : $Commande" }
        if ($Commande -ne 'libre' -and -not $key) { throw 'Ce point n''a pas de ticket : seule la session libre est possible.' }
        $context = @(
            "Point de ma synthèse mails : « $($Point.titre) »" + $(if ($Point.client) { ", client $($Point.client)" } else { '' }) + '.'
            $(if ($Point.action) { "Ce qui est attendu : $($Point.action)" })
            $(if ($Point.contact -and $Point.contact.resume) { "Où on en est : $($Point.contact.resume)" })
            $(if ($Point.noteMoi) { "Ma note : $($Point.noteMoi)" })
        ) | Where-Object { $_ }
        $prompt = $(switch ($Commande) {
            'libre' { "Reprends ce sujet avec moi$(if ($key) { " (ticket $key)" }). " + ($context -join ' ') + ' Commence par relire le dossier du ticket s''il existe, puis propose la suite.' }
            default { "/$Commande $key" }
        })
        # Windows PowerShell 5.1 transmet mal les guillemets droits aux programmes externes.
        $prompt = $prompt -replace '"', "'"
        $session = [guid]::NewGuid().ToString()
        $name = "$label · $($SmClaudeCommandes[$Commande])" -replace '"', "'"
        $line = "& $(& $quote $claude) --session-id $session -n $(& $quote $name) $(& $quote $prompt)"
        $title = "Claude · $name"
    }

    # Petit script de lancement (UTF-8 avec BOM pour les accents sous Windows PowerShell 5.1).
    $launchDir = Join-Path $WorkDir 'claude'
    New-Item -ItemType Directory -Force -Path $launchDir | Out-Null
    $launcher = Join-Path $launchDir ("lancer-{0}-{1:yyyyMMddHHmmss}.ps1" -f ($label -replace '[^A-Za-z0-9-]', ''), (Get-Date))
    $body = "Set-Location -LiteralPath $(& $quote $dir)`r`n`$Host.UI.RawUI.WindowTitle = $(& $quote $title)`r`n$line`r`n"
    [IO.File]::WriteAllText($launcher, $body, (New-Object Text.UTF8Encoding($true)))
    Get-ChildItem $launchDir -Filter 'lancer-*.ps1' | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-7) } | Remove-Item -Force

    $terminal = Get-Command wt.exe -ErrorAction SilentlyContinue
    $psArgs = "-NoExit -NoProfile -ExecutionPolicy Bypass -File `"$launcher`""
    if ($terminal) { Start-Process $terminal.Source -ArgumentList "-w new --title `"$title`" -d `"$dir`" powershell.exe $psArgs" }
    else { Start-Process powershell.exe -ArgumentList $psArgs -WorkingDirectory $dir }
    $session
}

# -Mail : synthèse envoyée ; sinon actualisation (analyse et registres, sans mail).
function Start-Run([switch]$Mail) {
    $launcher = Join-Path $PSScriptRoot 'Lancer-SyntheseMails.ps1'
    $options = $(if ($Mail) { '-Force' } else { '-Force -PasDeMail' })
    Start-Process -FilePath "$env:WINDIR\System32\conhost.exe" -WindowStyle Hidden `
        -ArgumentList "--headless powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$launcher`" $options" | Out-Null
}

function Install-Task {
    & (Join-Path $PSScriptRoot 'Installer-TacheSyntheseMails.ps1') -At (Get-SmConfig $WorkDir).heure -TaskName $TaskName | Out-Null
    [void](Get-TaskState -Refresh)
}

if (Test-PageRunning) {
    # Le service (ou une autre fenêtre) sert déjà la page : il suffit de l'ouvrir.
    if (-not $Service -and -not $PasDeNavigateur) { Start-Process $url }
    Write-Host "Page déjà active : $url"
    return
}

$listener = New-Object Net.Sockets.TcpListener([Net.IPAddress]::Loopback, $port)
try { $listener.Start() } catch { throw "Port $port indisponible (utilisé par un autre programme ?) : changer « portPage » dans config.json. $($_.Exception.Message)" }
Write-Host "Page de suivi et de configuration : $url"
if ($Service) { Write-Host 'Mode permanent.' }
else { Write-Host "Arrêt depuis la page, Ctrl+C ici, ou automatiquement après $IdleMinutes min sans activité." }
if (-not $Service -and -not $PasDeNavigateur) { Start-Process $url }

[void](Get-TaskState)   # premier calcul avant toute requête
$pageHtml = $null
$pageHtmlAt = [datetime]::MinValue
$lastActivity = Get-Date
$stop = $false
try {
    while (-not $stop) {
        if (-not $listener.Pending()) {
            if (-not $Service -and ((Get-Date) - $lastActivity).TotalMinutes -ge $IdleMinutes) { Write-Host 'Arrêt après inactivité.'; break }
            Start-Sleep -Milliseconds 200
            continue
        }
        $client = $listener.AcceptTcpClient()
        $lastActivity = Get-Date
        try {
            $stream = $client.GetStream()
            $stream.ReadTimeout = 5000
            $request = Read-Request $stream
            if (-not $request) { continue }

            # Seule la page servie ici, avec le jeton de la session, peut appeler le serveur.
            $host_ = $request.Headers['host']
            $tokenOk = ($request.Query['t'] -eq $token) -or ($request.Headers['x-token'] -eq $token)
            if (@("127.0.0.1:$port", "localhost:$port") -notcontains $host_ -or -not $tokenOk) {
                Send-Response $stream 403 'text/plain' ("Accès refusé : lien incomplet ou ancien (il manque le jeton). Ouvrir la page avec le raccourci " +
                    "« Panoramix » du menu Démarrer, le lien « Ouvrir Panoramix » d'une synthèse, ou Configurer-SyntheseMails.cmd.")
                continue
            }

            switch -Regex ("$($request.Method) $($request.Path)") {
                '^GET /$' {
                    # Page relue seulement si le fichier a changé.
                    $written = (Get-Item $pagePath).LastWriteTime
                    if (-not $pageHtml -or $written -ne $pageHtmlAt) { $pageHtml = [IO.File]::ReadAllText($pagePath, [Text.Encoding]::UTF8); $pageHtmlAt = $written }
                    Send-Response $stream 200 'text/html' $pageHtml
                }
                '^GET /api/etat$' { Send-Json $stream 200 (Get-State) }
                '^POST /api/config$' {
                    $config = ConvertFrom-PageConfig $request.Body
                    $errors = Test-SmConfig $config
                    if ($errors.Count) { Send-Json $stream 400 ([ordered]@{ erreurs = $errors }); break }
                    Save-SmConfig $config $WorkDir
                    $message = 'Configuration enregistrée.'
                    if ((Get-TaskState).installee) { Install-Task; $message += ' Tâche planifiée mise à jour.' }
                    Send-Json $stream 200 ([ordered]@{ message = $message; etat = (Get-State) })
                }
                '^GET /api/suivi$' {
                    $suivi = Get-SmSuivi $WorkDir
                    Send-Json $stream 200 ([ordered]@{ majLe = $suivi.majLe; points = @($suivi.points) })
                }
                '^POST /api/suivi$' {
                    $posted = $request.Body | ConvertFrom-Json
                    $suivi = Get-SmSuivi $WorkDir
                    $point = Set-SmSuiviStatut $suivi ([string]$posted.id) ([string]$posted.statut) ([string]$posted.note)
                    Save-SmSuivi $suivi $WorkDir
                    Send-Json $stream 200 ([ordered]@{ message = "« $($point.titre) » : enregistré."; point = $point; points = @($suivi.points) })
                }
                '^POST /api/claude$' {
                    # Ouvre une session Claude Code sur un point (skill ou session libre), ou reprend une session.
                    $posted = $request.Body | ConvertFrom-Json
                    $suivi = Get-SmSuivi $WorkDir
                    $point = $suivi.points | Where-Object { $_.id -eq [string]$posted.id -and -not $_.estGroupe } | Select-Object -First 1
                    if (-not $point) { throw 'Point introuvable.' }
                    $session = Start-ClaudeSession $point ([string]$posted.commande) ([string]$posted.reprendre)
                    if ($session) {
                        $entry = [ordered]@{ id = $session; commande = [string]$posted.commande; nom = $SmClaudeCommandes[[string]$posted.commande]; le = (Get-Date).ToString('o') }
                        $point.claudeSessions = @(@($entry) + @($point.claudeSessions | Where-Object { $_ }) | Select-Object -First 5)
                        Save-SmSuivi $suivi $WorkDir
                        $message = "Session « $($entry.nom) » ouverte dans Windows Terminal."
                    } else {
                        $message = 'Session reprise dans Windows Terminal.'
                    }
                    Send-Json $stream 200 ([ordered]@{ message = $message; point = $point })
                }
                '^POST /api/redaction$' {
                    # Rédaction d'une prise de contact avec tout le dossier : en arrière-plan, 1 à 2 minutes.
                    $id = [string]($request.Body | ConvertFrom-Json).id
                    if ($id -notmatch '^[A-Za-z0-9-]+$') { throw "Identifiant invalide : $id" }
                    $state = Get-RedactionState $id
                    if ($state -and $state.etat -eq 'en_cours') { Send-Json $stream 400 ([ordered]@{ erreurs = @('Rédaction déjà en cours pour ce point.') }); break }
                    Write-SmJson (Join-Path $WorkDir "redaction-$id.json") ([ordered]@{ id = $id; etat = 'en_cours'; message = 'Démarrage'; le = (Get-Date).ToString('o'); pid = $null })
                    $script = Join-Path $PSScriptRoot 'Rediger-Contact.ps1'
                    Start-Process -FilePath "$env:WINDIR\System32\conhost.exe" -WindowStyle Hidden `
                        -ArgumentList "--headless powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$script`" -Id $id" | Out-Null
                    Send-Json $stream 200 ([ordered]@{ message = 'Rédaction lancée : relecture de tout le dossier, 1 à 2 minutes.' })
                }
                '^GET /api/redaction$' {
                    $id = [string]$request.Query['id']
                    if ($id -notmatch '^[A-Za-z0-9-]+$') { throw "Identifiant invalide : $id" }
                    Send-Json $stream 200 (Get-RedactionState $id)
                }
                '^POST /api/brouillon$' {
                    # Brouillon Outlook d'une prise de contact : enregistré dans Brouillons et ouvert pour relecture,
                    # jamais envoyé par le script.
                    $posted = $request.Body | ConvertFrom-Json
                    $to = ([string]$posted.a).Trim()
                    if ($to -and $to -notmatch '^[^@\s;,]+@[^@\s;,]+\.[^@\s;,]+([;,]\s*[^@\s;,]+@[^@\s;,]+\.[^@\s;,]+)*$') { throw "Adresse invalide : $to" }
                    $outlook = New-Object -ComObject Outlook.Application
                    $draft = $outlook.CreateItem(0)   # olMailItem
                    $draft.To = $to
                    $draft.Subject = [string]$posted.objet
                    $draft.Body = [string]$posted.message
                    $draft.Save()
                    try { $draft.Display($false) } catch { }
                    Send-Json $stream 200 ([ordered]@{ message = 'Brouillon créé dans Outlook (dossier Brouillons) : relis-le puis envoie-le toi-même.' })
                }
                '^POST /api/groupe$' {
                    # Groupement manuel : créer, ajouter, retirer, renommer, dissoudre.
                    $posted = $request.Body | ConvertFrom-Json
                    $suivi = Get-SmSuivi $WorkDir
                    $ids = @($posted.ids | Where-Object { $_ } | ForEach-Object { [string]$_ })
                    switch ([string]$posted.action) {
                        'creer' {
                            if (-not $ids.Count) { throw 'Aucune tâche sélectionnée.' }
                            $titre = ([string]$posted.titre).Trim()
                            if (-not $titre) { throw 'Donner un nom au groupe.' }
                            $group = Get-SmGroupe $suivi $titre
                            if (-not $group) { $group = New-SmGroupe $suivi $titre 'moi' }
                            Set-SmParent $suivi $ids $group.id
                            $message = "Groupe « $($group.titre) » : $($ids.Count) tâche(s) ajoutée(s)."
                        }
                        'ajouter' {
                            $group = $suivi.points | Where-Object { $_.id -eq [string]$posted.groupe -and $_.estGroupe } | Select-Object -First 1
                            if (-not $group) { throw 'Groupe introuvable.' }
                            Set-SmParent $suivi $ids $group.id
                            $message = "$($ids.Count) tâche(s) ajoutée(s) à « $($group.titre) »."
                        }
                        'retirer' {
                            Set-SmParent $suivi $ids $null
                            $message = "$($ids.Count) tâche(s) sortie(s) du groupe."
                        }
                        'renommer' {
                            $group = $suivi.points | Where-Object { $_.id -eq [string]$posted.groupe -and $_.estGroupe } | Select-Object -First 1
                            $titre = ([string]$posted.titre).Trim()
                            if (-not $group -or -not $titre) { throw 'Groupe ou nom manquant.' }
                            $group.titre = $titre
                            $group.creePar = 'moi'
                            $message = "Groupe renommé « $titre »."
                        }
                        'dissoudre' {
                            $children = @($suivi.points | Where-Object { $_.parent -eq [string]$posted.groupe } | ForEach-Object { $_.id })
                            Set-SmParent $suivi $children $null
                            $message = "Groupe dissous : $($children.Count) tâche(s) remise(s) dans la liste."
                        }
                        default { throw "Action inconnue : $($posted.action)." }
                    }
                    Save-SmSuivi $suivi $WorkDir
                    Send-Json $stream 200 ([ordered]@{ message = $message; points = @($suivi.points) })
                }
                '^GET /api/lancement$' { Send-Json $stream 200 (Get-RunState) }
                '^POST /api/lancement$' {
                    if ((Get-RunState).enCours) { Send-Json $stream 400 ([ordered]@{ erreurs = @('Une analyse est déjà en cours.') }); break }
                    $mail = $false
                    if ($request.Body) { $mail = [bool]($request.Body | ConvertFrom-Json).mail }
                    Start-Run -Mail:$mail
                    $message = $(if ($mail) { 'Synthèse lancée : environ 8 minutes, puis le mail arrive.' }
                                 else { 'Actualisation lancée : environ 7 minutes, sans mail.' })
                    Send-Json $stream 200 ([ordered]@{ message = $message })
                }
                '^GET /api/archives$' {
                    $archives = Get-SmArchives $WorkDir
                    Send-Json $stream 200 ([ordered]@{ majLe = $archives.majLe; elements = @($archives.elements) })
                }
                '^POST /api/archives$' {
                    $posted = $request.Body | ConvertFrom-Json
                    $archives = Get-SmArchives $WorkDir
                    $item = Set-SmArchiveStatut $archives ([string]$posted.id) ([string]$posted.statut) ([string]$posted.note)
                    Save-SmArchives $archives $WorkDir
                    Send-Json $stream 200 ([ordered]@{ message = 'Enregistré.'; element = $item })
                }
                '^POST /api/tache$' {
                    Install-Task
                    Send-Json $stream 200 ([ordered]@{ message = "Tâches installées : synthèse à $((Get-SmConfig $WorkDir).heure), page permanente à l'ouverture de session."; etat = (Get-State) })
                }
                '^GET /synthese$' {
                    $name = [string]$request.Query['d']
                    $file = Join-Path $WorkDir "syntheses\$name.md"
                    if ($name -notmatch '^\d{4}-\d{2}-\d{2}$' -or -not (Test-Path $file)) { Send-Response $stream 404 'text/plain' 'Synthèse introuvable.'; break }
                    Send-Response $stream 200 'text/html' (ConvertTo-SmHtml ([IO.File]::ReadAllText($file, [Text.Encoding]::UTF8)))
                }
                '^POST /api/quitter$' {
                    if ($Service) { Send-Json $stream 200 ([ordered]@{ message = 'Page permanente : elle reste active.' }); break }
                    Send-Json $stream 200 ([ordered]@{ message = 'Page arrêtée.' })
                    $stop = $true
                }
                default { Send-Response $stream 404 'text/plain' 'Introuvable.' }
            }
        } catch {
            try { Send-Json $stream 500 ([ordered]@{ erreurs = @($_.Exception.Message) }) } catch { }
            Write-Warning $_.Exception.Message
        } finally {
            $client.Close()
        }
    }
} finally {
    $listener.Stop()
    # Fenêtre ponctuelle arrêtée : la page permanente reprend le relais si elle est installée.
    if (-not $Service -and (Get-ScheduledTask -TaskName "$TaskName - page" -ErrorAction SilentlyContinue)) {
        Start-ScheduledTask -TaskName "$TaskName - page"
    }
}
