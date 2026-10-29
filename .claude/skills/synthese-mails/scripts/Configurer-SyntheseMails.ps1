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
        etape   = $step
        dernier = Read-SmJson (Join-Path $WorkDir 'dernier-lancement.json')
    }
}

function Start-Run {
    $launcher = Join-Path $PSScriptRoot 'Lancer-SyntheseMails.ps1'
    Start-Process -FilePath "$env:WINDIR\System32\conhost.exe" -WindowStyle Hidden `
        -ArgumentList "--headless powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$launcher`" -Force" | Out-Null
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
                    "« Synthèse mails – suivi » du menu Démarrer, le lien « Ouvrir mon suivi » d'une synthèse, ou Configurer-SyntheseMails.cmd.")
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
                    Send-Json $stream 200 ([ordered]@{ message = "« $($point.titre) » : enregistré."; point = $point })
                }
                '^GET /api/lancement$' { Send-Json $stream 200 (Get-RunState) }
                '^POST /api/lancement$' {
                    if ((Get-RunState).enCours) { Send-Json $stream 400 ([ordered]@{ erreurs = @('Une synthèse est déjà en cours.') }); break }
                    Start-Run
                    Send-Json $stream 200 ([ordered]@{ message = 'Synthèse lancée : environ 8 minutes, puis le mail arrive.' })
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
