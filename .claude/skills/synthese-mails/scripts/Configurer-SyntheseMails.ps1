<#
.SYNOPSIS
    Ouvre la page locale de configuration de la synthèse mails.

.DESCRIPTION
    Démarre un petit serveur web sur 127.0.0.1 (port libre choisi au hasard,
    jamais accessible depuis le réseau), ouvre la page dans le navigateur et
    s'arrête quand on clique sur « Fermer » ou après 30 minutes sans activité.
    Chaque requête doit porter le jeton aléatoire de la session.

    La page règle <WorkDir>\config.json : synthèse active, heure, jours
    d'envoi, fenêtre en jours, envoi par mail, absences, dossiers. Elle
    affiche la prochaine synthèse prévue, l'état de la tâche planifiée et les
    dernières synthèses, et peut installer ou mettre à jour la tâche.

.EXAMPLE
    .\Configurer-SyntheseMails.ps1      # ou double-clic sur Configurer-SyntheseMails.cmd
#>
[CmdletBinding()]
param(
    [string]$WorkDir = (Join-Path $env:USERPROFILE 'SyntheseMails'),
    [string]$TaskName = 'Synthese mails ANSAM',
    [int]$IdleMinutes = 30,
    # Affiche l'adresse sans ouvrir le navigateur.
    [switch]$PasDeNavigateur
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'SyntheseMails.Commun.ps1')

$utf8 = New-Object Text.UTF8Encoding($false)
$pagePath = Join-Path $PSScriptRoot 'config.html'
$token = [guid]::NewGuid().ToString('N')

function Get-TaskState {
    $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    if (-not $task) { return [ordered]@{ installee = $false } }
    $info = Get-ScheduledTaskInfo -TaskName $TaskName
    $trigger = @($task.Triggers)[0]
    [ordered]@{
        installee          = $true
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

function Install-Task {
    & (Join-Path $PSScriptRoot 'Installer-TacheSyntheseMails.ps1') -At (Get-SmConfig $WorkDir).heure -TaskName $TaskName | Out-Null
}

$listener = New-Object Net.Sockets.TcpListener([Net.IPAddress]::Loopback, 0)
$listener.Start()
$port = $listener.LocalEndpoint.Port
$url = "http://127.0.0.1:$port/?t=$token"
Write-Host "Page de configuration : $url"
Write-Host "Fermer depuis la page, ou Ctrl+C ici. Arrêt automatique après $IdleMinutes min sans activité."
if (-not $PasDeNavigateur) { Start-Process $url }

$lastActivity = Get-Date
$stop = $false
try {
    while (-not $stop) {
        if (-not $listener.Pending()) {
            if (((Get-Date) - $lastActivity).TotalMinutes -ge $IdleMinutes) { Write-Host 'Arrêt après inactivité.'; break }
            Start-Sleep -Milliseconds 150
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
            if ($host_ -ne "127.0.0.1:$port" -or -not $tokenOk) {
                Send-Response $stream 403 'text/plain' 'Accès refusé : ouvrir la page depuis Configurer-SyntheseMails.'
                continue
            }

            switch -Regex ("$($request.Method) $($request.Path)") {
                '^GET /$' {
                    Send-Response $stream 200 'text/html' ([IO.File]::ReadAllText($pagePath, [Text.Encoding]::UTF8))
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
                '^POST /api/tache$' {
                    Install-Task
                    Send-Json $stream 200 ([ordered]@{ message = "Tâche planifiée installée ($((Get-SmConfig $WorkDir).heure))."; etat = (Get-State) })
                }
                '^GET /synthese$' {
                    $name = [string]$request.Query['d']
                    $file = Join-Path $WorkDir "syntheses\$name.md"
                    if ($name -notmatch '^\d{4}-\d{2}-\d{2}$' -or -not (Test-Path $file)) { Send-Response $stream 404 'text/plain' 'Synthèse introuvable.'; break }
                    Send-Response $stream 200 'text/html' (ConvertTo-SmHtml ([IO.File]::ReadAllText($file, [Text.Encoding]::UTF8)))
                }
                '^POST /api/quitter$' {
                    Send-Json $stream 200 ([ordered]@{ message = 'Page fermée.' })
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
}
