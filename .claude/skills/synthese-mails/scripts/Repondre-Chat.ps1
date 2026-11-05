<#
.SYNOPSIS
    Produit la réponse du chat de Panoramix au dernier message d'une conversation.

.DESCRIPTION
    La page dépose le message dans <WorkDir>\chat\<id>.json et lance ce script.
    Claude Code (claude -p) répond en lecture seule : registres de Panoramix,
    dossiers AnsamDiag et clients, Jira et Confluence par le connecteur, Outlook
    en direct par Serveur-OutilsChat.ps1. Il ne peut ni écrire dans un fichier,
    ni lancer de commande, ni aller sur le web, ni écrire dans Jira.

    La réponse arrive en continu dans le fichier de la conversation, que la page
    relit. D'un message à l'autre, la session Claude Code est reprise (--resume).
    Le modèle par défaut est Sonnet : une question n'a pas besoin d'Opus, et le
    forfait est limité.

    Les conversations de plus de -KeepDays jours sont effacées, comme les exports :
    elles contiennent des extraits de mails.

.EXAMPLE
    .\Repondre-Chat.ps1 -Id 20261104-120000-ab12
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string]$Id,
    [string]$WorkDir = (Join-Path $env:USERPROFILE 'SyntheseMails'),
    [string]$RepoDir,
    [string]$ClaudeExe,
    [string]$Modele = 'sonnet',
    [int]$KeepDays = 14,
    [int]$TimeoutMinutes = 10
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'SyntheseMails.Commun.ps1')
if ($Id -notmatch '^[0-9]{8}-[0-9]{6}-[0-9a-f]{4}$') { throw "Identifiant invalide : $Id" }
if (-not $RepoDir) { $RepoDir = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path }
$ClaudeExe = Resolve-SmClaude $ClaudeExe

$chatDir = Join-Path $WorkDir 'chat'
$file = Join-Path $chatDir "$Id.json"
$conv = Read-SmJson $file
if (-not $conv) { throw "Conversation introuvable : $Id" }
$utf8 = New-Object Text.UTF8Encoding($false)

$prefix = 'mcp__claude_ai_Atlassian_MCP__'
$lectureAtlassian = 'atlassianUserInfo', 'getAccessibleAtlassianResources', 'getJiraIssue', 'searchJiraIssuesUsingJql', 'listJiraIssueComments',
                    'discover', 'executeRead', 'searchConfluence', 'getConfluenceContent', 'search'
$ecritureAtlassian = 'executeWrite', 'executeDestructive', 'createConfluenceContent', 'updateConfluenceContent'
$allowed = (@('Read', 'Grep', 'Glob', 'mcp__panoramix__chercher_mails', 'mcp__panoramix__lire_mail', 'mcp__panoramix__lire_agenda') +
            ($lectureAtlassian | ForEach-Object { $prefix + $_ })) -join ','
$denied = (@('Bash', 'PowerShell', 'Write', 'Edit', 'NotebookEdit', 'WebFetch', 'WebSearch', 'Agent', 'Task', 'Read(**/page.token)','Read(**/jira.jeton)') +
           ($ecritureAtlassian | ForEach-Object { $prefix + $_ })) -join ','

# Libellés affichés pendant que Claude travaille.
function Get-Activite([string]$Outil) {
    switch -Regex ($Outil) {
        'chercher_mails'                      { return 'Recherche dans les mails' }
        'lire_mail'                           { return 'Lecture d''un mail' }
        'lire_agenda'                         { return 'Lecture de l''agenda' }
        'Confluence'                          { return 'Lecture de Confluence' }
        'Comment|Jira|executeRead|discover'   { return 'Lecture de Jira' }
        '^(Read|Grep|Glob)$'                  { return 'Lecture des dossiers' }
        default                               { return 'Recherche en cours' }
    }
}

function Save-Conversation {
    $conv.le = (Get-Date).ToString('o')
    $temporaire = "$file.tmp"
    [IO.File]::WriteAllText($temporaire, (ConvertTo-SmJson $conv), $utf8)
    # Remplacement d'un coup : la page ne doit jamais relire un fichier à moitié écrit. Il échoue si elle le lit à cet instant.
    for ($essai = 0; ; $essai++) {
        try { [IO.File]::Replace($temporaire, $file, [NullString]::Value); break }
        catch { if ($essai -ge 10) { throw }; Start-Sleep -Milliseconds 40 }
    }
}

function Add-Champ([string]$Nom, $Valeur) { $conv | Add-Member -NotePropertyName $Nom -NotePropertyValue $Valeur -Force }

function Set-Fin([string]$Etat, [string]$Erreur) {
    Add-Champ 'etat' $Etat
    Add-Champ 'erreur' $Erreur
    Add-Champ 'activite' $null
    Add-Champ 'partiel' $null
    Add-Champ 'pid' $null
    Save-Conversation
}

# Le processus de Claude ne doit pas survivre à ce script (arrêt demandé dans la page ou délai dépassé).
function Stop-Claude($Process) {
    if ($Process -and -not $Process.HasExited) { try { & taskkill.exe /T /F /PID $Process.Id 2>&1 | Out-Null } catch { } }
}

$process = $null
try {
    Add-Champ 'pid' $PID
    Add-Champ 'activite' 'Lecture de ta question'
    Add-Champ 'partiel' $null
    Save-Conversation

    Get-ChildItem $chatDir -File -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$KeepDays) } | Remove-Item -ErrorAction SilentlyContinue

    $config = Get-SmConfig $WorkDir
    $consigne = [IO.File]::ReadAllText((Join-Path $PSScriptRoot '..\references\consigne-chat.md'), $utf8).
        Replace('{{WorkDir}}', $WorkDir).Replace('{{AnsamDiag}}', [string]$config.dossierAnsamDiag).Replace('{{Clients}}', [string]$config.dossierClients)
    $consigneFile = Join-Path $chatDir 'consigne.md'
    [IO.File]::WriteAllText($consigneFile, $consigne, $utf8)

    $serveur = Join-Path $PSScriptRoot 'Serveur-OutilsChat.ps1'
    $mcpFile = Join-Path $chatDir 'mcp.json'
    Write-SmJson $mcpFile ([ordered]@{ mcpServers = [ordered]@{ panoramix = [ordered]@{
        type = 'stdio'; command = 'powershell.exe'; args = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $serveur) } } })

    # Le message passe par l'entrée standard : guillemets et sauts de ligne n'ont pas à être protégés.
    $dernier = @($conv.messages)[-1]
    $entete = "Nous sommes le $((Get-Date).ToString('dddd dd.MM.yyyy, HH:mm', [Globalization.CultureInfo]::GetCultureInfo('fr-CH')))."
    if ($conv.pointOuvert) { $entete += " Tâche ouverte dans la page : $($conv.pointOuvert | ConvertTo-Json -Compress)." }
    $message = "$entete`n`n$($dernier.texte)"

    $arguments = @('-p', '--output-format', 'stream-json', '--verbose', '--include-partial-messages', '--model', $Modele,
                   '--append-system-prompt-file', "`"$consigneFile`"", '--mcp-config', "`"$mcpFile`"",
                   '--allowedTools', "`"$allowed`"", '--disallowedTools', "`"$denied`"")
    foreach ($dossier in $WorkDir, $config.dossierAnsamDiag, $config.dossierClients) {
        if ($dossier -and (Test-Path $dossier)) { $arguments += '--add-dir'; $arguments += "`"$dossier`"" }
    }
    if ($conv.sessionId) { $arguments += '--resume'; $arguments += [string]$conv.sessionId }

    $info = New-Object Diagnostics.ProcessStartInfo
    $info.FileName = $ClaudeExe
    $info.Arguments = $arguments -join ' '
    $info.WorkingDirectory = $RepoDir
    $info.UseShellExecute = $false
    $info.CreateNoWindow = $true
    $info.RedirectStandardInput = $true
    $info.RedirectStandardOutput = $true
    $info.RedirectStandardError = $true
    $info.StandardOutputEncoding = [Text.Encoding]::UTF8
    $info.StandardErrorEncoding = [Text.Encoding]::UTF8
    $process = [Diagnostics.Process]::Start($info)
    $erreurs = $process.StandardError.ReadToEndAsync()
    $octets = $utf8.GetBytes($message)
    $process.StandardInput.BaseStream.Write($octets, 0, $octets.Length)
    $process.StandardInput.Close()

    $limite = (Get-Date).AddMinutes($TimeoutMinutes)
    $texte = New-Object Text.StringBuilder
    $resultat = $null
    $sauve = Get-Date
    while ($true) {
        $lecture = $process.StandardOutput.ReadLineAsync()
        while (-not $lecture.Wait(500)) {
            if ((Get-Date) -gt $limite) { Stop-Claude $process; throw "Pas de réponse après $TimeoutMinutes minutes." }
        }
        $ligne = $lecture.Result
        if ($null -eq $ligne) { break }
        if (-not $ligne.StartsWith('{')) { continue }
        # Les résultats d'outils (type user) peuvent peser lourd et ne servent pas ici.
        if ($ligne.StartsWith('{"type":"user"')) { continue }
        try { $event = $ligne | ConvertFrom-Json } catch { continue }

        if ($event.session_id -and -not $conv.sessionId) { Add-Champ 'sessionId' ([string]$event.session_id) }
        if ($event.type -eq 'stream_event') {
            $flux = $event.event
            if ($flux.type -eq 'message_start') { [void]$texte.Clear() }
            elseif ($flux.type -eq 'content_block_delta' -and $flux.delta.type -eq 'text_delta') { [void]$texte.Append([string]$flux.delta.text) }
            elseif ($flux.type -eq 'content_block_start' -and $flux.content_block.type -eq 'tool_use') {
                Add-Champ 'activite' (Get-Activite ([string]$flux.content_block.name))
                Add-Champ 'partiel' $null
                Save-Conversation
                continue
            }
            if ($texte.Length -and ((Get-Date) - $sauve).TotalMilliseconds -ge 350) {
                Add-Champ 'partiel' $texte.ToString()
                Save-Conversation
                $sauve = Get-Date
            }
        } elseif ($event.type -eq 'result') {
            $resultat = $event
        }
    }
    $process.WaitForExit()

    if ($resultat -and -not $resultat.is_error -and $resultat.result) {
        $conv.messages = @(@($conv.messages) + [pscustomobject]@{ role = 'claude'; texte = [string]$resultat.result; le = (Get-Date).ToString('o') })
        Set-Fin 'pret' $null
    } else {
        $detail = $(if ($resultat -and $resultat.result) { [string]$resultat.result } else { ([string]$erreurs.Result).Trim() })
        if (-not $detail) { $detail = "Claude Code s'est arrêté sans réponse (code $($process.ExitCode))." }
        if ($detail.Length -gt 400) { $detail = $detail.Substring(0, 400) }
        Set-Fin 'echec' $detail
    }
} catch {
    Stop-Claude $process
    Set-Fin 'echec' $_.Exception.Message
    throw
}
