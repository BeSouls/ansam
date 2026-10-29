<#
.SYNOPSIS
    Réécrit la prise de contact d'un point (brouillon de mail ou points
    d'appel) en relisant tout son dossier.

.DESCRIPTION
    Claude Code (claude -p, lecture seule) relit, pour ce seul point :
    - le dossier AnsamDiag du ticket (<dossierAnsamDiag>\Tickets\<CLE>) :
      toutes les notes de ticket.json, TICKET.md, POUR-AGENT.md, les messages
      déjà rédigés (message-*.md), cloture.md, et les résultats des tests
      (<horodatage>_<Outil>\resume.md) ;
    - le ticket Jira avec ses commentaires (outils de lecture du connecteur
      Atlassian seulement, écriture refusée) ;
    - tous les mails du dernier export qui concernent le point.
    Il renvoie la prise de contact au format du registre, plus « resume » (où
    on en est : ce qui a été fait, testé, la suite prévue). Le résultat
    remplace le champ contact du point dans suivi.json.

    L'avancement est écrit dans <WorkDir>\redaction-<id>.json (en_cours, ok,
    echec) pour la page de suivi. Rien n'est envoyé.

.EXAMPLE
    .\Rediger-Contact.ps1 -Id ASD-26470
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string]$Id,
    [string]$WorkDir = (Join-Path $env:USERPROFILE 'SyntheseMails'),
    [string]$RepoDir,
    [string]$ClaudeExe
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'SyntheseMails.Commun.ps1')
if (-not $RepoDir) { $RepoDir = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path }
$ClaudeExe = Resolve-SmClaude $ClaudeExe
if ($Id -notmatch '^[A-Za-z0-9-]+$') { throw "Identifiant invalide : $Id" }

$statusFile = Join-Path $WorkDir "redaction-$Id.json"
function Set-Status([string]$State, [string]$Message) {
    Write-SmJson $statusFile ([ordered]@{ id = $Id; etat = $State; message = $Message; le = (Get-Date).ToString('o'); pid = $PID })
}
Set-Status 'en_cours' 'Lecture du dossier'

try {
    $point = (Get-SmSuivi $WorkDir).points | Where-Object { $_.id -eq $Id } | Select-Object -First 1
    if (-not $point) { throw "Point « $Id » introuvable." }
    $config = Get-SmConfig $WorkDir

    $dirs = @($WorkDir)
    $sources = @()
    $ticketDir = $null
    if ($point.ticket) {
        $candidate = Join-Path $config.dossierAnsamDiag "Tickets\$($point.ticket)"
        if (Test-Path $candidate) { $ticketDir = $candidate; $dirs += $candidate; $sources += "dossier AnsamDiag : $candidate" }
    }
    $export = Get-ChildItem (Join-Path $WorkDir 'exports') -Filter 'mails-*.json' -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notlike '*-fils*' } | Sort-Object Name | Select-Object -Last 1
    if ($export) { $sources += "mails : $($export.FullName)" }

    $current = $(if ($point.contact) { $point.contact | ConvertTo-Json -Depth 4 -Compress } else { 'aucune' })
    $pointInfo = [ordered]@{ id = $point.id; titre = $point.titre; client = $point.client; ticket = $point.ticket; statut = $point.statut;
                             action = $point.action; source = $point.source; noteMoi = $point.noteMoi } | ConvertTo-Json -Compress

    $prompt = "Tu prépares une prise de contact pour un point de ma synthèse mails ANSAM (skill synthese-mails : applique " +
              "« Prise de contact » de references/regles-classement.md). Point : $pointInfo. Prise de contact actuelle : $current. " +
              "Avant d'écrire, relis TOUT le contexte disponible : " +
              $(if ($ticketDir) { "le dossier $ticketDir (toutes les notes de ticket.json, TICKET.md, POUR-AGENT.md, brief-*.md, message-*.md, cloture.md, et chaque <horodatage>_<Outil>\resume.md : ce qui a été testé et trouvé) ; " } else { '' }) +
              $(if ($point.ticket) { "le ticket Jira $($point.ticket) avec tous ses commentaires (connecteur Atlassian, lecture seule) ; " } else { '' }) +
              $(if ($export) { "dans $($export.FullName), tous les mails du fil (même objet, même ticket, mêmes interlocuteurs), corps complets ; " } else { '' }) +
              "Ne repose pas une question à laquelle on a déjà répondu, ne redemande pas un test déjà fait, tiens compte du plan de résolution. " +
              "Garde le mode (mail ou appel) sauf si le dossier montre clairement que l'autre convient mieux. " +
              "Réponds UNIQUEMENT par un objet JSON : {""mode"":""mail|appel"",""nom"":"""",""role"":"""",""email"":null,""telephone"":null," +
              """objet"":null,""message"":null,""points"":[],""resume"":""où on en est : ce qui a été fait et testé, ce qui est établi, ce qui manque, la suite prévue (4 à 8 lignes)""}."

    $prefix = 'mcp__claude_ai_Atlassian_MCP__'
    $allowed = (@('Read', 'Grep', 'Glob', 'Skill') + (@('atlassianUserInfo', 'getAccessibleAtlassianResources', 'getJiraIssue', 'searchJiraIssuesUsingJql') | ForEach-Object { $prefix + $_ })) -join ','
    $denied = (@('executeWrite', 'executeDestructive', 'createConfluenceContent', 'updateConfluenceContent') | ForEach-Object { $prefix + $_ }) -join ','
    $addDirs = @(); foreach ($d in $dirs) { $addDirs += '--add-dir'; $addDirs += $d }

    Set-Status 'en_cours' 'Rédaction par Claude Code'
    [Console]::OutputEncoding = [Text.Encoding]::UTF8
    Push-Location $RepoDir
    try {
        $output = & $ClaudeExe -p $prompt @addDirs --allowedTools $allowed --disallowedTools $denied
        if ($LASTEXITCODE -ne 0) { throw "Claude Code a échoué (code $LASTEXITCODE) : $($output -join ' ')" }
    } finally { Pop-Location }

    $text = $output -join "`n"
    $start = $text.IndexOf('{'); $end = $text.LastIndexOf('}')
    if ($start -lt 0 -or $end -le $start) { throw 'Réponse sans objet JSON.' }
    $contact = $text.Substring($start, $end - $start + 1) | ConvertFrom-Json
    if (@('mail', 'appel') -notcontains $contact.mode) { throw "Mode de contact inattendu : $($contact.mode)" }
    $contact | Add-Member -NotePropertyName redigeLe -NotePropertyValue (Get-Date).ToString('o') -Force
    $contact | Add-Member -NotePropertyName sources -NotePropertyValue $sources -Force

    # Relecture juste avant l'écriture : une synthèse a pu modifier le registre entre-temps.
    $suivi = Get-SmSuivi $WorkDir
    $target = $suivi.points | Where-Object { $_.id -eq $Id } | Select-Object -First 1
    if (-not $target) { throw "Point « $Id » disparu du registre." }
    $target.contact = $contact
    Save-SmSuivi $suivi $WorkDir
    Set-Status 'ok' $(if ($contact.mode -eq 'appel') { 'Points d''appel rédigés avec tout le dossier.' } else { 'Brouillon rédigé avec tout le dossier.' })
} catch {
    Set-Status 'echec' $_.Exception.Message
    throw
}
