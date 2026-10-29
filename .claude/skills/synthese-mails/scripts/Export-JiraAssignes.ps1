<#
.SYNOPSIS
    Écrit jira-assignes.json : mes tickets Jira ouverts, lus par Claude Code
    avec le connecteur Atlassian, en lecture seule.

.DESCRIPTION
    Claude Code (claude -p) n'a le droit qu'aux trois outils de lecture du
    connecteur Atlassian ; tout outil d'écriture est refusé. Le résultat est
    contrôlé (tableau JSON avec une clé par ticket) avant d'être écrit. En cas
    d'échec, le fichier précédent est laissé tel quel et le script renvoie
    $null : la synthèse continue sans Jira et le signale.

.EXAMPLE
    .\Export-JiraAssignes.ps1 -OutputFile "$env:USERPROFILE\SyntheseMails\jira-assignes.json"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string]$OutputFile,
    [string]$ClaudeExe,
    [string]$Jql = 'assignee = currentUser() AND statusCategory != Done ORDER BY updated DESC'
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'SyntheseMails.Commun.ps1')
$ClaudeExe = Resolve-SmClaude $ClaudeExe

$prefix = 'mcp__claude_ai_Atlassian_MCP__'
$allowed = (@('atlassianUserInfo', 'getAccessibleAtlassianResources', 'searchJiraIssuesUsingJql') | ForEach-Object { $prefix + $_ }) -join ','
$denied = (@('executeWrite', 'executeDestructive', 'createConfluenceContent', 'updateConfluenceContent') | ForEach-Object { $prefix + $_ }) -join ','

$prompt = "Avec le connecteur Atlassian, en lecture seule, cherche mes tickets Jira avec la JQL « $Jql » " +
          "(200 résultats au plus ; champs summary, status, project, priority, duedate, updated, reporter, created). " +
          "Réponds UNIQUEMENT par un tableau JSON, sans texte autour : " +
          "[{""cle"":"""",""titre"":"""",""statut"":"""",""projet"":"""",""priorite"":"""",""echeance"":null,""creeLe"":"""",""misAJour"":"""",""rapporteur"":""""}]."

[Console]::OutputEncoding = [Text.Encoding]::UTF8
try {
    $output = & $ClaudeExe -p $prompt --allowedTools $allowed --disallowedTools $denied
    if ($LASTEXITCODE -ne 0) { throw "Claude Code a échoué (code $LASTEXITCODE) : $($output -join ' ')" }
    $text = $output -join "`n"
    $start = $text.IndexOf('[')
    $end = $text.LastIndexOf(']')
    if ($start -lt 0 -or $end -le $start) { throw "Réponse sans tableau JSON : $($text.Substring(0, [math]::Min(300, $text.Length)))" }
    # Windows PowerShell 5.1 renvoie un tableau JSON comme un seul objet : on le déroule.
    $parsed = $text.Substring($start, $end - $start + 1) | ConvertFrom-Json
    $tickets = @($parsed | ForEach-Object { $_ })
    if ($tickets.Count -and -not $tickets[0].cle) { throw 'Tableau JSON sans clé de ticket.' }
} catch {
    Write-Warning "Tickets Jira non récupérés : $($_.Exception.Message)"
    return $null
}

$result = [ordered]@{ lusLe = (Get-Date).ToString('o'); jql = $Jql; tickets = $tickets }
Write-SmJson $OutputFile $result
Write-Host "Jira : $($tickets.Count) tickets assignés ouverts -> $OutputFile"
$OutputFile
