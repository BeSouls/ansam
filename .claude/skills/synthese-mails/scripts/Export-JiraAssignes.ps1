<#
.SYNOPSIS
    Écrit jira-assignes.json : mes tickets Jira ouverts, et le statut réel des
    tickets que je suis, lus par Claude Code avec le connecteur Atlassian, en
    lecture seule.

.DESCRIPTION
    Deux recherches :
    - « tickets » : mes tickets ouverts (-Jql) ;
    - « suivis » : les tickets passés en -Cles (ceux des points ouverts du
      registre), quel que soit leur statut ou leur assigné. C'est ce qui dit
      si un ticket est clos, résolu, rouvert ou réassigné : un commentaire du
      client sur un ticket resté clos n'est pas une tâche.

    Claude Code (claude -p) n'a le droit qu'aux trois outils de lecture du
    connecteur Atlassian ; tout outil d'écriture est refusé. Le résultat est
    contrôlé avant d'être écrit. En cas d'échec, le fichier précédent est
    laissé tel quel et le script renvoie $null : la synthèse continue sans
    Jira et le signale.

.EXAMPLE
    .\Export-JiraAssignes.ps1 -OutputFile "$env:USERPROFILE\SyntheseMails\jira-assignes.json" -Cles ASD-27857, ASD-27228
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string]$OutputFile,
    [string]$ClaudeExe,
    [string]$Jql = 'assignee = currentUser() AND statusCategory != Done ORDER BY updated DESC',
    # Tickets dont il faut le statut réel (100 au plus).
    [string[]]$Cles = @()
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'SyntheseMails.Commun.ps1')
$ClaudeExe = Resolve-SmClaude $ClaudeExe

$prefix = 'mcp__claude_ai_Atlassian_MCP__'
$allowed = (@('atlassianUserInfo', 'getAccessibleAtlassianResources', 'searchJiraIssuesUsingJql') | ForEach-Object { $prefix + $_ }) -join ','
$denied = (@('executeWrite', 'executeDestructive', 'createConfluenceContent', 'updateConfluenceContent') | ForEach-Object { $prefix + $_ }) -join ','

$keys = @($Cles | Where-Object { $_ -match '^[A-Z][A-Z0-9]+-\d+$' } | Sort-Object -Unique | Select-Object -First 100)
$second = $(if ($keys.Count) {
    "Puis une seconde recherche, JQL « key in ($($keys -join ', ')) » (champs summary, status, assignee, resolution, updated) ; " +
    "ignore les clés introuvables. "
} else { '' })

$prompt = "Avec le connecteur Atlassian, en lecture seule, cherche mes tickets Jira avec la JQL « $Jql » " +
          "(200 résultats au plus ; champs summary, status, project, priority, duedate, updated, reporter (nom et adresse), created, " +
          "et les organisations du Service Desk du ticket, champ « Organizations »). " +
          $second +
          "Parcours toutes les pages de résultats de la première recherche, et indique dans « total » le nombre total de résultats annoncé par Jira. " +
          "Réponds UNIQUEMENT par un objet JSON, sans texte autour : " +
          "{""total"":0,""tickets"":[{""cle"":"""",""titre"":"""",""statut"":"""",""projet"":"""",""priorite"":"""",""echeance"":null,""creeLe"":"""",""misAJour"":"""",""rapporteur"":"""",""rapporteurEmail"":null,""organisations"":[]}]," +
          """suivis"":[{""cle"":"""",""titre"":"""",""statut"":"""",""categorie"":""À faire | En cours | Terminé"",""resolution"":null,""assigne"":"""",""misAJour"":""""}]}."

[Console]::OutputEncoding = [Text.Encoding]::UTF8
try {
    $output = & $ClaudeExe -p $prompt --allowedTools $allowed --disallowedTools $denied
    if ($LASTEXITCODE -ne 0) { throw "Claude Code a échoué (code $LASTEXITCODE) : $($output -join ' ')" }
    $text = $output -join "`n"
    $start = $text.IndexOf('{')
    $end = $text.LastIndexOf('}')
    if ($start -lt 0 -or $end -le $start) { throw "Réponse sans objet JSON : $($text.Substring(0, [math]::Min(300, $text.Length)))" }
    $parsed = $text.Substring($start, $end - $start + 1) | ConvertFrom-Json
    # Windows PowerShell 5.1 renvoie un tableau JSON comme un seul objet : on le déroule.
    $tickets = @($parsed.tickets | ForEach-Object { $_ } | Where-Object { $_ })
    $tracked = @($parsed.suivis | ForEach-Object { $_ } | Where-Object { $_ })
    if (($tickets.Count -and -not $tickets[0].cle) -or ($tracked.Count -and -not $tracked[0].cle)) { throw 'Tickets sans clé.' }
} catch {
    Write-Warning "Tickets Jira non récupérés : $($_.Exception.Message)"
    return $null
}

# Contrôle d'exhaustivité : la liste doit contenir autant de tickets que Jira en annonce.
$total = 0
if ($parsed.total) { $total = [int]$parsed.total }
$complete = (-not $total) -or ($tickets.Count -ge $total)
if (-not $complete) { Write-Warning "Liste Jira incomplète : $($tickets.Count) tickets reçus sur $total annoncés." }

$result = [ordered]@{ lusLe = (Get-Date).ToString('o'); jql = $Jql; total = $total; complet = $complete; tickets = $tickets; suivis = $tracked }
Write-SmJson $OutputFile $result
Write-Host "Jira : $($tickets.Count) tickets assignés ouverts (total annoncé : $total), $($tracked.Count) tickets suivis -> $OutputFile"
$OutputFile
