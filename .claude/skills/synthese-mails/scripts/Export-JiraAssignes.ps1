<#
.SYNOPSIS
    Écrit jira-assignes.json : mes tickets Jira ouverts, et le statut réel des
    tickets que je suis, lus par Claude Code avec le connecteur Atlassian, en
    lecture seule.

.DESCRIPTION
    Deux recherches :
    - « tickets » : mes tickets ouverts (-Jql) ;
    - « suivis » : parmi les tickets passés en -Cles (ceux des points ouverts du
      registre), ceux qui ne sont plus dans mes tickets ouverts. C'est ce qui
      dit si un ticket est clos, résolu ou réassigné : un commentaire du client
      sur un ticket resté clos n'est pas une tâche.

    Avec un jeton enregistré (Definir-JetonJira.ps1), les recherches vont
    directement à Jira, en quelques secondes. Sans jeton, ou si Jira le refuse :
    Claude Code (claude -p) n'a le droit qu'aux trois outils de lecture du
    connecteur Atlassian ; tout outil d'écriture est refusé. Recopier ces listes
    ne demande pas d'analyse : le modèle léger s'en charge. Le résultat est
    contrôlé avant d'être écrit : s'il manque un champ, le modèle standard
    refait la lecture. En cas d'échec, le fichier précédent est laissé tel quel
    et le script renvoie $null : la synthèse continue sans Jira et le signale.

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
    "Puis une seconde recherche, JQL « key in ($($keys -join ', ')) AND NOT (assignee = currentUser() AND statusCategory != Done) » (fields : summary, status, assignee, resolution, updated) ; " +
    "ignore les clés introuvables ; elle peut ne rien renvoyer. Chaque entrée de « suivis » contient toujours ses 7 clés : « assigne » = le nom affiché de assignee (null seulement s'il n'y en a pas), " +
    "« resolution » = le nom de resolution ou null, « categorie » = la catégorie du statut, « misAJour » = updated. "
} else { '' })

$prompt = "Avec le connecteur Atlassian, en lecture seule, cherche mes tickets Jira avec la JQL « $Jql » " +
          "(200 résultats au plus ; champs summary, status, project, priority, duedate, updated, reporter (nom et adresse), created, " +
          "et les organisations du Service Desk du ticket, champ « Organizations »). " +
          $second +
          "Parcours toutes les pages de résultats de la première recherche, et indique dans « total » le nombre total de résultats annoncé par Jira. " +
          "Indique dans « moi » mon nom affiché dans Jira (atlassianUserInfo). " +
          "Réponds UNIQUEMENT par un objet JSON, sans texte autour : " +
          "{""moi"":"""",""total"":0,""tickets"":[{""cle"":"""",""titre"":"""",""statut"":"""",""projet"":"""",""priorite"":"""",""echeance"":null,""creeLe"":"""",""misAJour"":"""",""rapporteur"":"""",""rapporteurEmail"":null,""organisations"":[]}]," +
          """suivis"":[{""cle"":"""",""titre"":"""",""statut"":"""",""categorie"":""À faire | En cours | Terminé"",""resolution"":null,""assigne"":"""",""misAJour"":""""}]}."

# Lecture directe de Jira, sans Claude : mêmes champs que par le connecteur.
function Read-JiraDirect($Acces) {
    $orgField = @(Invoke-SmJira $Acces '/rest/api/3/field' | ForEach-Object { $_ } | Where-Object { $_.name -eq 'Organizations' -and $_.custom } | Select-Object -First 1).id
    $fields = @('summary', 'status', 'project', 'priority', 'duedate', 'updated', 'reporter', 'created')
    if ($orgField) { $fields += $orgField }
    $tickets = @(Find-SmJira $Acces $Jql $fields | ForEach-Object {
        $f = $_.fields
        $orgs = @($(if ($orgField) { $f.$orgField }) | ForEach-Object { if ($_) { $_.name } })
        [ordered]@{ cle = $_.key; titre = ([string]$f.summary).Trim(); statut = $f.status.name; projet = $f.project.name; priorite = $f.priority.name; echeance = $f.duedate
                    creeLe = $f.created; misAJour = $f.updated; rapporteur = $f.reporter.displayName; rapporteurEmail = $f.reporter.emailAddress
                    organisations = $orgs }
    })

    # Une clé que Jira ne connaît plus fait échouer toute la recherche : on la retire et on recommence.
    $categories = @{ new = 'À faire'; indeterminate = 'En cours'; done = 'Terminé' }
    $wanted = @($keys)
    $issues = @()
    while ($wanted.Count) {
        try {
            $issues = @(Find-SmJira $Acces "key in ($($wanted -join ', ')) AND NOT (assignee = currentUser() AND statusCategory != Done)" @('summary', 'status', 'assignee', 'resolution', 'updated'))
            break
        } catch {
            $message = $_.Exception.Message
            $unknown = @($wanted | Where-Object { $message -match "\b$([regex]::Escape($_))\b" })
            if (-not $unknown.Count) { throw }
            $wanted = @($wanted | Where-Object { $unknown -notcontains $_ })
        }
    }
    $tracked = @($issues | ForEach-Object {
        $f = $_.fields
        [ordered]@{ cle = $_.key; titre = ([string]$f.summary).Trim(); statut = $f.status.name; categorie = $categories[[string]$f.status.statusCategory.key]
                    resolution = $f.resolution.name; assigne = $f.assignee.displayName; misAJour = $f.updated }
    })
    $me = Invoke-SmJira $Acces '/rest/api/3/myself'
    [pscustomobject]@{ parsed = [pscustomobject]@{ moi = $me.displayName; total = $tickets.Count }; tickets = $tickets; tracked = $tracked }
}

# Une lecture complète, ou une exception qui dit ce qui manque.
function Read-JiraReponse([string[]]$Modele) {
    $output = Invoke-SmClaude $ClaudeExe $prompt ($Modele + '--allowedTools', $allowed, '--disallowedTools', $denied)
    if ($LASTEXITCODE -ne 0) { throw "Claude Code a échoué (code $LASTEXITCODE) : $($output -join ' ')" }
    $text = $output -join "`n"
    $start = $text.IndexOf('{')
    $end = $text.LastIndexOf('}')
    if ($start -lt 0 -or $end -le $start) { throw "Réponse sans objet JSON : $($text.Substring(0, [math]::Min(300, $text.Length)))" }
    $parsed = $text.Substring($start, $end - $start + 1) | ConvertFrom-Json
    # Windows PowerShell 5.1 renvoie un tableau JSON comme un seul objet : on le déroule.
    $tickets = @($parsed.tickets | ForEach-Object { $_ } | Where-Object { $_ })
    $tracked = @($parsed.suivis | ForEach-Object { $_ } | Where-Object { $_ })
    if (@($tickets + $tracked | Where-Object { -not $_.cle -or -not $_.statut -or -not $_.misAJour }).Count) { throw 'Ticket sans clé, statut ou date.' }
    # Un ticket suivi est dans mes tickets ouverts ou dans la seconde recherche. Le modèle standard, lui,
    # peut laisser de côté une clé que Jira ne connaît plus.
    $listed = @($tickets + $tracked | ForEach-Object { [string]$_.cle })
    $absent = @($keys | Where-Object { $listed -notcontains $_ })
    if ($Modele.Count -and $absent.Count) { throw "Tickets suivis absents : $($absent -join ', ')." }
    $names = @('cle', 'titre', 'statut', 'categorie', 'resolution', 'assigne', 'misAJour')
    foreach ($entry in $tracked) {
        $missing = @($names | Where-Object { $entry.PSObject.Properties.Name -notcontains $_ })
        if ($missing.Count) { throw "Ticket suivi $($entry.cle) sans $($missing -join ', ')." }
    }
    if (-not $parsed.moi) { throw 'Nom Jira absent.' }
    [pscustomobject]@{ parsed = $parsed; tickets = $tickets; tracked = $tracked }
}

[Console]::OutputEncoding = [Text.Encoding]::UTF8
$reply = $null
$acces = Get-SmJiraAcces
if ($acces) {
    try { $reply = Read-JiraDirect $acces }
    catch { Write-Warning "Lecture directe de Jira impossible, passage par Claude : $($_.Exception.Message)" }
}
foreach ($modele in $(if ($reply) { @() } else { @(, @('--model', 'haiku')) + @(, @()) })) {
    try { $reply = Read-JiraReponse $modele; break }
    catch { Write-Warning "Tickets Jira, lecture $(if ($modele.Count) { 'légère' } else { 'standard' }) refusée : $($_.Exception.Message)" }
}
if (-not $reply) {
    Write-Warning 'Tickets Jira non récupérés.'
    return $null
}
$parsed = $reply.parsed
$tickets = $reply.tickets
$tracked = $reply.tracked

# Contrôle d'exhaustivité : la liste doit contenir autant de tickets que Jira en annonce.
$total = 0
if ($parsed.total) { $total = [int]$parsed.total }
$complete = (-not $total) -or ($tickets.Count -ge $total)
if (-not $complete) { Write-Warning "Liste Jira incomplète : $($tickets.Count) tickets reçus sur $total annoncés." }

$result = [ordered]@{ lusLe = (Get-Date).ToString('o'); moi = $parsed.moi; jql = $Jql; total = $total; complet = $complete; tickets = $tickets; suivis = $tracked }
Write-SmJson $OutputFile $result
Write-Host "Jira : $($tickets.Count) tickets assignés ouverts (total annoncé : $total), $($tracked.Count) tickets suivis -> $OutputFile"
$OutputFile
