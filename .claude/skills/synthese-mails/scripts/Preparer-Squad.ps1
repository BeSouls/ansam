<#
.SYNOPSIS
    Prépare mes lignes du formulaire de la réunion squad de la semaine.

.DESCRIPTION
    1. Date de la réunion : celle de la semaine (réglage « squadJour »), ou -Date.
       Sans -Force, rien n'est préparé si la réunion est désactivée ou si ce
       jour-là est une absence ou un jour férié.
    2. Export-OutlookAgenda.ps1 : mon agenda, de la semaine écoulée aux quatre
       semaines suivantes (lecture seule).
    3. Claude Code (claude -p, lecture seule) applique references/squad.md à
       l'agenda, au dernier export des mails, aux registres (suivi, éléments à
       garder), aux tickets Jira, au contexte AnsamDiag et à la fiche de la
       réunion précédente.
    4. Les lignes préparées remplacent les précédentes lignes préparées de la
       fiche (<WorkDir>\squad\AAAA-MM-JJ.json) ; ce que j'ai écrit ou modifié
       dans la page reste.

    L'avancement est écrit dans <WorkDir>\squad-preparation.json (en_cours,
    ok, echec, ignore) pour la page. Rien n'est envoyé.

.EXAMPLE
    .\Preparer-Squad.ps1                       # réunion de la semaine
    .\Preparer-Squad.ps1 -Date 2026-11-02 -Force
#>
[CmdletBinding()]
param(
    [string]$Date,
    [string]$WorkDir = (Join-Path $env:USERPROFILE 'SyntheseMails'),
    [string]$RepoDir,
    [string]$ClaudeExe,
    # Prépare même si la réunion est désactivée ou tombe un jour d'absence ou férié.
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'SyntheseMails.Commun.ps1')
if (-not $RepoDir) { $RepoDir = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path }

$config = Get-SmConfig $WorkDir
if (-not $Date) { $Date = Get-SmSquadDate $config }
$meeting = ConvertTo-SmDate $Date
if (-not $meeting) { throw "Date de réunion invalide : $Date" }

$statusFile = Join-Path $WorkDir 'squad-preparation.json'
function Set-Status([string]$State, [string]$Message) {
    Write-SmJson $statusFile ([ordered]@{ date = $Date; etat = $State; message = $Message; le = (Get-Date).ToString('o'); pid = $PID })
}

if (-not $Force) {
    $reason = $(if (-not $config.squadActif) { 'réunion squad désactivée' } else { Get-SmSquadEmpechement $config $meeting })
    if ($reason) {
        Set-Status 'ignore' "Pas de préparation pour le $($meeting.ToString('dd.MM')) : $reason."
        Write-Host "Réunion squad du $($meeting.ToString('dd.MM')) non préparée : $reason."
        return
    }
}

Set-Status 'en_cours' 'Lecture de l''agenda'
try {
    $ClaudeExe = Resolve-SmClaude $ClaudeExe
    $exportDir = Join-Path $WorkDir 'exports'
    $agendaFile = Join-Path $exportDir "agenda-$Date.json"
    & (Join-Path $PSScriptRoot 'Export-OutlookAgenda.ps1') -Debut $meeting.AddDays(-7) -Fin $meeting.AddDays(27) -OutputFile $agendaFile | Out-Null

    $export = Get-ChildItem $exportDir -Filter 'mails-*.json' -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notlike '*-fils*' } | Sort-Object Name | Select-Object -Last 1
    $threads = $(if ($export) { [IO.Path]::ChangeExtension($export.FullName, $null).TrimEnd('.') + '-fils.txt' })
    $previous = @(Get-SmSquadDates $WorkDir) | Where-Object { $_ -lt $Date } | Select-Object -First 1
    $files = [ordered]@{
        'agenda'                   = $agendaFile
        'mails, vue par fil'       = $threads
        'mails, détail'            = $(if ($export) { $export.FullName })
        'registre de suivi'        = Join-Path $WorkDir 'suivi.json'
        'éléments à garder'        = Join-Path $WorkDir 'archives.json'
        'tickets Jira assignés'    = Join-Path $WorkDir 'jira-assignes.json'
        'contexte AnsamDiag'       = Join-Path $WorkDir 'contexte-ansam.json'
        'fiche précédente'         = $(if ($previous) { Get-SmSquadPath $previous $WorkDir })
    }
    $sources = ($files.Keys | Where-Object { $files[$_] -and (Test-Path $files[$_]) } | ForEach-Object { "$_ : $($files[$_])" }) -join ' ; '
    $absences = (@($config.absences) | Where-Object { $_.fin -ge $Date } | ForEach-Object { "$($_.libelle) du $($_.debut) au $($_.fin)" }) -join ', '

    $prompt = "Prépare mes lignes du formulaire de la réunion squad du $($meeting.ToString('dd.MM.yyyy')) : applique " +
              ".claude/skills/synthese-mails/references/squad.md (lis-le d'abord). Mes initiales : $($config.squadMoi). " +
              "Semaine écoulée : du $($meeting.AddDays(-7).ToString('dd.MM')) au $($meeting.AddDays(-3).ToString('dd.MM')) ; " +
              "semaine de la réunion : du $($meeting.ToString('dd.MM')) au $($meeting.AddDays(4).ToString('dd.MM')). " +
              "Mes absences prévues (réglages) : $(if ($absences) { $absences } else { 'aucune' }). " +
              "Sources : $sources. " +
              "Réponds UNIQUEMENT par l'objet JSON décrit dans squad.md."

    Set-Status 'en_cours' 'Préparation par Claude Code'
    Push-Location $RepoDir
    try {
        $output = Invoke-SmClaude $ClaudeExe $prompt @('--add-dir', $WorkDir, '--allowedTools', 'Read,Grep,Glob', '--disallowedTools', 'Bash,PowerShell,Write,Edit')
        if ($LASTEXITCODE -ne 0) { throw "Claude Code a échoué (code $LASTEXITCODE) : $($output -join ' ')" }
    } finally { Pop-Location }

    $text = $output -join "`n"
    $start = $text.IndexOf('{'); $end = $text.LastIndexOf('}')
    if ($start -lt 0 -or $end -le $start) { throw 'Réponse sans objet JSON.' }
    $answer = $text.Substring($start, $end - $start + 1) | ConvertFrom-Json
    $ids = @($SmSquadSujets | ForEach-Object { $_.id })
    $lines = @(@($answer.lignes) | Where-Object { $_ -and $ids -contains [string]$_.sujet })

    # Relecture juste avant l'écriture : la fiche a pu être modifiée dans la page entre-temps.
    $doc = Get-SmSquad $config $Date $WorkDir
    if (-not $doc) { $doc = New-SmSquad $config $Date }
    Merge-SmSquad $doc $lines
    $doc.resume = ([string]$answer.resume).Trim()
    $doc.prepareLe = (Get-Date).ToString('o')
    Save-SmSquad $doc $WorkDir
    Set-Status 'ok' "Réunion du $($meeting.ToString('dd.MM')) préparée : $($lines.Count) lignes."
    Write-Host "Réunion squad du $($meeting.ToString('dd.MM')) préparée : $($lines.Count) lignes."
} catch {
    Set-Status 'echec' $_.Exception.Message
    throw
}
