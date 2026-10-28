<#
.SYNOPSIS
    Exporte les mails Outlook puis fait rédiger la synthèse matinale par
    Claude Code, en local.

.DESCRIPTION
    1. Export-OutlookMails.ps1 : export JSON de la boîte principale.
    2. claude -p (Claude Code en mode non interactif), lancé depuis ce repo
       pour que le skill synthese-mails soit chargé, en lecture seule.
    3. Synthèse écrite dans <WorkDir>\syntheses\AAAA-MM-JJ.md, journal dans
       <WorkDir>\logs.

    Prérequis : Outlook classique ouvert ou ouvrable, Claude Code installé et
    connecté (lancer « claude » une fois à la main).
#>
[CmdletBinding()]
param(
    [string]$WorkDir = (Join-Path $env:USERPROFILE 'SyntheseMails'),
    [int]$Days = 14,
    [string]$Mailbox,
    [string]$RepoDir = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path,
    [string]$ClaudeExe = 'claude'
)

$ErrorActionPreference = 'Stop'
$today = '{0:yyyy-MM-dd}' -f (Get-Date)
$exportDir = Join-Path $WorkDir 'exports'
$synthesisDir = Join-Path $WorkDir 'syntheses'
$logDir = Join-Path $WorkDir 'logs'
New-Item -ItemType Directory -Force -Path $exportDir, $synthesisDir, $logDir | Out-Null

Start-Transcript -Path (Join-Path $logDir "$today.log") -Append | Out-Null
try {
    $exportArgs = @{ Days = $Days; OutputDir = $exportDir }
    if ($Mailbox) { $exportArgs.Mailbox = $Mailbox }
    $exportFile = & (Join-Path $PSScriptRoot 'Export-OutlookMails.ps1') @exportArgs | Select-Object -Last 1

    $prompt = "Utilise le skill synthese-mails pour produire la synthèse matinale de ma boîte mail pro ANSAM " +
              "à partir de l'export local $exportFile. Réponds uniquement avec la synthèse en Markdown."

    # Sortie de claude en UTF-8, sinon les accents sont perdus sous Windows PowerShell.
    [Console]::OutputEncoding = [Text.Encoding]::UTF8
    Push-Location $RepoDir
    try {
        $synthesis = & $ClaudeExe -p $prompt --add-dir $exportDir --allowedTools 'Read,Grep,Glob,Skill'
        if ($LASTEXITCODE -ne 0) { throw "Claude Code a échoué (code $LASTEXITCODE)." }
    } finally {
        Pop-Location
    }

    $synthesisFile = Join-Path $synthesisDir "$today.md"
    [IO.File]::WriteAllText($synthesisFile, ($synthesis -join "`n"), (New-Object Text.UTF8Encoding($false)))
    Write-Host "Synthèse prête : $synthesisFile"
} finally {
    Stop-Transcript | Out-Null
}
