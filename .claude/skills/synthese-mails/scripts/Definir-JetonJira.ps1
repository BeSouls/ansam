<#
.SYNOPSIS
    Enregistre le jeton d'API Atlassian qui permet de lire Jira sans passer par Claude.

.DESCRIPTION
    Demande l'adresse du compte Atlassian et le jeton, les essaie sur Jira, puis
    les enregistre dans <WorkDir>\jira.jeton. Le jeton est chiffré par Windows
    (DPAPI) : il n'est lisible que par ce compte, sur ce poste. Il ne s'affiche
    jamais et n'est écrit dans aucun journal. Sans ce fichier, ou si Jira refuse
    le jeton, Panoramix repasse par Claude et son connecteur Atlassian.

    Jeton à créer sur https://id.atlassian.com/manage-profile/security/api-tokens :
    « Create API token with scopes », application Jira, droits read:jira-work et
    read:jira-user. Il est limité à un an : relancer ce script après l'avoir renouvelé.

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Definir-JetonJira.ps1
#>
[CmdletBinding()]
param([string]$WorkDir = (Join-Path $env:USERPROFILE 'SyntheseMails'))

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'SyntheseMails.Commun.ps1')

$email = (Read-Host 'Adresse de ton compte Atlassian').Trim()
if ($email -notmatch '^[\w.+-]+@[\w-]+(\.[\w-]+)+$') { throw 'Adresse non valable.' }
$secure = Read-Host 'Jeton (rien ne s''affiche pendant la saisie)' -AsSecureString
if (-not $secure.Length) { throw 'Jeton vide.' }

$saved = [ordered]@{ email = $email; jeton = (ConvertFrom-SecureString $secure); enregistreLe = (Get-Date).ToString('o') }
$acces = ConvertTo-SmJiraAcces $saved
try {
    $moi = Invoke-SmJira $acces '/rest/api/3/myself'
    $essai = Invoke-SmJira $acces '/rest/api/3/search/jql' @{ jql = 'assignee = currentUser() ORDER BY updated DESC'; fields = @('summary'); maxResults = 1 }
} catch {
    throw "Jira refuse ce jeton : rien n'a été enregistré. $($_.Exception.Message)"
}
Write-SmJson (Join-Path $WorkDir 'jira.jeton') $saved
Write-Host "Jeton enregistré pour $($moi.displayName) : lecture de Jira vérifiée ($(@($essai.issues).Count) ticket lu)." -ForegroundColor Green
