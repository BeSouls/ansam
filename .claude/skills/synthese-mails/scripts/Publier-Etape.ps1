<#
.SYNOPSIS
    Vérifie dans Jira, en lecture seule, ce que j'ai publié d'une clôture préparée.

.DESCRIPTION
    Lancé par la page (onglet « Publier ») en arrière-plan. Pour chaque ticket de la file (ou -Cle) :
    commentaire client public, note interne, worklogs jour par jour et statut. Le résultat va dans
    <WorkDir>\publications.json ; l'état du travail dans <WorkDir>\travaux\publier.json.
    Jeton Jira en lecture seule (Definir-JetonJira.ps1) : rien n'est écrit dans Jira.

.EXAMPLE
    .\Publier-Etape.ps1 -Cle ASD-27334
    .\Publier-Etape.ps1            # toute la file non vérifiée
#>
[CmdletBinding()]
param(
    [string]$WorkDir = (Join-Path $env:USERPROFILE 'SyntheseMails'),
    [string]$Cle = ''
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'SyntheseMails.Commun.ps1')
. (Join-Path $PSScriptRoot 'SyntheseMails.Publier.ps1')

$stateFile = Join-Path $WorkDir 'travaux\publier.json'
$started = (Get-Date).ToString('o')
function Write-Etat([string]$Etat, [string]$Message) {
    Write-SmJson $stateFile ([ordered]@{ travail = 'publier'; id = $Cle; etat = $Etat; message = $Message; le = $started; fin = $(if ($Etat -ne 'en_cours') { (Get-Date).ToString('o') } else { $null }); pid = $PID })
}

try {
    Write-Etat 'en_cours' 'Lecture de Jira'
    $acces = Get-SmJiraAcces $WorkDir
    if (-not $acces) { throw 'Pas de jeton Jira : Definir-JetonJira.ps1 pour en enregistrer un (lecture seule).' }
    $me = Invoke-SmJira $acces '/rest/api/3/myself'
    $config = Get-SmConfig $WorkDir
    $list = @(Get-SmPublications $config $WorkDir (Get-SmSuivi $WorkDir) | Where-Object {
        if ($Cle) { $_.cle -eq $Cle } else { -not ($_.verification -and $_.verification.complet) }
    })
    $done = 0; $complete = 0; $errors = @()
    foreach ($pub in $list) {
        Write-Etat 'en_cours' "Vérification de $($pub.cle) ($($done + 1)/$($list.Count))"
        try {
            $result = Test-SmPublication $acces $pub ([string]$me.accountId)
            # Relu juste avant d'écrire : un clic dans la page pendant la vérification n'est pas perdu.
            $etat = Get-SmPublicationsEtat $WorkDir
            $entry = $etat[$pub.cle]
            if (-not $entry -or [string]$entry.fichierLe -ne [string]$pub.fichierLe) { $entry = [ordered]@{ cle = $pub.cle; fichierLe = $pub.fichierLe } }
            $entry.verification = $result
            # Tout est dans Jira sans que je l'aie signalé : c'est publié.
            if ($result.complet -and -not $entry.publieLe) { $entry.publieLe = $result.le }
            $etat[$pub.cle] = $entry
            Save-SmPublicationsEtat $etat $WorkDir
            if ($result.complet) { $complete++ }
        } catch { $errors += "$($pub.cle) : $($_.Exception.Message)" }
        $done++
    }
    $message = "$done ticket$(if ($done -gt 1) { 's' }) vérifié$(if ($done -gt 1) { 's' }), $complete complet$(if ($complete -gt 1) { 's' })" +
               $(if ($errors.Count) { ' ; ' + ($errors -join ' ; ') } else { '' })
    Write-Etat $(if ($errors.Count -and $errors.Count -eq $list.Count) { 'echec' } else { 'ok' }) $message
} catch {
    Write-Etat 'echec' $_.Exception.Message
    exit 1
}
