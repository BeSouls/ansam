<#
.SYNOPSIS
    Produit la synthèse matinale des mails et me l'envoie par mail.

.DESCRIPTION
    1. Lit la configuration (<WorkDir>\config.json, réglée dans la page de
       Configurer-SyntheseMails.cmd) : jours d'envoi, absences, fenêtre.
       Pas de synthèse un jour non prévu ou pendant une absence ; au retour,
       une seule synthèse couvre toute la période depuis la dernière.
    2. Export-OutlookMails.ps1 : export de la boîte principale.
    3. Export-ContexteAnsam.ps1 : fiches clients et tickets AnsamDiag.
    4. Export-JiraAssignes.ps1 : mes tickets Jira ouverts (lecture seule).
    5. claude -p (Claude Code non interactif, lecture seule) depuis ce repo,
       pour que le skill synthese-mails soit chargé, avec le registre de
       suivi (suivi.json). La réponse se termine par un bloc de suivi, qui
       est intégré au registre (Merge-SmSuivi) puis retiré du mail.
    6. Synthèse gardée dans <WorkDir>\syntheses\AAAA-MM-JJ.md et envoyée par
       mail à ma propre adresse via Outlook. En cas d'échec, un mail
       « Synthèse mails – échec » le signale. Journal dans <WorkDir>\logs.

    Prérequis : Outlook classique ouvert ou ouvrable, Claude Code installé et
    connecté (lancer « claude » une fois à la main).

.EXAMPLE
    .\Lancer-SyntheseMails.ps1              # comme la tâche planifiée
    .\Lancer-SyntheseMails.ps1 -Force       # même un jour non prévu ou déjà traité
    .\Lancer-SyntheseMails.ps1 -Force -PasDeMail
#>
[CmdletBinding()]
param(
    [string]$WorkDir = (Join-Path $env:USERPROFILE 'SyntheseMails'),
    [string]$Mailbox,
    # Par défaut : la racine de ce repo.
    [string]$RepoDir,
    # Par défaut : « claude » du PATH, sinon l'installation native
    # (%USERPROFILE%\.local\bin), que l'installeur n'ajoute pas toujours au PATH.
    [string]$ClaudeExe,
    # Lance même un jour non prévu, pendant une absence ou déjà traité.
    [switch]$Force,
    # N'envoie pas le mail (la synthèse est seulement enregistrée).
    [switch]$PasDeMail
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'SyntheseMails.Commun.ps1')

# $PSScriptRoot est vide dans les valeurs par défaut du param() sous Windows PowerShell 5.1.
if (-not $RepoDir) { $RepoDir = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path }
$ClaudeExe = Resolve-SmClaude $ClaudeExe

function Send-SelfMail([string]$Subject, [string]$HtmlBody, [string]$Address) {
    $outlook = New-Object -ComObject Outlook.Application
    if (-not $Address) {
        $Address = $outlook.Session.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress
    }
    if (-not $Address) { throw 'Adresse de la boîte introuvable : mail non envoyé.' }
    $mail = $outlook.CreateItem(0)   # olMailItem
    $mail.To = $Address              # toujours ma propre adresse, jamais un tiers
    $mail.Subject = $Subject
    $mail.HTMLBody = $HtmlBody
    $mail.Send()
    Write-Host "Synthèse envoyée à $Address."
}

$start = Get-Date
$today = '{0:yyyy-MM-dd}' -f $start
$exportDir = Join-Path $WorkDir 'exports'
$synthesisDir = Join-Path $WorkDir 'syntheses'
$logDir = Join-Path $WorkDir 'logs'
New-Item -ItemType Directory -Force -Path $exportDir, $synthesisDir, $logDir | Out-Null

$config = Get-SmConfig $WorkDir
$etat = Get-SmEtat $WorkDir
$plan = Get-SmPlan $config $etat $start

Start-Transcript -Path (Join-Path $logDir "$today.log") -Append | Out-Null
$address = $null
try {
    if (-not $plan.lancer) {
        if (-not $Force) { Write-Host "Pas de synthèse : $($plan.raison)"; return }
        Write-Host "Lancement forcé malgré : $($plan.raison)"
        if (-not $plan.depuis) { $plan.depuis = $start.Date.AddDays(-1).AddHours(7) }
    }

    $exportArgs = @{ Days = $plan.jours; OutputDir = $exportDir }
    if ($Mailbox) { $exportArgs.Mailbox = $Mailbox }
    $exportFile = & (Join-Path $PSScriptRoot 'Export-OutlookMails.ps1') @exportArgs | Select-Object -Last 1
    $threadFile = [IO.Path]::ChangeExtension($exportFile, $null).TrimEnd('.') + '-fils.txt'
    $address = (Read-SmJson $exportFile).address

    $contextFile = Join-Path $WorkDir 'contexte-ansam.json'
    & (Join-Path $PSScriptRoot 'Export-ContexteAnsam.ps1') -DataDir $config.dossierAnsamDiag -ExportFile $exportFile `
        -DossierPortefeuille $config.dossierPortefeuille -OutputFile $contextFile | Out-Null

    # Tickets Jira assignés (lecture seule) ; en cas d'échec, la synthèse continue sans eux.
    $jiraFile = Join-Path $WorkDir 'jira-assignes.json'
    $jiraOk = & (Join-Path $PSScriptRoot 'Export-JiraAssignes.ps1') -OutputFile $jiraFile -ClaudeExe $ClaudeExe | Select-Object -Last 1
    $jiraNote = $(if ($jiraOk) { "tickets Jira assignés : $jiraFile" }
                  elseif (Test-Path $jiraFile) { "tickets Jira assignés : $jiraFile (lecture du jour en échec : liste de la veille, à signaler)" }
                  else { "tickets Jira assignés : indisponibles aujourd'hui (à signaler dans les points d'attention)" })

    # Registre de suivi : les points restent ouverts d'une synthèse à l'autre.
    $suivi = Get-SmSuivi $WorkDir
    $suiviFile = Join-Path $WorkDir 'suivi.json'
    if (-not (Test-Path $suiviFile)) { Save-SmSuivi $suivi $WorkDir }

    $since = $plan.depuis.ToString('dd.MM.yyyy HH:mm')
    $prompt = "Utilise le skill synthese-mails pour produire la synthèse matinale de ma boîte mail pro ANSAM. " +
              "Vue par fil : $threadFile ; détail complet : $exportFile ; contexte AnsamDiag : $contextFile ; " +
              "$jiraNote ; registre de suivi : $suiviFile ($(@($suivi.points).Count) points). " +
              "Nouveautés depuis la dernière synthèse : $since."
    $absences = @($plan.absences)
    if ($absences.Count) {
        $periods = ($absences | ForEach-Object { "$($_.libelle) du $($_.debut) au $($_.fin)" }) -join ', '
        $prompt += " Retour d'absence ($periods) : utilise la variante « retour d'absence » du modèle."
    }
    $prompt += " Réponds uniquement avec la synthèse en Markdown, suivie du bloc de suivi (marqueur <!-- SUIVI --> puis le tableau JSON)."

    # Sortie de claude en UTF-8, sinon les accents sont perdus sous Windows PowerShell.
    [Console]::OutputEncoding = [Text.Encoding]::UTF8
    Push-Location $RepoDir
    try {
        $synthesis = & $ClaudeExe -p $prompt --add-dir $WorkDir --allowedTools 'Read,Grep,Glob,Skill'
        if ($LASTEXITCODE -ne 0) { throw "Claude Code a échoué (code $LASTEXITCODE) : $($synthesis -join ' ')" }
    } finally {
        Pop-Location
    }

    # La synthèse commence au premier titre ; tout préambule éventuel est retiré.
    $markdown = $synthesis -join "`n"
    $titleAt = [regex]::Match($markdown, '(?m)^# ')
    if ($titleAt.Success) { $markdown = $markdown.Substring($titleAt.Index) }

    # Bloc de suivi : intégré au registre, retiré du mail.
    $marker = $markdown.IndexOf('<!-- SUIVI -->')
    if ($marker -lt 0) {
        Write-Warning 'Bloc de suivi absent : registre inchangé.'
        $markdown += "`n`n> ⚠ Registre de suivi non mis à jour aujourd'hui (bloc de suivi absent)."
    } else {
        $block = $markdown.Substring($marker)
        $markdown = $markdown.Substring(0, $marker).TrimEnd()
        try {
            $jsonStart = $block.IndexOf('[')
            $jsonEnd = $block.LastIndexOf(']')
            if ($jsonStart -lt 0 -or $jsonEnd -le $jsonStart) { throw 'tableau JSON introuvable' }
            $updates = $block.Substring($jsonStart, $jsonEnd - $jsonStart + 1) | ConvertFrom-Json
            $changes = Merge-SmSuivi $suivi $updates
            Save-SmSuivi $suivi $WorkDir
            Write-Host "Suivi : $(@($updates).Count) points revus, $(@($changes).Count) changements de statut, $(@($suivi.points).Count) points au registre."
        } catch {
            Write-Warning "Bloc de suivi illisible, registre inchangé : $($_.Exception.Message)"
            $markdown += "`n`n> ⚠ Registre de suivi non mis à jour aujourd'hui (bloc de suivi illisible)."
        }
    }
    $markdown += "`n`n---`n[Ouvrir mon suivi]($(Get-SmPageUrl $WorkDir)) pour marquer un point traité, en attente d'un tiers ou sans objet."
    $synthesisFile = Join-Path $synthesisDir "$today.md"
    [IO.File]::WriteAllText($synthesisFile, $markdown, (New-Object Text.UTF8Encoding($false)))
    Write-Host "Synthèse prête : $synthesisFile"

    if ($config.envoiMail -and -not $PasDeMail) {
        $subject = ((($markdown -split "`n")[0]) -replace '^#\s*', '').Trim()
        if ($subject -notlike 'Synthèse mails – *') { $subject = "Synthèse mails – $($start.ToString('dd.MM.yyyy'))" }
        Send-SelfMail $subject (ConvertTo-SmHtml $markdown) $address
    }

    $etat = [ordered]@{
        derniereSynthese = $start.ToString('o')
        dernierMode      = $(if ($absences.Count) { 'retour' } else { 'quotidien' })
        dernierFichier   = $synthesisFile
    }
    Save-SmEtat $etat $WorkDir
} catch {
    Write-Warning "Échec : $($_.Exception.Message)"
    if ($config.envoiMail -and -not $PasDeMail) {
        try {
            $body = "<p>La synthèse du $($start.ToString('dd.MM.yyyy')) n'a pas pu être produite.</p>" +
                    "<p><b>Erreur :</b> $([Net.WebUtility]::HtmlEncode($_.Exception.Message))</p>" +
                    "<p>Journal : <code>$(Join-Path $logDir "$today.log")</code></p>"
            Send-SelfMail "Synthèse mails – échec du $($start.ToString('dd.MM.yyyy'))" $body $address
        } catch { Write-Warning "Mail d'échec non envoyé : $($_.Exception.Message)" }
    }
    throw
} finally {
    Stop-Transcript | Out-Null
}
