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
    .\Lancer-SyntheseMails.ps1 -Force -PasDeMail   # actualisation (bouton « Actualiser » de la page)
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
    # Actualisation : analyse complète et registres mis à jour, sans mail. La synthèse du jour et la
    # date de la dernière synthèse ne changent pas ; le résultat va dans syntheses\actualisations.
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

Start-Transcript -Path (Join-Path $logDir "$today.log") -Append | Out-Null
$address = $null
$config = $null
$lockFile = $null
$outcome = [ordered]@{ debut = $start.ToString('o'); fin = $null; resultat = 'ignore'; message = $null; force = $Force.IsPresent }
try {
    # Une seule synthèse à la fois (tâche du matin, bouton de la page).
    $lockPath = Join-Path $WorkDir 'lancement.lock'
    $other = Read-SmJson $lockPath
    if ($other -and $other.pid -ne $PID -and (Get-Process -Id $other.pid -ErrorAction SilentlyContinue)) {
        Write-Host "Synthèse déjà en cours (processus $($other.pid))."
        $outcome = $null
        return
    }
    $lockFile = $lockPath
    Write-SmJson $lockFile ([ordered]@{ pid = $PID; debut = $start.ToString('o'); mail = -not $PasDeMail.IsPresent })

    $config = Get-SmConfig $WorkDir
    $etat = Get-SmEtat $WorkDir
    $plan = Get-SmPlan $config $etat $start
    if (-not $plan.lancer) {
        if (-not $Force) { Write-Host "Pas de synthèse : $($plan.raison)"; $outcome.message = $plan.raison; return }
        Write-Host "Lancement forcé malgré : $($plan.raison)"
        if (-not $plan.depuis) { $plan.depuis = $start.Date.AddDays(-1).AddHours(7) }
    }

    $exportArgs = @{ Days = $plan.jours; OutputDir = $exportDir }
    if ($Mailbox) { $exportArgs.Mailbox = $Mailbox }
    Write-Host "$(Get-Date -Format HH:mm:ss) Étape : export Outlook ($($plan.jours) jours)"
    $exportFile = & (Join-Path $PSScriptRoot 'Export-OutlookMails.ps1') @exportArgs | Select-Object -Last 1
    $threadFile = [IO.Path]::ChangeExtension($exportFile, $null).TrimEnd('.') + '-fils.txt'
    $address = (Read-SmJson $exportFile).address

    $contextFile = Join-Path $WorkDir 'contexte-ansam.json'
    Write-Host "$(Get-Date -Format HH:mm:ss) Étape : contexte AnsamDiag"
    & (Join-Path $PSScriptRoot 'Export-ContexteAnsam.ps1') -DataDir $config.dossierAnsamDiag -ExportFile $exportFile `
        -DossierPortefeuille $config.dossierPortefeuille -OutputFile $contextFile | Out-Null

    # Tickets Jira assignés (lecture seule) ; en cas d'échec, la synthèse continue sans eux.
    $jiraFile = Join-Path $WorkDir 'jira-assignes.json'
    Write-Host "$(Get-Date -Format HH:mm:ss) Étape : tickets Jira"
    # Statut réel des tickets des points ouverts : clos, rouvert, réassigné.
    $trackedKeys = @((Get-SmSuivi $WorkDir).points | Where-Object { $_.ticket -and ($_.statut -eq 'a_faire' -or $_.statut -eq 'attente_tiers') } |
        ForEach-Object { [string]$_.ticket })
    $jiraOk = & (Join-Path $PSScriptRoot 'Export-JiraAssignes.ps1') -OutputFile $jiraFile -ClaudeExe $ClaudeExe -Cles $trackedKeys | Select-Object -Last 1
    $jiraNote = $(if ($jiraOk) { "tickets Jira assignés : $jiraFile" }
                  elseif (Test-Path $jiraFile) { "tickets Jira assignés : $jiraFile (lecture du jour en échec : liste de la veille, à signaler)" }
                  else { "tickets Jira assignés : indisponibles aujourd'hui (à signaler dans les points d'attention)" })

    # Registres : les points restent ouverts d'une synthèse à l'autre, les éléments à garder aussi.
    $suivi = Get-SmSuivi $WorkDir
    $suiviFile = Join-Path $WorkDir 'suivi.json'
    # Chaque ticket Jira ouvert qui m'est assigné a un point, avant même que l'agent lise les mails.
    $jira = $(if ($jiraOk) { Read-SmJson $jiraFile })
    $contextData = Read-SmJson $contextFile
    if ($jira) {
        $sync = Sync-SmJira $suivi $jira $contextData
        Write-Host "Jira → suivi : $($sync.ajoutes) ajoutés, $($sync.regles) réglés, $($sync.rouverts) rouverts."
        if ($jira.complet -eq $false) { $jiraNote += " ATTENTION : liste Jira incomplète ($(@($jira.tickets).Count) sur $($jira.total)), à signaler" }
    }
    Save-SmSuivi $suivi $WorkDir
    $archivesFile = Join-Path $WorkDir 'archives.json'
    if (-not (Test-Path $archivesFile)) { Save-SmArchives (Get-SmArchives $WorkDir) $WorkDir }

    $since = $plan.depuis.ToString('dd.MM.yyyy HH:mm')
    $prompt = "Utilise le skill synthese-mails pour produire la synthèse matinale de ma boîte mail pro ANSAM. " +
              "Vue par fil : $threadFile ; détail complet : $exportFile ; contexte AnsamDiag : $contextFile ; " +
              "$jiraNote ; registre de suivi : $suiviFile ($(@($suivi.points).Count) points) ; " +
              "registre des éléments à garder : $archivesFile. " +
              "Nouveautés depuis la dernière synthèse : $since."
    $absences = @($plan.absences)
    if ($absences.Count) {
        $periods = ($absences | ForEach-Object { "$($_.libelle) du $($_.debut) au $($_.fin)" }) -join ', '
        $prompt += " Retour d'absence ($periods) : utilise la variante « retour d'absence » du modèle."
    }
    $prompt += " Réponds uniquement avec la synthèse en Markdown, suivie du bloc de suivi (<!-- SUIVI --> puis son tableau JSON) " +
               "et du bloc des éléments à garder (<!-- ARCHIVES --> puis son tableau JSON)."

    # Sortie de claude en UTF-8, sinon les accents sont perdus sous Windows PowerShell.
    [Console]::OutputEncoding = [Text.Encoding]::UTF8
    Push-Location $RepoDir
    try {
        Write-Host "$(Get-Date -Format HH:mm:ss) Étape : rédaction par Claude Code"
        $synthesis = & $ClaudeExe -p $prompt --add-dir $WorkDir --allowedTools 'Read,Grep,Glob,Skill'
        if ($LASTEXITCODE -ne 0) { throw "Claude Code a échoué (code $LASTEXITCODE) : $($synthesis -join ' ')" }
    } finally {
        Pop-Location
    }

    # La synthèse commence au premier titre ; tout préambule éventuel est retiré.
    $markdown = $synthesis -join "`n"
    $titleAt = [regex]::Match($markdown, '(?m)^# ')
    if ($titleAt.Success) { $markdown = $markdown.Substring($titleAt.Index) }

    # Blocs de fin (suivi, éléments à garder) : intégrés aux registres, retirés du mail.
    $raw = $markdown
    $firstBlock = [regex]::Match($markdown, '<!-- (SUIVI|ARCHIVES) -->')
    if ($firstBlock.Success) { $markdown = $markdown.Substring(0, $firstBlock.Index).TrimEnd() }
    try {
        $updates = Get-SmBlocJson $raw 'SUIVI'
        if ($null -eq $updates) { throw 'bloc absent' }
        $changes = Merge-SmSuivi $suivi $updates
        if ($jira) { [void](Sync-SmJira $suivi $jira $contextData) }   # Jira a le dernier mot sur ce qui est encore ouvert
        Save-SmSuivi $suivi $WorkDir
        Write-Host "Suivi : $(@($updates).Count) points revus, $(@($changes).Count) changements de statut, $(@($suivi.points).Count) points au registre."
    } catch {
        Write-Warning "Registre de suivi inchangé : $($_.Exception.Message)"
        $markdown += "`n`n> ⚠ Registre de suivi non mis à jour aujourd'hui ($($_.Exception.Message))."
    }
    try {
        $kept = Get-SmBlocJson $raw 'ARCHIVES'
        if ($null -eq $kept) { throw 'bloc absent' }
        $archives = Get-SmArchives $WorkDir
        $added = Merge-SmArchives $archives $kept
        Save-SmArchives $archives $WorkDir
        Write-Host "Éléments à garder : $(@($kept).Count) revus, $added nouveaux, $(@($archives.elements).Count) au registre."
    } catch {
        Write-Warning "Registre des éléments à garder inchangé : $($_.Exception.Message)"
        $markdown += "`n`n> ⚠ Variabilité / facturation / points d'attention non enregistrés aujourd'hui ($($_.Exception.Message))."
    }
    $markdown += "`n`n---`n[Ouvrir mon suivi]($(Get-SmPageUrl $WorkDir)) : marquer un point traité ou en attente, reporter la variabilité, traiter la facturation."
    # Une actualisation (-PasDeMail) ne remplace pas la synthèse du jour : fichier à part, et la date de
    # la dernière synthèse reste celle du dernier envoi, sinon celle du matin serait sautée ou raccourcie.
    $analysisOnly = $PasDeMail.IsPresent
    if ($analysisOnly) {
        $refreshDir = Join-Path $synthesisDir 'actualisations'
        New-Item -ItemType Directory -Force -Path $refreshDir | Out-Null
        $synthesisFile = Join-Path $refreshDir ('{0:yyyy-MM-dd_HHmm}.md' -f $start)
    } else {
        $synthesisFile = Join-Path $synthesisDir "$today.md"
    }
    [IO.File]::WriteAllText($synthesisFile, $markdown, (New-Object Text.UTF8Encoding($false)))
    Write-Host "Synthèse prête : $synthesisFile"

    if ($config.envoiMail -and -not $analysisOnly) {
        $subject = ((($markdown -split "`n")[0]) -replace '^#\s*', '').Trim()
        if ($subject -notlike 'Synthèse mails – *') { $subject = "Synthèse mails – $($start.ToString('dd.MM.yyyy'))" }
        Send-SelfMail $subject (ConvertTo-SmHtml $markdown) $address
    }

    $newEtat = [ordered]@{
        derniereSynthese = $etat.derniereSynthese
        dernierMode      = $etat.dernierMode
        dernierFichier   = $etat.dernierFichier
        derniereAnalyse  = $start.ToString('o')
    }
    if (-not $analysisOnly) {
        $newEtat.derniereSynthese = $start.ToString('o')
        $newEtat.dernierMode = $(if ($absences.Count) { 'retour' } else { 'quotidien' })
        $newEtat.dernierFichier = $synthesisFile
    }
    Save-SmEtat $newEtat $WorkDir
    $outcome.resultat = 'ok'
    $outcome.message = $(if ($analysisOnly) { 'Analyse actualisée (sans mail).' }
                         elseif ($config.envoiMail) { 'Synthèse envoyée.' }
                         else { 'Synthèse enregistrée (envoi par mail désactivé).' })
} catch {
    $outcome.resultat = 'echec'
    $outcome.message = $_.Exception.Message
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
    # Les déclenchements sans synthèse (déverrouillage, jour non prévu) n'écrasent pas le dernier vrai résultat.
    if ($outcome -and $outcome.resultat -ne 'ignore') {
        $outcome.fin = (Get-Date).ToString('o')
        Write-SmJson (Join-Path $WorkDir 'dernier-lancement.json') $outcome
    }
    if ($lockFile) { Remove-Item $lockFile -ErrorAction SilentlyContinue }
    Stop-Transcript | Out-Null
}
