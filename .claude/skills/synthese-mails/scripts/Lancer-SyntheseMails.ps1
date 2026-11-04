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
    Le jour de la réunion squad (hors absence et jour férié), Preparer-Squad.ps1
       prépare mes lignes du formulaire ; la synthèse y mène par un lien.
    6. Synthèse gardée dans <WorkDir>\syntheses\AAAA-MM-JJ.md et envoyée par
       mail à ma propre adresse via Outlook. En cas d'échec, un mail
       « Synthèse mails – échec » le signale. Journal dans <WorkDir>\logs.

    Prérequis : Outlook classique ouvert ou ouvrable, Claude Code installé et
    connecté (lancer « claude » une fois à la main).

.EXAMPLE
    .\Lancer-SyntheseMails.ps1              # comme la tâche planifiée
    .\Lancer-SyntheseMails.ps1 -Force       # même un jour non prévu ou déjà traité
    .\Lancer-SyntheseMails.ps1 -Force -PasDeMail   # actualisation (bouton « Actualiser » de la page)
    .\Lancer-SyntheseMails.ps1 -Force -Essai       # synthèse d'essai, sans mail
    .\Lancer-SyntheseMails.ps1 -Force -Complet     # analyse complète, même un jour incrémental

    Trois modes d'analyse (section du skill correspondante) :
    - complète : toute la fenêtre (14 jours) est relue. Le jour « analyseCompleteJour » (lundi),
      au retour d'absence, sans analyse depuis plus de 3 jours, ou avec -Complet ;
    - incrémentale (les autres matins) : seuls les fils qui ont du nouveau depuis la dernière
      analyse sont relus ; les points ouverts sont repris du registre ;
    - actualisation (-PasDeMail, analyse déjà faite le jour même) : idem, compte rendu court.
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
    # Actualisation : registres mis à jour, sans mail. La synthèse du jour et la date de la
    # dernière synthèse ne changent pas ; le résultat va dans syntheses\actualisations.
    [switch]$PasDeMail,
    # Analyse complète de la fenêtre, même un jour où la synthèse serait incrémentale.
    [switch]$Complet,
    # Essai : synthèse produite comme le matin (même mode), gardée dans syntheses\essais,
    # sans mail ; la date de la dernière synthèse ne change pas.
    [switch]$Essai
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

    # Tickets Jira assignés (lecture seule), lus en parallèle de l'export Outlook : les deux sont
    # indépendants. En cas d'échec, la synthèse continue sans eux.
    $jiraFile = Join-Path $WorkDir 'jira-assignes.json'
    # Statut réel des tickets des points ouverts : clos, rouvert, réassigné.
    $trackedKeys = @((Get-SmSuivi $WorkDir).points | Where-Object { $_.ticket -and ($_.statut -eq 'a_faire' -or $_.statut -eq 'attente_tiers') } |
        ForEach-Object { [string]$_.ticket })
    Write-Host "$(Get-Date -Format HH:mm:ss) Étape : export Outlook ($($plan.jours) jours) et tickets Jira"
    $activityFile = Join-Path $WorkDir 'jira-activite.json'
    $jiraJob = Start-Job -ScriptBlock {
        param($Script, $ActivityScript, $OutputFile, $ActivityFile, $Claude, $Keys)
        Set-ExecutionPolicy -Scope Process Bypass -Force
        $ok = & $Script -OutputFile $OutputFile -ClaudeExe $Claude -Cles $Keys | Select-Object -Last 1
        # Commentaires : les mails n'en disent rien (voir Export-JiraActivite.ps1). Un échec n'arrête rien.
        $activity = $null
        if ($ok) {
            try { $activity = & $ActivityScript -JiraFile $OutputFile -OutputFile $ActivityFile -ClaudeExe $Claude | Select-Object -Last 1 }
            catch { Write-Warning "Commentaires Jira non lus : $($_.Exception.Message)" }
        }
        [pscustomobject]@{ jira = $ok; activite = $activity }
    } -ArgumentList (Join-Path $PSScriptRoot 'Export-JiraAssignes.ps1'), (Join-Path $PSScriptRoot 'Export-JiraActivite.ps1'), $jiraFile, $activityFile, $ClaudeExe, $trackedKeys

    $exportArgs = @{ Days = $plan.jours; OutputDir = $exportDir }
    if ($Mailbox) { $exportArgs.Mailbox = $Mailbox }
    $exportFile = & (Join-Path $PSScriptRoot 'Export-OutlookMails.ps1') @exportArgs | Select-Object -Last 1
    $threadFile = [IO.Path]::ChangeExtension($exportFile, $null).TrimEnd('.') + '-fils.txt'
    $address = (Read-SmJson $exportFile).address

    $contextFile = Join-Path $WorkDir 'contexte-ansam.json'
    Write-Host "$(Get-Date -Format HH:mm:ss) Étape : contexte AnsamDiag"
    & (Join-Path $PSScriptRoot 'Export-ContexteAnsam.ps1') -DataDir $config.dossierAnsamDiag -ClientsDir $config.dossierClients -ExportFile $exportFile `
        -DossierPortefeuille $config.dossierPortefeuille -OutputFile $contextFile | Out-Null

    Write-Host "$(Get-Date -Format HH:mm:ss) Étape : tickets Jira"
    $jiraResult = $jiraJob | Wait-Job | Receive-Job | Select-Object -Last 1
    Remove-Job $jiraJob -Force
    $jiraOk = $jiraResult.jira
    $jiraNote = $(if ($jiraOk) { "tickets Jira assignés : $jiraFile" }
                  elseif (Test-Path $jiraFile) { "tickets Jira assignés : $jiraFile (lecture du jour en échec : liste de la veille, à signaler)" }
                  else { "tickets Jira assignés : indisponibles aujourd'hui (à signaler dans les points d'attention)" })
    $activity = $(if ($jiraResult.activite) { Read-SmJson $activityFile })
    $activityGaps = @($activity.echecs | ForEach-Object { $_ } | Where-Object { $_ })
    $jiraNote += $(if (-not $activity) { " ; commentaires Jira non lus aujourd'hui (à signaler dans les points d'attention)" }
                   elseif ($activityGaps.Count) { " ; commentaires Jira : $activityFile (non relus : $($activityGaps -join ', '), à signaler)" }
                   else { " ; commentaires Jira : $activityFile" })

    # Registres : les points restent ouverts d'une synthèse à l'autre, les éléments à garder aussi.
    $suivi = Get-SmSuivi $WorkDir
    $suiviFile = Join-Path $WorkDir 'suivi.json'
    # Chaque ticket Jira ouvert qui m'est assigné a un point, avant même que l'agent lise les mails.
    $jira = $(if ($jiraOk) { Read-SmJson $jiraFile })
    $contextData = Read-SmJson $contextFile
    if ($jira) {
        $sync = Sync-SmJira $suivi $jira $contextData
        Update-SmActiviteJira $suivi $activity ([string]$jira.moi)
        Write-Host "Jira → suivi : $($sync.ajoutes) ajoutés, $($sync.regles) réglés, $($sync.rouverts) rouverts, $($sync.changes) changés."
        if ($jira.complet -eq $false) { $jiraNote += " ATTENTION : liste Jira incomplète ($(@($jira.tickets).Count) sur $($jira.total)), à signaler" }
    }
    Save-SmSuivi $suivi $WorkDir
    $archivesFile = Join-Path $WorkDir 'archives.json'
    if (-not (Test-Path $archivesFile)) { Save-SmArchives (Get-SmArchives $WorkDir) $WorkDir }

    # Mode d'analyse. Les registres gardent l'état de chaque point d'une analyse à l'autre : hors
    # analyse complète, seuls les fils qui ont du nouveau depuis la dernière analyse sont relus, en
    # entier. L'analyse complète hebdomadaire (et au retour d'absence) rattrape un fil mal classé.
    $absences = @($plan.absences)
    $lastAnalysis = ConvertTo-SmDateTime $etat.derniereAnalyse
    $recent = $lastAnalysis -and ($start - $lastAnalysis).TotalDays -lt 3
    $mode = $(if ($PasDeMail -and $lastAnalysis -and $lastAnalysis.Date -eq $start.Date) { 'actualisation' }
              elseif ($PasDeMail -or $Complet -or -not $recent -or $absences.Count -or (Get-SmJourSemaine $start) -eq $config.analyseCompleteJour) { 'complete' }
              else { 'incrementale' })
    $delta = $null
    if ($mode -ne 'complete') {
        $deltaSince = $lastAnalysis.AddMinutes(-5)
        $deltaFile = [IO.Path]::ChangeExtension($exportFile, $null).TrimEnd('.') + '-delta.txt'
        $delta = New-SmFilsDelta $threadFile $deltaSince $deltaFile
        Write-Host "Analyse $mode : $($delta.nouveaux) fils avec du nouveau depuis $($deltaSince.ToString('dd.MM HH:mm')) ($($delta.selection) fils relus sur $($delta.total))."
    } else {
        Write-Host "Analyse complète de la fenêtre."
    }

    # Points dont une source a bougé depuis leur dernière revue : la synthèse les reprend, quel que soit le mode.
    $toReview = Update-SmARevoir $suivi $activity $contextData $(if ($delta) { $delta.cles })
    $reviewFile = Join-Path $WorkDir 'a-revoir.json'
    Write-SmJson $reviewFile @($toReview)
    Save-SmSuivi $suivi $WorkDir
    Write-Host "Points à revoir : $($toReview.Count)."

    $sources = "contexte AnsamDiag : $contextFile ; $jiraNote ; registre de suivi : $suiviFile ($(@($suivi.points).Count) points) ; " +
               "registre des éléments à garder : $archivesFile ; " +
               "points à reprendre obligatoirement, avec ce qui a bougé et leurs derniers commentaires Jira : $reviewFile ($($toReview.Count) points) ; " +
               "annuaire des tiers (canal officiel quand c'est ce tiers qu'il faut contacter) : $(Format-SmContactsTiers $config.contactsTiers)."
    $deltaNote = $(if ($delta -and $delta.nouveaux) { "Fils avec du nouveau depuis $($deltaSince.ToString('dd.MM.yyyy HH:mm')), complets : $deltaFile ($($delta.selection) fils) ; " }
                   elseif ($delta) { "Aucun fil n'a de message depuis $($deltaSince.ToString('dd.MM.yyyy HH:mm')) ; " })
    if ($mode -eq 'incrementale') {
        $prompt = "Utilise le skill synthese-mails pour produire la synthèse matinale de ma boîte mail pro ANSAM, en mode « Synthèse incrémentale » (section du skill). " +
                  $deltaNote + "vue de tous les fils : $threadFile et détail complet : $exportFile, pour le contexte seulement ; $sources " +
                  "Nouveautés depuis la dernière synthèse : $($plan.depuis.ToString('dd.MM.yyyy HH:mm')). " +
                  "Réponds uniquement avec la synthèse en Markdown, suivie du bloc de suivi (<!-- SUIVI --> puis son tableau JSON, " +
                  "seulement les points nouveaux ou modifiés) et du bloc des éléments à garder (<!-- ARCHIVES --> puis son tableau JSON, " +
                  "seulement les éléments nouveaux ou modifiés)."
    } elseif ($mode -eq 'actualisation') {
        $prompt = "Utilise le skill synthese-mails en mode « Actualisation » (section du skill) sur ma boîte mail pro ANSAM. " +
                  $deltaNote + "vue de tous les fils : $threadFile ; détail complet : $exportFile ; $sources " +
                  "Réponds uniquement avec le compte rendu court en Markdown, suivi du bloc de suivi (<!-- SUIVI --> puis son tableau JSON, " +
                  "seulement les points nouveaux ou modifiés) et du bloc des éléments à garder (<!-- ARCHIVES --> puis son tableau JSON, " +
                  "seulement les éléments nouveaux ou modifiés)."
    } else {
        $prompt = "Utilise le skill synthese-mails pour produire la synthèse matinale de ma boîte mail pro ANSAM. " +
                  "Vue par fil : $threadFile ; détail complet : $exportFile ; $sources " +
                  "Nouveautés depuis la dernière synthèse : $($plan.depuis.ToString('dd.MM.yyyy HH:mm'))."
        if ($absences.Count) {
            $periods = ($absences | ForEach-Object { "$($_.libelle) du $($_.debut) au $($_.fin)" }) -join ', '
            $prompt += " Retour d'absence ($periods) : utilise la variante « retour d'absence » du modèle."
        }
        $prompt += " Réponds uniquement avec la synthèse en Markdown, suivie du bloc de suivi (<!-- SUIVI --> puis son tableau JSON) " +
                   "et du bloc des éléments à garder (<!-- ARCHIVES --> puis son tableau JSON)."
    }
    if ($mode -eq 'actualisation' -and -not $delta.nouveaux -and -not $toReview.Count) {
        $markdown = "# Actualisation – $($start.ToString('dd.MM.yyyy HH:mm'))`n`nAucun nouveau message ni activité sur les tickets depuis $($deltaSince.ToString('HH:mm')) : " +
                    "statuts Jira et prises de contact mis à jour, rien d'autre à relire."
        $raw = $null
        $completed = Update-SmContacts $suivi $config
        Save-SmSuivi $suivi $WorkDir
    } else {
        # Sortie de claude en UTF-8, sinon les accents sont perdus sous Windows PowerShell.
        [Console]::OutputEncoding = [Text.Encoding]::UTF8
        Push-Location $RepoDir
        try {
            Write-Host "$(Get-Date -Format HH:mm:ss) Étape : $(@{ complete = 'rédaction'; incrementale = 'rédaction (nouveautés)'; actualisation = 'analyse des nouveautés' }[$mode]) par Claude Code"
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
    }
    if ($raw) { try {
        $updates = Get-SmBlocJson $raw 'SUIVI'
        if ($null -eq $updates) { throw 'bloc absent' }
        $changes = Merge-SmSuivi $suivi $updates $start.ToString('o')
        if ($jira) { [void](Sync-SmJira $suivi $jira $contextData) }   # Jira a le dernier mot sur ce qui est encore ouvert
        $completed = Update-SmContacts $suivi $config
        if ($completed) { Write-Host "Prises de contact complétées : $completed." }
        Save-SmSuivi $suivi $WorkDir
        $unreviewed = @($suivi.points | Where-Object { $_.aRevoir -and ($_.statut -eq 'a_faire' -or $_.statut -eq 'attente_tiers') })
        Write-Host "Suivi : $(@($updates).Count) points revus, $(@($changes).Count) changements de statut, $(@($suivi.points).Count) points au registre, $($unreviewed.Count) restés à revoir."
        if ($unreviewed.Count) { $markdown += "`n`n> ⚠ $($unreviewed.Count) points à revérifier toi-même, non repris par l'analyse : $(($unreviewed | ForEach-Object { $_.id }) -join ', ')." }
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
    } }
    # Réunion squad : préparée le matin du jour de réunion, hors absence et jour férié ; le mail y mène.
    if ($config.squadActif -and -not $PasDeMail -and -not $Essai -and (Get-SmJourSemaine $start) -eq $config.squadJour -and -not (Get-SmSquadEmpechement $config $start)) {
        $squadUrl = "$(Get-SmPageUrl $WorkDir)#squad"
        Write-Host "$(Get-Date -Format HH:mm:ss) Étape : réunion squad"
        try {
            & (Join-Path $PSScriptRoot 'Preparer-Squad.ps1') -Date $today -WorkDir $WorkDir -RepoDir $RepoDir -ClaudeExe $ClaudeExe | Out-Null
            $markdown += "`n`n**Réunion squad du $($start.ToString('dd.MM'))** : tes lignes sont prêtes, à relire et copier dans OneNote. [Ouvrir la préparation]($squadUrl)"
        } catch {
            Write-Warning "Réunion squad non préparée : $($_.Exception.Message)"
            $markdown += "`n`n> ⚠ Réunion squad non préparée ($($_.Exception.Message)). [La préparer dans Panoramix]($squadUrl)"
        }
    }
    $markdown += "`n`n---`n[Ouvrir Panoramix]($(Get-SmPageUrl $WorkDir)) : marquer un point traité ou en attente, reporter la variabilité, traiter la facturation."
    # Une actualisation (-PasDeMail) ne remplace pas la synthèse du jour : fichier à part, et la date de
    # la dernière synthèse reste celle du dernier envoi, sinon celle du matin serait sautée ou raccourcie.
    # Un essai (-Essai) non plus : fichier dans syntheses\essais, sans mail.
    $analysisOnly = $PasDeMail.IsPresent -or $Essai.IsPresent
    if ($analysisOnly) {
        $refreshDir = Join-Path $synthesisDir $(if ($Essai) { 'essais' } else { 'actualisations' })
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
        # Ce que la page affiche pour dire sur quoi reposent les statuts.
        sources          = [ordered]@{
            mails        = $start.ToString('o')
            jira         = $(if ($jiraOk) { $jira.lusLe })
            commentaires = $(if ($activity) { $activity.lusLe })
            nonRelus     = $activityGaps
        }
    }
    if (-not $analysisOnly) {
        $newEtat.derniereSynthese = $start.ToString('o')
        $newEtat.dernierMode = $(if ($absences.Count) { 'retour' } else { 'quotidien' })
        $newEtat.dernierFichier = $synthesisFile
    }
    Save-SmEtat $newEtat $WorkDir
    $outcome.resultat = 'ok'
    $outcome.message = $(if ($Essai) { "Synthèse d'essai enregistrée (sans mail)." }
                         elseif ($analysisOnly) { 'Analyse actualisée (sans mail).' }
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
