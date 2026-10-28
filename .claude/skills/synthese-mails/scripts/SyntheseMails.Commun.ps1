<#
.SYNOPSIS
    Fonctions communes de la synthèse mails : configuration, état, calendrier
    (jours d'envoi, absences) et mise en forme HTML. À charger par
    dot-sourcing : . "$PSScriptRoot\SyntheseMails.Commun.ps1"
#>

$SmWorkDir = Join-Path $env:USERPROFILE 'SyntheseMails'
$SmInvariant = [Globalization.CultureInfo]::InvariantCulture
$SmJours = @{ 1 = 'lundi'; 2 = 'mardi'; 3 = 'mercredi'; 4 = 'jeudi'; 5 = 'vendredi'; 6 = 'samedi'; 7 = 'dimanche' }

function Get-SmDefaultConfig {
    [ordered]@{
        actif               = $true
        heure               = '07:10'
        joursEnvoi          = @(1, 2, 3, 4, 5)   # 1 = lundi … 7 = dimanche
        fenetreJours        = 14
        envoiMail           = $true
        dossierAnsamDiag    = Join-Path $env:USERPROFILE 'OneDrive - ANSAM group\Documents\AnsamDiag-Donnees'
        dossierPortefeuille = 'Boîte de réception\MES CLIENTS'
        absences            = @()                # [{ debut: 'AAAA-MM-JJ'; fin: 'AAAA-MM-JJ'; libelle: '…' }]
    }
}

function Read-SmJson([string]$Path) {
    if (-not (Test-Path $Path)) { return $null }
    [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8) | ConvertFrom-Json
}

function Write-SmJson([string]$Path, $Value) {
    New-Item -ItemType Directory -Force -Path (Split-Path $Path) | Out-Null
    [IO.File]::WriteAllText($Path, ($Value | ConvertTo-Json -Depth 6), (New-Object Text.UTF8Encoding($false)))
}

# Configuration du fichier, complétée par les valeurs par défaut.
function Get-SmConfig([string]$WorkDir = $SmWorkDir) {
    $config = Get-SmDefaultConfig
    $saved = Read-SmJson (Join-Path $WorkDir 'config.json')
    if ($saved) {
        foreach ($property in $saved.PSObject.Properties) {
            if ($config.Contains($property.Name)) { $config[$property.Name] = $property.Value }
        }
    }
    $config.joursEnvoi = @($config.joursEnvoi | ForEach-Object { [int]$_ })
    $config.absences = @($config.absences | Where-Object { $_ } | ForEach-Object {
        [ordered]@{ debut = [string]$_.debut; fin = [string]$_.fin; libelle = [string]$_.libelle }
    })
    $config
}

# Contrôle une configuration ; renvoie la liste des erreurs (vide si tout va bien).
function Test-SmConfig($Config) {
    $errors = @()
    if ($Config.heure -notmatch '^([01]\d|2[0-3]):[0-5]\d$') { $errors += "Heure invalide : « $($Config.heure) » (format HH:mm)." }
    $days = @($Config.joursEnvoi)
    if (-not $days.Count) { $errors += 'Choisir au moins un jour d''envoi.' }
    foreach ($day in $days) { if ([int]$day -lt 1 -or [int]$day -gt 7) { $errors += "Jour d'envoi invalide : $day." } }
    if ([int]$Config.fenetreJours -lt 1 -or [int]$Config.fenetreJours -gt 60) { $errors += 'La fenêtre doit être comprise entre 1 et 60 jours.' }
    foreach ($absence in @($Config.absences)) {
        $start = ConvertTo-SmDate $absence.debut
        $end = ConvertTo-SmDate $absence.fin
        if (-not $start -or -not $end) { $errors += "Absence « $($absence.libelle) » : dates invalides (AAAA-MM-JJ)." }
        elseif ($end -lt $start) { $errors += "Absence « $($absence.libelle) » : la fin précède le début." }
    }
    , $errors
}

function Save-SmConfig($Config, [string]$WorkDir = $SmWorkDir) {
    Write-SmJson (Join-Path $WorkDir 'config.json') $Config
}

function Get-SmEtat([string]$WorkDir = $SmWorkDir) {
    $etat = Read-SmJson (Join-Path $WorkDir 'etat.json')
    if (-not $etat) { $etat = [pscustomobject]@{ derniereSynthese = $null; dernierMode = $null; dernierFichier = $null } }
    $etat
}

function Save-SmEtat($Etat, [string]$WorkDir = $SmWorkDir) {
    Write-SmJson (Join-Path $WorkDir 'etat.json') $Etat
}

function ConvertTo-SmDate([string]$Text) {
    $date = [datetime]::MinValue
    if ([datetime]::TryParseExact($Text, 'yyyy-MM-dd', $SmInvariant, 'None', [ref]$date)) { return $date }
    $null
}

function Get-SmJourSemaine([datetime]$Date) {
    $day = [int]$Date.DayOfWeek
    if ($day -eq 0) { 7 } else { $day }
}

function Get-SmAbsence($Config, [datetime]$Date) {
    foreach ($absence in @($Config.absences)) {
        $start = ConvertTo-SmDate $absence.debut
        $end = ConvertTo-SmDate $absence.fin
        if ($start -and $end -and $Date.Date -ge $start -and $Date.Date -le $end) { return $absence }
    }
    $null
}

# Décide si une synthèse doit partir à $Moment, et sur quelle période.
function Get-SmPlan($Config, $Etat, [datetime]$Moment = (Get-Date)) {
    $plan = [ordered]@{ lancer = $false; raison = $null; depuis = $null; jours = [int]$Config.fenetreJours; absences = @() }

    if (-not $Config.actif) { $plan.raison = 'Synthèse désactivée dans la configuration.'; return $plan }
    $absence = Get-SmAbsence $Config $Moment
    if ($absence) { $plan.raison = "Absence en cours : $($absence.libelle) ($($absence.debut) → $($absence.fin))."; return $plan }
    if (@($Config.joursEnvoi) -notcontains (Get-SmJourSemaine $Moment)) {
        $plan.raison = "Pas d'envoi le $($SmJours[(Get-SmJourSemaine $Moment)])."
        return $plan
    }

    $last = $null
    if ($Etat.derniereSynthese) { $last = [datetime]::Parse([string]$Etat.derniereSynthese, $SmInvariant) }
    if ($last -and $last.Date -eq $Moment.Date) { $plan.raison = "Synthèse déjà produite aujourd'hui à $($last.ToString('HH:mm'))."; return $plan }

    if (-not $last) {
        # Première synthèse : depuis le jour d'envoi précédent, 7h00.
        $last = $Moment.Date.AddDays(-1)
        while (@($Config.joursEnvoi) -notcontains (Get-SmJourSemaine $last) -and ($Moment.Date - $last).Days -lt 7) { $last = $last.AddDays(-1) }
        $last = $last.AddHours(7)
    }
    $plan.depuis = $last
    $plan.absences = @(@($Config.absences) | Where-Object {
        $end = ConvertTo-SmDate $_.fin
        $end -and $end -ge $last.Date -and $end -lt $Moment.Date
    })
    # La fenêtre couvre au moins la période écoulée depuis la dernière synthèse, plus une semaine.
    $gap = [int][math]::Ceiling(($Moment - $last).TotalDays)
    $plan.jours = [math]::Max([int]$Config.fenetreJours, $gap + 7)
    $plan.lancer = $true
    $plan
}

# Prochaine synthèse prévue (pour la page de configuration).
function Get-SmProchaine($Config, $Etat, [datetime]$Moment = (Get-Date)) {
    $parts = $Config.heure -split ':'
    for ($i = 0; $i -le 90; $i++) {
        $candidate = $Moment.Date.AddDays($i).AddHours([int]$parts[0]).AddMinutes([int]$parts[1])
        if ($candidate -lt $Moment) { continue }
        $plan = Get-SmPlan $Config $Etat $candidate
        if ($plan.lancer) { return [ordered]@{ moment = $candidate; plan = $plan } }
        if (-not $Config.actif) { return $null }
    }
    $null
}

# --- Markdown -> HTML (le sous-ensemble produit par la synthèse) -----------

function ConvertTo-SmInlineHtml([string]$Text) {
    $html = [Net.WebUtility]::HtmlEncode($Text)
    $html = [regex]::Replace($html, '`([^`]+)`', '<code>$1</code>')
    $html = [regex]::Replace($html, '\*\*(.+?)\*\*', '<strong>$1</strong>')
    $html = [regex]::Replace($html, '\[([^\]]+)\]\((https?://[^)\s]+)\)', '<a href="$2">$1</a>')
    $html
}

function ConvertTo-SmHtml([string]$Markdown) {
    $out = New-Object Text.StringBuilder
    $lines = $Markdown -replace "`r", '' -split "`n"
    $inList = $false
    $inTable = $false
    $headerDone = $false
    foreach ($line in $lines) {
        $isRow = $line -match '^\s*\|.*\|\s*$'
        $isItem = $line -match '^\s*[-*]\s+'
        if ($inList -and -not $isItem) { [void]$out.Append('</ul>'); $inList = $false }
        if ($inTable -and -not $isRow) { [void]$out.Append('</tbody></table>'); $inTable = $false }

        if ($isRow) {
            $cells = ($line.Trim().Trim('|') -split '\|') | ForEach-Object { $_.Trim() }
            if (($cells -join '') -match '^[-:\s]+$') { continue }   # ligne de séparation
            if (-not $inTable) {
                [void]$out.Append('<table><thead><tr>')
                foreach ($cell in $cells) { [void]$out.Append("<th>$(ConvertTo-SmInlineHtml $cell)</th>") }
                [void]$out.Append('</tr></thead><tbody>')
                $inTable = $true
                continue
            }
            [void]$out.Append('<tr>')
            foreach ($cell in $cells) { [void]$out.Append("<td>$(ConvertTo-SmInlineHtml $cell)</td>") }
            [void]$out.Append('</tr>')
        } elseif ($isItem) {
            if (-not $inList) { [void]$out.Append('<ul>'); $inList = $true }
            [void]$out.Append("<li>$(ConvertTo-SmInlineHtml ($line -replace '^\s*[-*]\s+', ''))</li>")
        } elseif ($line -match '^(#{1,4})\s+(.*)$') {
            $level = $Matches[1].Length
            [void]$out.Append("<h$level>$(ConvertTo-SmInlineHtml $Matches[2])</h$level>")
        } elseif ($line -match '^>\s?(.*)$') {
            [void]$out.Append("<blockquote>$(ConvertTo-SmInlineHtml $Matches[1])</blockquote>")
        } elseif ($line.Trim()) {
            [void]$out.Append("<p>$(ConvertTo-SmInlineHtml $line)</p>")
        }
    }
    if ($inList) { [void]$out.Append('</ul>') }
    if ($inTable) { [void]$out.Append('</tbody></table>') }

    $style = 'body{font-family:Segoe UI,Arial,sans-serif;font-size:10.5pt;color:#1f2328}' +
             'h1{font-size:15pt}h2{font-size:13pt;margin-top:18px}h3{font-size:11pt}' +
             'table{border-collapse:collapse;margin:6px 0}th,td{border:1px solid #d0d7de;padding:4px 6px;vertical-align:top;text-align:left}' +
             'th{background:#f3f4f6}code{font-family:Consolas,monospace;font-size:9.5pt;background:#f3f4f6}' +
             'blockquote{color:#57606a;border-left:3px solid #d0d7de;margin:6px 0;padding-left:8px}'
    "<html><head><meta charset=`"utf-8`"><style>$style</style></head><body>$($out.ToString())</body></html>"
}
