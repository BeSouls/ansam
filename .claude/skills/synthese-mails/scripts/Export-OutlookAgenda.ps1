<#
.SYNOPSIS
    Exporte en JSON les rendez-vous de mon calendrier Outlook sur une période.

.DESCRIPTION
    Pilote Outlook classique par COM, en lecture seule, comme Export-OutlookMails.ps1 :
    seul le calendrier par défaut de ma boîte est lu, jamais celui d'un autre.
    Pour chaque rendez-vous (occurrences des séries comprises) : objet, début, fin,
    lieu, organisateur, participants (nom, adresse), catégories, disponibilité,
    réponse, et le début de la description (sans le bloc de connexion Teams).
    Ignorés : rendez-vous annulés ou refusés. Un rendez-vous privé ne donne que
    ses horaires (« Privé »).

    Sert à préparer la réunion squad (RDV clients, visites, revues). Renvoie le
    chemin du JSON écrit.

.EXAMPLE
    .\Export-OutlookAgenda.ps1 -Debut 2026-10-26 -Fin 2026-11-30 -OutputFile "$env:USERPROFILE\SyntheseMails\exports\agenda-2026-11-02.json"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [datetime]$Debut,
    [Parameter(Mandatory)] [datetime]$Fin,
    [Parameter(Mandatory)] [string]$OutputFile,
    [int]$MaxBodyChars = 400
)

$ErrorActionPreference = 'Stop'
$PR_SMTP_ADDRESS = 'http://schemas.microsoft.com/mapi/proptag/0x39FE001E'
$invariant = [Globalization.CultureInfo]::InvariantCulture
# OlBusyStatus et OlResponseStatus.
$busy = @{ 0 = 'libre'; 1 = 'provisoire'; 2 = 'occupe'; 3 = 'absent'; 4 = 'ailleurs' }
$responses = @{ 0 = 'aucune'; 1 = 'organisateur'; 2 = 'provisoire'; 3 = 'accepte'; 4 = 'refuse'; 5 = 'sans_reponse' }

function Format-Date($value) { ([datetime]$value).ToString('yyyy-MM-ddTHH:mm:sszzz', $invariant) }

function Get-Attendees($item) {
    $list = @()
    foreach ($recipient in $item.Recipients) {
        $address = $recipient.Address
        try { $address = $recipient.PropertyAccessor.GetProperty($PR_SMTP_ADDRESS) } catch { }
        $list += [ordered]@{ name = $recipient.Name; address = $address; type = $(if ($recipient.Type -eq 2) { 'facultatif' } elseif ($recipient.Type -eq 3) { 'ressource' } else { 'requis' }) }
    }
    , $list
}

function Get-ShortBody($item) {
    $body = [string]$item.Body
    # Bloc de connexion Teams / Zoom et liens : inutiles ici.
    $cut = [regex]::Match($body, '(?m)^\s*(_{10,}|Réunion Microsoft Teams|Microsoft Teams meeting|Rejoindre|Join)')
    if ($cut.Success) { $body = $body.Substring(0, $cut.Index) }
    $body = ($body -replace '<(https?|mailto):[^>\s]*>', '' -replace "`r", '' -replace "(`n\s*){2,}", "`n").Trim()
    if ($body.Length -gt $MaxBodyChars) { $body = $body.Substring(0, $MaxBodyChars) + ' [...]' }
    $body
}

$outlook = New-Object -ComObject Outlook.Application
$calendar = $outlook.GetNamespace('MAPI').GetDefaultFolder(9)   # olFolderCalendar
$items = $calendar.Items
# Ordre imposé par Outlook pour obtenir les occurrences des séries : tri, puis récurrences, puis filtre.
$items.Sort('[Start]')
$items.IncludeRecurrences = $true
$filter = "[Start] < '$($Fin.Date.AddDays(1).ToString('g'))' AND [End] > '$($Debut.Date.ToString('g'))'"
$events = @()
foreach ($item in $items.Restrict($filter)) {
    if ($item.Class -ne 26) { continue }   # olAppointment
    if ($item.MeetingStatus -eq 5 -or $item.MeetingStatus -eq 7 -or [string]$item.Subject -match '^(Annulé|Canceled|Cancelled)\s*:') { continue }   # annulé
    if ($item.ResponseStatus -eq 4) { continue }                              # refusé
    $start = [datetime]$item.Start
    if ($start -ge $Fin.Date.AddDays(1)) { break }
    $private = $item.Sensitivity -eq 2
    $events += [ordered]@{
        objet        = $(if ($private) { 'Privé' } else { [string]$item.Subject })
        debut        = Format-Date $item.Start
        fin          = Format-Date $item.End
        journee      = [bool]$item.AllDayEvent
        lieu         = $(if ($private) { $null } else { [string]$item.Location })
        organisateur = $(if ($private) { $null } else { [string]$item.Organizer })
        participants = $(if ($private) { @() } else { Get-Attendees $item })
        categories   = [string]$item.Categories
        disponibilite = $busy[[int]$item.BusyStatus]
        reponse      = $responses[[int]$item.ResponseStatus]
        enLigne      = [bool]($item.IsOnlineMeeting -or ([string]$item.Location -match 'Teams|Zoom|Webex'))
        description  = $(if ($private) { $null } else { Get-ShortBody $item })
    }
}

$result = [ordered]@{
    exporteLe = (Get-Date).ToString('o')
    debut     = $Debut.ToString('yyyy-MM-dd')
    fin       = $Fin.ToString('yyyy-MM-dd')
    nombre    = $events.Count
    evenements = $events
}
New-Item -ItemType Directory -Force -Path (Split-Path $OutputFile) | Out-Null
[IO.File]::WriteAllText($OutputFile, ($result | ConvertTo-Json -Depth 5), (New-Object Text.UTF8Encoding($false)))
Write-Host "Agenda : $($events.Count) rendez-vous du $($Debut.ToString('dd.MM')) au $($Fin.ToString('dd.MM'))."
$OutputFile
