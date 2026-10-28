<#
.SYNOPSIS
    Exporte en JSON les mails récents de la boîte principale Outlook :
    tous les dossiers et sous-dossiers, sans les boîtes partagées.

.DESCRIPTION
    Pilote Outlook classique (Windows) par COM : aucune autorisation
    Microsoft 365 n'est nécessaire, le script lit ce qu'Outlook voit déjà.
    Seul le magasin par défaut (ta boîte) est parcouru ; les boîtes
    partagées et déléguées sont des magasins distincts et ne sont jamais
    ouvertes. Exclus : Brouillons, Boîte d'envoi, Courrier indésirable,
    Problèmes de synchronisation, et par défaut Éléments supprimés.

    Le script écrit un fichier mails-AAAA-MM-JJ.json dans -OutputDir et
    renvoie son chemin. Les exports de plus de -KeepDays jours sont effacés.

    Ne fonctionne pas avec le « nouvel Outlook » (pas d'accès COM).

.EXAMPLE
    .\Export-OutlookMails.ps1
    .\Export-OutlookMails.ps1 -Days 30 -IncludeDeletedItems
    .\Export-OutlookMails.ps1 -Mailbox "prenom.nom@ansam.group"
#>
[CmdletBinding()]
param(
    # Fenêtre d'export, en jours.
    [int]$Days = 14,
    [string]$OutputDir = (Join-Path $env:USERPROFILE 'SyntheseMails\exports'),
    # Nom du magasin tel qu'affiché dans Outlook, si la boîte ANSAM n'est pas
    # le compte par défaut.
    [string]$Mailbox,
    # Longueur maximale du corps conservé (après retrait de l'historique cité).
    [int]$MaxBodyChars = 2000,
    # Durée de conservation des exports, qui contiennent des données sensibles.
    [int]$KeepDays = 14,
    [switch]$IncludeDeletedItems
)

$ErrorActionPreference = 'Stop'

$PR_SMTP_ADDRESS        = 'http://schemas.microsoft.com/mapi/proptag/0x39FE001E'
$PR_SENDER_SMTP_ADDRESS = 'http://schemas.microsoft.com/mapi/proptag/0x5D01001F'
$PR_LAST_VERB_EXECUTED  = 'http://schemas.microsoft.com/mapi/proptag/0x10810003'
$PR_LAST_VERB_TIME      = 'http://schemas.microsoft.com/mapi/proptag/0x10820040'
$PR_ATTACHMENT_HIDDEN   = 'http://schemas.microsoft.com/mapi/proptag/0x7FFE000B'

# OlDefaultFolders : 16 Brouillons, 4 Boîte d'envoi, 23 Courrier indésirable,
# 20 Problèmes de synchronisation, 3 Éléments supprimés.
$excludedDefaultFolders = @(16, 4, 23, 20)
if (-not $IncludeDeletedItems) { $excludedDefaultFolders += 3 }

$lastVerbs = @{ 102 = 'reply'; 103 = 'replyAll'; 104 = 'forward' }
$invariant = [Globalization.CultureInfo]::InvariantCulture

function Format-Date($value) {
    if ($null -eq $value) { return $null }
    $date = [datetime]$value
    if ($date.Year -ge 4500) { return $null }   # 01.01.4501 = « aucune date » pour Outlook
    $date.ToString('yyyy-MM-ddTHH:mm:sszzz', $invariant)
}

function Get-SenderAddress($item) {
    try {
        if ($item.SenderEmailType -eq 'EX') {
            return $item.PropertyAccessor.GetProperty($PR_SENDER_SMTP_ADDRESS)
        }
    } catch { }
    $item.SenderEmailAddress
}

function Get-Recipients($item, [int]$type) {
    $list = @()
    foreach ($recipient in $item.Recipients) {
        if ($recipient.Type -ne $type) { continue }
        $address = $recipient.Address
        try { $address = $recipient.PropertyAccessor.GetProperty($PR_SMTP_ADDRESS) } catch { }
        $list += [ordered]@{ name = $recipient.Name; address = $address }
    }
    , $list
}

# Corps sans l'historique cité : les messages précédents du fil sont exportés
# séparément, inutile de les répéter.
function Get-CleanBody($item) {
    $body = [string]$item.Body
    $quote = [regex]::Match($body, '(?m)^\s*(_{20,}|-{3,}\s*(Original Message|Message d.origine)|(De|From)\s*:\s)')
    if ($quote.Success -and $quote.Index -gt 0) { $body = $body.Substring(0, $quote.Index) }
    $body = ($body -replace "`r", '' -replace "(`n\s*){3,}", "`n`n").Trim()
    if ($body.Length -gt $MaxBodyChars) { $body = $body.Substring(0, $MaxBodyChars) + "`n[... tronqué]" }
    $body
}

function Convert-Mail($item, [string]$folderPath) {
    $accessor = $item.PropertyAccessor
    $lastAction = $null
    try {
        $verb = $accessor.GetProperty($PR_LAST_VERB_EXECUTED)
        if ($lastVerbs.ContainsKey([int]$verb)) {
            $lastAction = [ordered]@{
                type = $lastVerbs[[int]$verb]
                at   = Format-Date ($accessor.UTCToLocalTime($accessor.GetProperty($PR_LAST_VERB_TIME)))
            }
        }
    } catch { }

    $flag = $null
    if ($item.FlagStatus -ne 0) {
        $flag = [ordered]@{
            status  = $(if ($item.FlagStatus -eq 1) { 'complete' } else { 'flagged' })
            request = $item.FlagRequest
            due     = Format-Date $item.TaskDueDate
        }
    }

    $attachments = @()
    foreach ($attachment in $item.Attachments) {
        $hidden = $false
        try { $hidden = [bool]$attachment.PropertyAccessor.GetProperty($PR_ATTACHMENT_HIDDEN) } catch { }
        if (-not $hidden) { $attachments += $attachment.FileName }
    }

    [ordered]@{
        id                = $item.EntryID
        conversationId    = $item.ConversationID
        conversationTopic = $item.ConversationTopic
        folder            = $folderPath
        subject           = $item.Subject
        from              = [ordered]@{ name = $item.SenderName; address = (Get-SenderAddress $item) }
        to                = (Get-Recipients $item 1)
        cc                = (Get-Recipients $item 2)
        received          = Format-Date $item.ReceivedTime
        sent              = Format-Date $item.SentOn
        unread            = [bool]$item.UnRead
        importance        = @('low', 'normal', 'high')[$item.Importance]
        categories        = $item.Categories
        flag              = $flag
        lastAction        = $lastAction
        attachments       = $attachments
        body              = Get-CleanBody $item
    }
}

function Export-Folder($folder) {
    if ($excludedIds.ContainsKey($folder.EntryID)) { return }
    $path = $folder.FolderPath -replace '^\\\\[^\\]+\\?', ''
    if (-not $path) { $path = '(racine)' }

    if ($folder.DefaultItemType -eq 0) {   # dossiers de courrier uniquement
        $count = 0
        foreach ($item in $folder.Items.Restrict($filter)) {
            if ($item.Class -ne 43) { continue }   # MailItem
            if ($item.MessageClass -like 'IPM.Note.Microsoft.Conversation*') { continue }   # historique Teams/Skype
            try {
                $messages.Add((Convert-Mail $item $path))
                $count++
            } catch {
                Write-Warning "Mail ignoré dans « $path » : $($_.Exception.Message)"
            }
        }
        $folderStats.Add([ordered]@{ path = $path; exported = $count })
    }

    foreach ($subFolder in $folder.Folders) { Export-Folder $subFolder }
}

# --- Boîte principale -------------------------------------------------------

$outlook = New-Object -ComObject Outlook.Application   # se rattache à Outlook s'il est ouvert
$namespace = $outlook.GetNamespace('MAPI')

if ($Mailbox) {
    $store = $namespace.Stores | Where-Object { $_.DisplayName -eq $Mailbox } | Select-Object -First 1
    if (-not $store) { throw "Boîte « $Mailbox » introuvable dans Outlook." }
} else {
    $store = $namespace.DefaultStore
}
# ExchangeStoreType 1 = boîte d'un autre utilisateur (partagée ou déléguée).
if ($store.ExchangeStoreType -eq 1) { throw "« $($store.DisplayName) » est une boîte partagée : export refusé." }

$excludedIds = @{}
foreach ($id in $excludedDefaultFolders) {
    try { $excludedIds[$store.GetDefaultFolder($id).EntryID] = $true } catch { }
}

# Filtre DASL : dates en UTC, format indépendant de la langue de Windows.
$since = (Get-Date).Date.AddDays(-$Days)
$sinceUtc = $since.ToUniversalTime().ToString('yyyy-MM-dd HH:mm', $invariant)
$filter = "@SQL=(""urn:schemas:httpmail:datereceived"" >= '$sinceUtc' OR ""urn:schemas:httpmail:date"" >= '$sinceUtc')"

$messages = New-Object System.Collections.Generic.List[object]
$folderStats = New-Object System.Collections.Generic.List[object]
Export-Folder $store.GetRootFolder()

# --- Écriture ---------------------------------------------------------------

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
$file = Join-Path $OutputDir ('mails-{0:yyyy-MM-dd}.json' -f (Get-Date))
$export = [ordered]@{
    exportedAt   = Format-Date (Get-Date)
    mailbox      = $store.DisplayName
    since        = Format-Date $since
    folders      = $folderStats
    messageCount = $messages.Count
    messages     = $messages
}
$json = $export | ConvertTo-Json -Depth 6
[IO.File]::WriteAllText($file, $json, (New-Object Text.UTF8Encoding($false)))

Get-ChildItem -Path $OutputDir -Filter 'mails-*.json' |
    Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$KeepDays) } |
    Remove-Item -Force

Write-Host "$($messages.Count) mails exportés depuis $($folderStats.Count) dossiers de « $($store.DisplayName) » -> $file"
$file
