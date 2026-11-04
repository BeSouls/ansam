<#
.SYNOPSIS
    Outils de lecture Outlook pour le chat de Panoramix : petit serveur MCP
    (JSON-RPC sur stdin / stdout), lancé par Claude Code le temps d'une réponse.

.DESCRIPTION
    Les exports du matin ne couvrent que 14 jours, avec des corps tronqués. Ces
    outils interrogent Outlook en direct, en lecture seule, comme
    Export-OutlookMails.ps1 : uniquement ma boîte, jamais les boîtes partagées.

    - chercher_mails : recherche par mots et expéditeur, sur une période
    - lire_mail      : un mail complet, historique cité compris
    - lire_agenda    : mes rendez-vous sur une période

    Aucun outil n'écrit, n'envoie ni ne déplace quoi que ce soit.
#>
[CmdletBinding()]
param([switch]$Test)

$ErrorActionPreference = 'Stop'
$utf8 = New-Object Text.UTF8Encoding($false)
$invariant = [Globalization.CultureInfo]::InvariantCulture
$limiteSecondes = 45
$corpsMax = 8000

$outils = @(
    [ordered]@{
        name = 'chercher_mails'
        description = 'Cherche dans ma boîte Outlook (tous les dossiers, Envoyés et Supprimés compris), du plus récent au plus ancien. ' +
                      'Chaque mot de « texte » doit se trouver dans l''objet, le corps, l''expéditeur ou les destinataires. Renvoie l''id à passer à lire_mail.'
        inputSchema = [ordered]@{
            type = 'object'
            properties = [ordered]@{
                texte      = [ordered]@{ type = 'string'; description = 'Mots recherchés, séparés par des espaces.' }
                expediteur = [ordered]@{ type = 'string'; description = 'Nom ou adresse de l''expéditeur.' }
                depuis     = [ordered]@{ type = 'string'; description = 'AAAA-MM-JJ. Par défaut : 180 jours en arrière.' }
                jusqua     = [ordered]@{ type = 'string'; description = 'AAAA-MM-JJ inclus. Par défaut : aujourd''hui.' }
                max        = [ordered]@{ type = 'integer'; description = 'Nombre de résultats, 15 par défaut, 40 au plus.' }
            }
        }
    }
    [ordered]@{
        name = 'lire_mail'
        description = 'Lit un mail complet à partir de l''id rendu par chercher_mails ou d''un export : corps, historique cité, destinataires, pièces jointes.'
        inputSchema = [ordered]@{
            type = 'object'
            properties = [ordered]@{ id = [ordered]@{ type = 'string'; description = 'Identifiant du mail.' } }
            required = @('id')
        }
    }
    [ordered]@{
        name = 'lire_agenda'
        description = 'Lit mes rendez-vous Outlook sur une période (séries comprises, annulés et refusés exclus).'
        inputSchema = [ordered]@{
            type = 'object'
            properties = [ordered]@{
                debut = [ordered]@{ type = 'string'; description = 'AAAA-MM-JJ' }
                fin   = [ordered]@{ type = 'string'; description = 'AAAA-MM-JJ' }
            }
            required = @('debut', 'fin')
        }
    }
)

$script:outlook = $null
$script:dossiers = $null
function Get-Mapi {
    if (-not $script:outlook) { $script:outlook = New-Object -ComObject Outlook.Application }
    $script:outlook.GetNamespace('MAPI')
}

# Dossiers de courrier de ma boîte, comme l'export, avec les Supprimés : on y cherche aussi ce qui a été jeté.
function Get-Dossiers {
    if ($script:dossiers) { return $script:dossiers }
    $mapi = Get-Mapi
    $exclus = @{}
    foreach ($id in 16, 4, 23, 20) { try { $exclus[$mapi.GetDefaultFolder($id).EntryID] = $true } catch { } }
    $liste = New-Object Collections.ArrayList
    $parcourir = {
        param($dossier)
        if ($exclus.ContainsKey($dossier.EntryID)) { return }
        if ($dossier.DefaultItemType -eq 0) { [void]$liste.Add($dossier) }
        foreach ($enfant in $dossier.Folders) { & $parcourir $enfant }
    }
    & $parcourir $mapi.DefaultStore.GetRootFolder()
    # Les dossiers les plus remplis d'abord : ce sont eux qui ont le plus de chances de répondre avant la limite.
    $script:dossiers = @($liste | Sort-Object { $_.Items.Count } -Descending)
    $script:dossiers
}

function Format-Mot([string]$Mot) { $Mot -replace "'", "''" -replace '%', '' }

function Get-Filtre([string]$Texte, [string]$Expediteur, [datetime]$Depuis, [datetime]$Jusqua) {
    $champs = 'urn:schemas:httpmail:subject', 'urn:schemas:httpmail:textdescription', 'urn:schemas:httpmail:fromname', 'urn:schemas:httpmail:fromemail', 'urn:schemas:httpmail:displayto', 'urn:schemas:httpmail:displaycc'
    $clauses = @(
        """urn:schemas:httpmail:datereceived"" >= '$($Depuis.ToUniversalTime().ToString('yyyy-MM-dd HH:mm', $invariant))'"
        """urn:schemas:httpmail:datereceived"" < '$($Jusqua.AddDays(1).ToUniversalTime().ToString('yyyy-MM-dd HH:mm', $invariant))'"
    )
    foreach ($mot in @($Texte -split '\s+' | Where-Object { $_ })) {
        $clauses += '(' + (($champs | ForEach-Object { """$_"" LIKE '%$(Format-Mot $mot)%'" }) -join ' OR ') + ')'
    }
    if ($Expediteur) {
        $clauses += '(' + (($champs[2, 3] | ForEach-Object { """$_"" LIKE '%$(Format-Mot $Expediteur)%'" }) -join ' OR ') + ')'
    }
    '@SQL=' + ($clauses -join ' AND ')
}

function Get-Apercu($Mail, [int]$Max) {
    $corps = ([string]$Mail.Body -replace '<(https?|mailto):[^>\s]*>', '' -replace "`r", '' -replace "[ `t]+`n", "`n" -replace "(`n\s*){3,}", "`n`n").Trim()
    if ($corps.Length -gt $Max) { $corps = $corps.Substring(0, $Max) + ' [...]' }
    $corps
}

function Find-Mails($Arguments) {
    $max = [math]::Min(40, [math]::Max(1, $(if ($Arguments.max) { [int]$Arguments.max } else { 15 })))
    $depuis = $(if ($Arguments.depuis) { [datetime]$Arguments.depuis } else { (Get-Date).Date.AddDays(-180) })
    $jusqua = $(if ($Arguments.jusqua) { [datetime]$Arguments.jusqua } else { (Get-Date).Date })
    if (-not $Arguments.texte -and -not $Arguments.expediteur) { throw 'Préciser « texte » ou « expediteur ».' }
    $filtre = Get-Filtre ([string]$Arguments.texte) ([string]$Arguments.expediteur) $depuis $jusqua

    $debut = Get-Date
    $trouves = New-Object Collections.ArrayList
    $dossiers = @(Get-Dossiers)
    $lus = 0
    foreach ($dossier in $dossiers) {
        if (((Get-Date) - $debut).TotalSeconds -gt $limiteSecondes) { break }
        $lus++
        try {
            $items = $dossier.Items.Restrict($filtre)
            $items.Sort('[ReceivedTime]', $true)
            $nombre = 0
            foreach ($mail in $items) {
                if ($mail.Class -ne 43) { continue }
                [void]$trouves.Add([pscustomobject]@{ Mail = $mail; Date = [datetime]$mail.ReceivedTime; Dossier = $dossier.Name })
                if (++$nombre -ge $max) { break }
            }
        } catch { }   # un dossier illisible ne doit pas faire échouer la recherche
    }
    $garde = @($trouves | Sort-Object Date -Descending | Select-Object -First $max)
    [ordered]@{
        nombre    = $garde.Count
        complet   = ($lus -eq $dossiers.Count)
        periode   = "$($depuis.ToString('yyyy-MM-dd')) au $($jusqua.ToString('yyyy-MM-dd'))"
        resultats = @($garde | ForEach-Object {
            [ordered]@{
                id = $_.Mail.EntryID; date = $_.Date.ToString('yyyy-MM-dd HH:mm'); dossier = $_.Dossier
                objet = [string]$_.Mail.Subject; de = [string]$_.Mail.SenderName; a = [string]$_.Mail.To
                apercu = Get-Apercu $_.Mail 220
            }
        })
    }
}

function Read-Mail($Arguments) {
    if (-not $Arguments.id) { throw 'Préciser « id ».' }
    $mail = (Get-Mapi).GetItemFromID([string]$Arguments.id)
    if ($mail.Class -ne 43) { throw 'Cet élément n''est pas un mail.' }
    $pieces = @(); foreach ($piece in $mail.Attachments) { $pieces += [string]$piece.FileName }
    [ordered]@{
        date = ([datetime]$mail.ReceivedTime).ToString('yyyy-MM-dd HH:mm'); dossier = [string]$mail.Parent.Name
        objet = [string]$mail.Subject; de = "$($mail.SenderName) <$($mail.SenderEmailAddress)>"
        a = [string]$mail.To; cc = [string]$mail.CC; piecesJointes = $pieces
        corps = Get-Apercu $mail $corpsMax
    }
}

function Read-Agenda($Arguments) {
    $fichier = Join-Path $env:TEMP "panoramix-agenda-$PID.json"
    try {
        & (Join-Path $PSScriptRoot 'Export-OutlookAgenda.ps1') -Debut ([datetime]$Arguments.debut) -Fin ([datetime]$Arguments.fin) -OutputFile $fichier *> $null
        [IO.File]::ReadAllText($fichier, $utf8) | ConvertFrom-Json
    } finally { Remove-Item $fichier -ErrorAction SilentlyContinue }
}

function Invoke-Outil([string]$Nom, $Arguments) {
    switch ($Nom) {
        'chercher_mails' { return Find-Mails $Arguments }
        'lire_mail'      { return Read-Mail $Arguments }
        'lire_agenda'    { return Read-Agenda $Arguments }
        default          { throw "Outil inconnu : $Nom" }
    }
}

# Essai sans passer par Claude : .\Serveur-OutilsChat.ps1 -Test
if ($Test) { return }

$entree = New-Object IO.StreamReader([Console]::OpenStandardInput(), $utf8)
$sortie = New-Object IO.StreamWriter([Console]::OpenStandardOutput(), $utf8)
$sortie.AutoFlush = $true
function Send-Reponse($Id, $Resultat) {
    $sortie.WriteLine((ConvertTo-Json -InputObject ([ordered]@{ jsonrpc = '2.0'; id = $Id; result = $Resultat }) -Depth 12 -Compress))
}

while ($null -ne ($ligne = $entree.ReadLine())) {
    if (-not $ligne.Trim()) { continue }
    $appel = $ligne | ConvertFrom-Json
    if ($null -eq $appel.id) { continue }   # notification : pas de réponse
    switch ([string]$appel.method) {
        'initialize' {
            Send-Reponse $appel.id ([ordered]@{
                protocolVersion = $(if ($appel.params.protocolVersion) { [string]$appel.params.protocolVersion } else { '2025-06-18' })
                capabilities = [ordered]@{ tools = [ordered]@{} }
                serverInfo = [ordered]@{ name = 'panoramix'; version = '1.0.0' }
            })
        }
        'tools/list' { Send-Reponse $appel.id ([ordered]@{ tools = $outils }) }
        'tools/call' {
            try {
                $texte = ConvertTo-Json -InputObject (Invoke-Outil ([string]$appel.params.name) $appel.params.arguments) -Depth 8 -Compress
                Send-Reponse $appel.id ([ordered]@{ content = @([ordered]@{ type = 'text'; text = $texte }) })
            } catch {
                Send-Reponse $appel.id ([ordered]@{ isError = $true; content = @([ordered]@{ type = 'text'; text = $_.Exception.Message }) })
            }
        }
        'ping' { Send-Reponse $appel.id ([ordered]@{}) }
        default {
            $sortie.WriteLine((ConvertTo-Json -InputObject ([ordered]@{ jsonrpc = '2.0'; id = $appel.id; error = [ordered]@{ code = -32601; message = 'Méthode inconnue' } }) -Compress))
        }
    }
}
