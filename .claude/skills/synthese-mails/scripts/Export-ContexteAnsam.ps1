<#
.SYNOPSIS
    Construit contexte-ansam.json pour la synthèse : fiches clients et
    tickets suivis localement par AnsamDiag, portefeuille TAM.

.DESCRIPTION
    Lit, sans rien modifier :
    - <DataDir>\Clients\<CODE>.json : nom, modèle, domaines, particularités ;
    - <DataDir>\Tickets\<CLE>\ticket.json : client, domaine, utilisateur,
      symptôme, statut, 3 dernières notes, clôture préparée ;
    - dans l'export des mails, les sous-dossiers Outlook de
      -DossierPortefeuille (un sous-dossier par client du portefeuille).
    Un fichier illisible est ignoré avec un avertissement.

.EXAMPLE
    .\Export-ContexteAnsam.ps1 -DataDir 'C:\…\AnsamDiag-Donnees' -ExportFile '…\mails-2026-10-28.json' -OutputFile '…\contexte-ansam.json'
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string]$DataDir,
    [Parameter(Mandatory)] [string]$OutputFile,
    [string]$ExportFile,
    [string]$DossierPortefeuille = 'Boîte de réception\MES CLIENTS',
    # Longueur maximale conservée pour les textes (notes, particularités).
    [int]$MaxTextChars = 300
)

$ErrorActionPreference = 'Stop'

function Read-Json([string]$Path) {
    try { [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8) | ConvertFrom-Json }
    catch { Write-Warning "Fichier ignoré, illisible : $Path ($($_.Exception.Message))"; $null }
}

function Limit-Text($Text) {
    $value = ([string]$Text -replace '\s+', ' ').Trim()
    if ($value.Length -gt $MaxTextChars) { $value = $value.Substring(0, $MaxTextChars) + '…' }
    $value
}

$clients = New-Object System.Collections.Generic.List[object]
$clientDir = Join-Path $DataDir 'Clients'
if (Test-Path $clientDir) {
    foreach ($file in Get-ChildItem $clientDir -Filter '*.json' -File) {
        $fiche = Read-Json $file.FullName
        if (-not $fiche) { continue }
        $domains = @(@($fiche.DomainePrincipal) + @($fiche.Domaines) + @($fiche.DomaineInitial) |
            Where-Object { $_ -and $_ -notmatch '\.onmicrosoft\.com$' } | Sort-Object -Unique)
        $clients.Add([ordered]@{
            code           = $(if ($fiche.ClientName) { $fiche.ClientName } else { $file.BaseName })
            nom            = $fiche.NomComplet
            modele         = $fiche.Modele
            actif          = $fiche.Actif
            domaines       = $domains
            particularites = @(@($fiche.Particularites) | Where-Object { $_ } | ForEach-Object {
                [ordered]@{ theme = $_.Theme; texte = (Limit-Text $_.Texte); ticket = $_.Ticket; date = $_.Date }
            })
        })
    }
} else {
    Write-Warning "Dossier des fiches clients introuvable : $clientDir"
}

$tickets = New-Object System.Collections.Generic.List[object]
$ticketDir = Join-Path $DataDir 'Tickets'
if (Test-Path $ticketDir) {
    foreach ($dir in Get-ChildItem $ticketDir -Directory) {
        $path = Join-Path $dir.FullName 'ticket.json'
        if (-not (Test-Path $path)) { continue }
        $ticket = Read-Json $path
        if (-not $ticket) { continue }
        # Les dernières notes disent où j'en suis (action faite, attente, résolu).
        $notes = @(@($ticket.Notes) | Where-Object { $_ } | Select-Object -Last 3 |
            ForEach-Object { Limit-Text "$($_.horodatage) : $($_.texte)" })
        $tickets.Add([ordered]@{
            cle         = $(if ($ticket.Ticket) { $ticket.Ticket } else { $dir.Name })
            client      = $ticket.Client
            domaine     = $ticket.Domaine
            utilisateur = $ticket.Utilisateur
            symptome    = Limit-Text $ticket.Symptome
            statut      = $ticket.Statut              # Ouvert, EnAttente, Resolu
            creeLe      = $ticket.Cree
            misAJour    = $ticket.MisAJour
            cloture     = Test-Path (Join-Path $dir.FullName 'cloture.md')
            notes       = $notes
        })
    }
} else {
    Write-Warning "Dossier des tickets introuvable : $ticketDir"
}

$portfolio = @()
if ($ExportFile -and (Test-Path $ExportFile)) {
    $export = Read-Json $ExportFile
    $prefix = $DossierPortefeuille.TrimEnd('\') + '\'
    $portfolio = @($export.folders | ForEach-Object { $_.path } |
        Where-Object { $_.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase) } |
        ForEach-Object { ($_.Substring($prefix.Length) -split '\\')[0] } | Sort-Object -Unique)
}

$context = [ordered]@{
    genereLe     = (Get-Date).ToString('yyyy-MM-ddTHH:mm:sszzz')
    source       = $DataDir
    portefeuille = $portfolio
    clients      = $clients
    tickets      = $tickets
}
New-Item -ItemType Directory -Force -Path (Split-Path $OutputFile) | Out-Null
[IO.File]::WriteAllText($OutputFile, ($context | ConvertTo-Json -Depth 6), (New-Object Text.UTF8Encoding($false)))
Write-Host "Contexte : $($clients.Count) clients, $($tickets.Count) tickets, portefeuille de $($portfolio.Count) clients -> $OutputFile"
$OutputFile
