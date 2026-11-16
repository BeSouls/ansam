<#
.SYNOPSIS
    Variabilité mensuelle lancée depuis l'onglet « Mois » de Panoramix : collecte des
    inventaires M365, puis calcul des compteurs, par les outils AnsamDiag.

.DESCRIPTION
    collecte   Pour chaque client du périmètre sans inventaire M365 du mois (ou -Client) :
               Inventaire.M365 par le relais de la console AnsamDiag du TAM. La connexion
               au tenant (code d'appareil) se fait chez le TAM, notification à l'appui ;
               un client à la fois, chacun dans son ticket local VARIA-<CODE> pour que la
               console déconnecte le tenant précédent. Lecture seule.
    calcul     Variabilite.Metriques pour le mois (ticket local VARIA-<AAAA-MM>).

    Les inventaires AD (VETBE…) restent au TAM, par MeshCentral. Rien n'est écrit
    chez le client, dans Kiosk ni dans Jira. Avancement : <WorkDir>\travaux\variabilite-<etape>.json ;
    détail : <dossierAnsamDiag>\variabilite\etapes\<etape>-<AAAA-MM>.json.

.EXAMPLE
    .\Variabilite-Etape.ps1 -Etape collecte -Mois 2026-11
    .\Variabilite-Etape.ps1 -Etape calcul -Mois 2026-11 -Client STJAC
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [ValidateSet('collecte', 'calcul')] [string]$Etape,
    [Parameter(Mandatory)] [ValidatePattern('^\d{4}-(0[1-9]|1[0-2])$')] [string]$Mois,
    [string[]]$Client,
    [int]$AttenteMinutes = 15,
    [string]$WorkDir = (Join-Path $env:USERPROFILE 'SyntheseMails')
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'SyntheseMails.Commun.ps1')
. (Join-Path $PSScriptRoot 'SyntheseMails.Mois.ps1')
$config = Get-SmConfig $WorkDir
$data = [string]$config.dossierAnsamDiag
$clientsDir = [string]$config.dossierClients
$statusFile = Join-Path $WorkDir "travaux\variabilite-$Etape.json"
function Set-Status([string]$State, [string]$Message) {
    Write-SmJson $statusFile ([ordered]@{ etape = $Etape; mois = $Mois; etat = $State; message = $Message; le = (Get-Date).ToString('o'); pid = $PID })
}
Set-Status 'en_cours' 'Démarrage'

try {
    if ($Mois -gt (Get-Date).ToString('yyyy-MM')) { throw "$Mois n'est pas encore commencé." }
    $module = Join-Path (Split-Path $data) 'scripts\AnsamDiag\AnsamDiag.psd1'
    if (-not (Test-Path -LiteralPath $module)) { throw "Module AnsamDiag introuvable : $module" }
    Import-Module $module -Force -WarningAction SilentlyContinue | Out-Null

    $codes = if ($Client) { @($Client | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim().ToUpper() } | Where-Object { $_ }) } else { @(Get-SmMoisPerimetre $config) }
    $resultats = @()

    if ($Etape -eq 'collecte') {
        if ($Mois -ne (Get-Date).ToString('yyyy-MM')) { throw "Une collecte ne se fait que pour le mois en cours ($((Get-Date).ToString('yyyy-MM'))) : un inventaire est daté du jour où il est pris." }
        $aFaire = @(foreach ($c in $codes) {
            $f = Read-SmMoisJson (Join-Path $clientsDir "$c.json")
            if (-not $f) { $resultats += [ordered]@{ client = $c; statut = 'REFUSE'; message = 'Fiche client absente' }; continue }
            if ([string]$f.Sources.Users -eq 'AD') { $resultats += [ordered]@{ client = $c; statut = 'AD'; message = 'Utilisateurs comptés dans l AD : inventaire AD par MeshCentral (TAM)' }; continue }
            $deja = @(Get-SmMoisInventaires $clientsDir $c 'M365' ([string]$f.DossierM365)) | Where-Object { $_.StartsWith($Mois) } | Select-Object -First 1
            if ($deja -and -not $Client) { $resultats += [ordered]@{ client = $c; statut = 'DEJA'; message = "Inventaire du $deja" }; continue }
            $c
        })
        $i = 0
        foreach ($c in $aFaire) {
            $i++
            $ticket = "VARIA-$c"
            Start-AnsamTicket $ticket -Client $c | Out-Null      # le ticket local porte le client : la console change de tenant
            Set-Status 'en_cours' ("{0} ({1}/{2}) : inventaire M365, connexion attendue dans la console AnsamDiag (code d'appareil)" -f $c, $i, $aFaire.Count)
            $commande = "Invoke-AnsamOutil Inventaire.M365 @{ Client = '$c'; SkipMailboxes = `$true; SkipDevices = `$true }"
            $r = Invoke-AnsamRelais -Ticket $ticket -Commande $commande -AttenteSecondes 540
            $limite = (Get-Date).AddMinutes($AttenteMinutes)
            while ($r.Statut -eq 'EN_COURS' -and (Get-Date) -lt $limite) {
                Set-Status 'en_cours' ("{0} ({1}/{2}) : {3}" -f $c, $i, $aFaire.Count, $r.Erreur)
                $r = Get-AnsamRelaisReponse -Id $r.Id -AttenteSecondes 120
            }
            $res = [ordered]@{ client = $c; statut = [string]$r.Statut; message = [string]$r.Erreur; tenant = $null; publie = $null; dossier = $null }
            if ($r.Execution -and $r.Execution.Dossier) {
                $res.dossier = [string]$r.Execution.Dossier
                $j = Read-SmMoisJson (Join-Path $res.dossier 'resultat.json')
                if ($j) {
                    $res.tenant = [ordered]@{ nom = [string]$j.tenant.nom; id = [string]$j.tenant.id; domaine = [string]$j.tenant.domaineInitial }
                    $fiche = Read-SmMoisJson (Join-Path $clientsDir "$c.json")
                    if ($fiche.TenantId -and $j.tenant.id -and [string]$fiche.TenantId -ne [string]$j.tenant.id) {
                        $res.statut = 'TENANT'; $res.message = "Tenant $($j.tenant.nom) ($($j.tenant.id)) différent de celui de la fiche ($($fiche.TenantId)) : inventaire à écarter"
                    }
                }
            }
            $dates = @(Get-SmMoisInventaires $clientsDir $c 'M365' '') | Where-Object { $_.StartsWith($Mois) }
            $res.publie = @($dates)[0]
            if ($res.statut -eq 'OK' -and -not $res.publie) { $res.statut = 'ERREUR'; $res.message = 'Inventaire terminé mais non publié dans le dossier du client' }
            $resultats += $res
        }
    } else {
        Start-AnsamTicket "VARIA-$Mois" | Out-Null
        Set-Status 'en_cours' ("Calcul de {0} pour {1}" -f $Mois, ($codes -join ', '))
        $null = Invoke-AnsamOutil Variabilite.Metriques @{ Client = ($codes -join ','); Mois = $Mois } *>&1
        foreach ($c in $codes) {
            $calc = Read-SmMoisJson (Join-Path $data "variabilite\_etat\$c\Mois-$Mois.json")
            $ok = $calc -and ([string]$calc.Calcule).StartsWith((Get-Date).ToString('yyyy-MM-dd'))
            $resultats += [ordered]@{ client = $c; statut = $(if ($ok) { 'OK' } else { 'IGNORE' }); message = $(if ($ok) { "Calculé le $($calc.Calcule)" } else { "Pas de calcul pour $Mois (inventaire du mois manquant ?)" }) }
        }
    }

    $dir = Join-Path $data 'variabilite\etapes'
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    [IO.File]::WriteAllText((Join-Path $dir "$Etape-$Mois.json"), (ConvertTo-Json -InputObject ([ordered]@{ etape = $Etape; mois = $Mois; le = (Get-Date).ToString('o'); clients = $resultats }) -Depth 6), (New-Object Text.UTF8Encoding($false)))
    $bons = @($resultats | Where-Object { $_.statut -in 'OK', 'DEJA' }).Count
    $autres = @($resultats | Where-Object { $_.statut -notin 'OK', 'DEJA' } | ForEach-Object { "$($_.client) $($_.statut)" })
    Set-Status 'ok' ("{0} client(s) prêts sur {1}{2}" -f $bons, $resultats.Count, $(if ($autres.Count) { " ; " + ($autres -join ', ') } else { '' }))
    $resultats | ForEach-Object { "{0,-6} {1,-8} {2}" -f $_.client, $_.statut, $_.message }
} catch {
    Set-Status 'echec' $_.Exception.Message
    throw
}
