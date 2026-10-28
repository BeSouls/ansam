# Synthèse mails – installation sur le PC

Tant que le connecteur Microsoft 365 est bloqué, l'export se fait en local.
Le principe :

1. Un script PowerShell demande à **Outlook classique** les mails de **ta
   boîte principale** (14 jours par défaut), dans tous ses dossiers et
   sous-dossiers, et les écrit dans un JSON et dans une vue compacte par fil.
2. Un second script rassemble le **contexte AnsamDiag** : fiches clients
   (`Clients\*.json`) et tickets suivis (`Tickets\*\ticket.json`), pour
   rattacher chaque mail à son client et à son ticket.
3. **Claude Code**, installé sur le PC, lit ces fichiers avec le skill
   `synthese-mails` et rédige la synthèse.
4. La synthèse t'est **envoyée par mail** (à ta propre adresse, via
   Outlook) et gardée dans `%USERPROFILE%\SyntheseMails\syntheses`.
5. Une **tâche planifiée Windows** enchaîne le tout chaque matin. Les jours
   d'envoi et les absences sont respectés ; au retour d'une absence, une
   seule synthèse couvre toute la période.

Les boîtes partagées ne sont jamais lues : le script ne parcourt que le
magasin par défaut d'Outlook, et refuse explicitement une boîte partagée.
Aucune autorisation Microsoft 365 n'est nécessaire, car le script lit ce
qu'Outlook affiche déjà.

## Prérequis

- Windows avec **Outlook classique** configuré sur ta boîte ANSAM. Le
  « nouvel Outlook » n'offre pas l'accès nécessaire : la version classique
  doit être installée (elle peut tourner à côté du nouvel Outlook).
- **Claude Code** installé et connecté : lancer `claude` une fois dans un
  terminal et se connecter. Si l'installeur signale que
  `%USERPROFILE%\.local\bin` n'est pas dans le PATH, le lanceur le trouve
  quand même.
- Ce repo cloné sur le PC, par exemple dans `C:\Users\<toi>\dev\ansam`.

## Configuration

Double-clic sur **`Configurer-SyntheseMails.cmd`** : une page locale
s'ouvre dans le navigateur (servie sur 127.0.0.1 uniquement, fermée après
30 minutes sans activité). On y règle :

- synthèse active, envoi par mail, heure, jours d'envoi ;
- fenêtre relue (14 jours par défaut) ;
- **absences** (vacances, formation…) : pas de synthèse pendant
  l'absence, puis une synthèse de retour qui couvre toute la période ;
- dossier des données AnsamDiag et dossier Outlook du portefeuille.

La page affiche aussi la prochaine synthèse prévue, l'état de la tâche
planifiée et les dernières synthèses, et installe ou met à jour la tâche.
Les réglages sont dans `%USERPROFILE%\SyntheseMails\config.json`.

## Étapes

Dans PowerShell, depuis le dossier du repo :

```powershell
cd .claude\skills\synthese-mails\scripts
Get-ChildItem *.ps1 | Unblock-File          # fichiers venus d'Internet

# 1. Tester l'export seul
powershell -ExecutionPolicy Bypass -File .\Export-OutlookMails.ps1
#    -> %USERPROFILE%\SyntheseMails\exports\mails-AAAA-MM-JJ.json et -fils.txt

# 2. Tester la chaîne complète, mail compris, même si une synthèse est déjà partie aujourd'hui
powershell -ExecutionPolicy Bypass -File .\Lancer-SyntheseMails.ps1 -Force

# 3. Régler et planifier
.\Configurer-SyntheseMails.cmd
```

Options utiles :

| Script | Option | Effet |
|---|---|---|
| Lancer | `-Force` | Lance même un jour non prévu, pendant une absence ou déjà traité |
| Lancer | `-PasDeMail` | Enregistre la synthèse sans l'envoyer |
| Export / Lancer | `-Mailbox "prenom.nom@ansam.group"` | Si la boîte ANSAM n'est pas le compte par défaut d'Outlook (nom tel qu'affiché dans le volet des dossiers) |
| Export | `-Days 30` | Remonte plus loin (la fenêtre du lanceur se règle dans la page) |
| Export | `-IncludeDeletedItems` | Inclut les Éléments supprimés |
| Export | `-KeepDays 7` | Durée de conservation des exports (14 jours par défaut) |
| Export | `-SyncTimeoutSeconds 0` | Pas d'attente de synchronisation (par défaut : jusqu'à 10 min, au moins 3 min après le démarrage d'Outlook) |
| Installer | `-At '06:45'` | Heure de la tâche (par défaut : celle de la configuration) |

## Si ça coince

- **Mail « Synthèse mails – échec du … »** : la chaîne a échoué ; le mail
  donne l'erreur et le chemin du journal.
- **`0x80080005` (CO_E_SERVER_EXEC_FAILURE)** : l'Outlook ouvert ne répond
  pas aux scripts. Fermer complètement Outlook (aucun `OUTLOOK.EXE` restant
  dans le Gestionnaire des tâches), le rouvrir, relancer.
- **Dossiers ou Éléments envoyés manquants** : le cache d'Outlook classique
  n'était pas à jour. L'export attend la synchronisation ; `sync.status`
  vaut `timeout` dans le JSON si elle n'a pas été confirmée.
- **« Not logged in · Please run /login »** dans le journal : lancer
  `claude` une fois et se reconnecter.
- **Message d'Outlook « un programme tente d'accéder à des adresses
  e-mail » ou « … d'envoyer un message »** : Outlook juge l'antivirus non
  reconnu. L'informatique peut l'autoriser (Centre de gestion de la
  confidentialité → Accès par programme).
- **« l'exécution de scripts est désactivée »** : la stratégie du poste
  bloque PowerShell. Il faut voir avec l'informatique.
- **0 mail exporté** : vérifier que c'est bien la boîte ANSAM qui est le
  compte par défaut, sinon utiliser `-Mailbox`.
- Le journal de chaque exécution est dans `%USERPROFILE%\SyntheseMails\logs`.

## Données

Les exports contiennent le texte de tes mails : ils restent sur le PC,
dans ton profil, et sont effacés au bout de 14 jours. Pour rédiger la
synthèse, Claude Code envoie leur contenu, ainsi que le contexte AnsamDiag,
à Claude. Le connecteur Microsoft 365 étant bloqué chez ANSAM, vérifie avec
l'informatique que cet usage est autorisé.
