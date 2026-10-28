# Synthèse mails – installation sur le PC

Tant que le connecteur Microsoft 365 est bloqué, l'export se fait en local.
Le principe :

1. Un script PowerShell demande à **Outlook classique** les mails des 14
   derniers jours de **ta boîte principale**, dans tous ses dossiers et
   sous-dossiers, puis les écrit dans un fichier JSON.
2. **Claude Code**, installé sur le PC, lit ce fichier avec le skill
   `synthese-mails` et rédige la synthèse.
3. Une **tâche planifiée Windows** enchaîne les deux chaque matin à 7h10.

Les boîtes partagées ne sont jamais lues : le script ne parcourt que le
magasin par défaut d'Outlook, et refuse explicitement une boîte partagée.
Aucune autorisation Microsoft 365 n'est nécessaire, car le script lit ce
qu'Outlook affiche déjà.

## Prérequis

- Windows avec **Outlook classique** configuré sur ta boîte ANSAM. Le
  « nouvel Outlook » n'offre pas l'accès nécessaire : il faut être sur la
  version classique, qu'on retrouve en désactivant le bouton « Nouvel
  Outlook » en haut à droite.
- **Claude Code** installé et connecté : lancer `claude` une fois dans un
  terminal et se connecter.
- Ce repo cloné sur le PC, par exemple dans `C:\Users\<toi>\ansam`.

## Étapes

Dans PowerShell, depuis le dossier du repo :

```powershell
cd .claude\skills\synthese-mails\scripts
Get-ChildItem *.ps1 | Unblock-File          # fichiers venus d'Internet

# 1. Tester l'export seul
powershell -ExecutionPolicy Bypass -File .\Export-OutlookMails.ps1
#    -> crée %USERPROFILE%\SyntheseMails\exports\mails-AAAA-MM-JJ.json

# 2. Tester la chaîne complète
powershell -ExecutionPolicy Bypass -File .\Lancer-SyntheseMails.ps1
#    -> crée %USERPROFILE%\SyntheseMails\syntheses\AAAA-MM-JJ.md

# 3. Planifier tous les jours à 7h10
powershell -ExecutionPolicy Bypass -File .\Installer-TacheSyntheseMails.ps1
```

Options utiles :

| Script | Option | Effet |
|---|---|---|
| Export / Lancer | `-Days 30` | Remonte plus loin que 14 jours |
| Export / Lancer | `-Mailbox "prenom.nom@ansam.group"` | Si la boîte ANSAM n'est pas le compte par défaut d'Outlook (nom tel qu'affiché dans le volet des dossiers) |
| Export | `-IncludeDeletedItems` | Inclut les Éléments supprimés |
| Export | `-KeepDays 7` | Durée de conservation des exports (14 jours par défaut) |
| Installer | `-At '06:45'` | Change l'heure de la tâche |

## Si ça coince

- **Message d'Outlook « un programme tente d'accéder à des adresses
  e-mail »** : Outlook juge l'antivirus non reconnu. L'informatique peut
  l'autoriser (Centre de gestion de la confidentialité → Accès par
  programme).
- **« l'exécution de scripts est désactivée »** : la stratégie du poste
  bloque PowerShell. Il faut voir avec l'informatique.
- **0 mail exporté** : vérifier que c'est bien la boîte ANSAM qui est le
  compte par défaut, sinon utiliser `-Mailbox`.
- Le journal de chaque exécution est dans `%USERPROFILE%\SyntheseMails\logs`.

## Données

Les exports contiennent le texte de tes mails : ils restent sur le PC,
dans ton profil, et sont effacés au bout de 14 jours. Pour rédiger la
synthèse, Claude Code envoie leur contenu à Claude. Le connecteur Microsoft
365 étant bloqué chez ANSAM, vérifie avec l'informatique que cet usage est
autorisé.
