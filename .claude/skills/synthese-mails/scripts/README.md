# Synthèse mails – installation sur le PC

> Reprendre le sujet dans une nouvelle session : lire `AnsamDiag-Donnees/PANORAMIX-REPRISE.md` (Documents du OneDrive ANSAM).

Tant que le connecteur Microsoft 365 est bloqué, l'export se fait en local.
Le principe :

1. Un script PowerShell demande à **Outlook classique** les mails de **ta
   boîte principale** (14 jours par défaut), dans tous ses dossiers et
   sous-dossiers, et les écrit dans un JSON et dans une vue compacte par fil.
2. Un second script rassemble le **contexte AnsamDiag** : fiches clients
   (`Clients\*.json`) et tickets suivis (`Tickets\*\ticket.json`), pour
   rattacher chaque mail à son client et à son ticket.
3. **Claude Code** lit tes **tickets Jira assignés ouverts** avec le
   connecteur Atlassian, en lecture seule (outils d'écriture refusés), et
   vérifie que la liste est complète (total annoncé par Jira). Le lanceur
   donne ensuite **un point de suivi à chaque ticket ouvert**, même sans
   mail ni dossier AnsamDiag (statut et priorité repris de Jira, client
   déduit du dossier AnsamDiag, du rapporteur ou du projet) ; un ticket
   encore ouvert dans Jira n'est jamais considéré comme réglé en silence.
4. **Claude Code**, installé sur le PC, lit ces fichiers et le **registre
   de suivi** (`suivi.json`) avec le skill `synthese-mails`, rédige la
   synthèse et met à jour le registre : chaque point reste ouvert d'un jour
   à l'autre tant qu'il n'est pas réglé (marquage, AnsamDiag, Jira ou mail).
5. La synthèse t'est **envoyée par mail** (à ta propre adresse, via
   Outlook) et gardée dans `%USERPROFILE%\SyntheseMails\syntheses`.
6. Une **tâche planifiée Windows** enchaîne le tout chaque matin. Les jours
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

La page locale de suivi et de configuration tourne en permanence une fois
les tâches installées : elle démarre à l'ouverture de session, sur
`http://127.0.0.1:8765/` avec un jeton propre au PC (un lien sans jeton est
refusé). Pour l'ouvrir : le raccourci **« Panoramix »** du menu
Démarrer, le lien « Ouvrir Panoramix » en bas de chaque synthèse, ou
**`Configurer-SyntheseMails.cmd`** (qui démarre une page temporaire si le
service ne tourne pas).

Onglet **Suivi** : les points du registre, ouverts ou réglés. Pour chacun,
choisir « À faire », « En attente d'un tiers » (ton action est faite),
« Traité » ou « Sans objet », avec une note (« réglé par Teams le 28.10 »).
Ton marquage tient jusqu'à ce qu'un nouveau message arrive sur le point.

**Drapeau de priorité** : le drapeau en haut à droite d'une carte (tâche ou
groupe) la passe dans la section « Prioritaires », en tête de page, bordée
de rouge. Un groupe y monte aussi quand l'une de ses sous-tâches ouvertes
porte le drapeau. Seul toi le poses ou le retires : la synthèse n'y touche
pas. Une tâche traitée retourne dans sa section, drapeau ou non.

**Groupes** : les tâches d'une même activité se rangent sous une tâche
plus large (par exemple « Revues des tickets », une sous-tâche par client).
La synthèse regroupe d'elle-même les revues récurrentes et les chantiers ;
à la main, cocher les tâches (rond à gauche) puis « Grouper » ou « Ajouter
à un groupe » dans la barre qui apparaît en bas. Un clic sur un groupe
ouvre sa fenêtre : statut des sous-tâches, « Sortir du groupe », nom
modifiable, « Tout marquer traité », « Dissoudre le groupe ». Un point
rangé ou sorti à la main n'est plus déplacé par la synthèse.

Filtre **À contacter** : les tâches qui attendent un appel ou un mail de
ta part. Une tâche dont l'action est d'écrire ou d'appeler à quelqu'un
(« transmettre à Swisscom », « relancer », « répondre à »…) y figure même
si la synthèse a oublié la prise de contact : elle reçoit alors un brouillon
vide, à rédiger avec tout le dossier. « Préparer le mail » ouvre le brouillon (destinataire, objet,
texte) à relire, puis « Créer le brouillon Outlook » (rien n'est envoyé),
« Ouvrir dans ma messagerie » ou « Copier ». « Préparer l'appel » donne le
numéro et les points à aborder, et garde le compte rendu dans la note du
point.

Dans ces deux fenêtres, **« Rédiger avec tout le dossier »** (1 à 2 min)
fait relire à Claude Code, pour ce seul point : toutes les notes et les
résultats de tests du dossier AnsamDiag, `TICKET.md`, les messages déjà
rédigés, les commentaires Jira (lecture seule) et le fil de mails complet.
Le brouillon ou les points d'appel sont réécrits sans reposer une question
déjà traitée, et un résumé « Où on en est » s'affiche. Cette version est
gardée tant qu'aucun nouveau message n'arrive sur le point. Le champ
**« À savoir en plus et consignes pour la rédaction »**, juste au-dessus du bouton, sert
à donner ce qui n'est pas encore écrit (un appel, un statut, une décision) et des consignes (simplifier, ajouter un élément), toutes appliquées :
il est gardé sur le point, et prime sur le dossier en cas de contradiction.

Bouton **« Claude »** sur chaque tâche : ouvre une session Claude Code
dans le dossier AnsamDiag (ses règles et ses skills), dans un onglet
Windows Terminal : « Analyser le ticket » (`/ticket`), « Timesheet »,
« Clôture », « Escalade », ou « Session libre » avec le contexte du point.
Chaque session est mémorisée sur la tâche (« Claude ● ») et se reprend
depuis le même menu (« Reprendre »), là où elle s'était arrêtée.

Vues **Variabilité**, **Facturation** et **Attention** : les éléments à
garder relevés par les synthèses (`archives.json`), établis ou à vérifier
(avec l'indice et la façon de vérifier). La variabilité se lit par mois et
par client ; on marque chaque élément reporté, traité, vu ou écarté, d'un
clic, avec possibilité d'annuler. Les informations TAM à archiver sont dans
la vue Attention.

Bouton **« Actualiser »** (en haut de la page, un clic) : relance toute
l'analyse (mails, AnsamDiag, Jira) et met à jour les points et les
éléments à garder, **sans envoyer de mail**. Utile quand des mails ont été
envoyés ou reçus depuis la synthèse du matin. La synthèse du jour et sa
date ne changent pas : celle du lendemain couvre toujours tout depuis le
dernier envoi. Le texte produit est gardé dans `syntheses\actualisations`.
Comme l'analyse du jour a déjà tout traité, l'actualisation ne fait relire
à Claude que les fils qui ont un message depuis la dernière analyse, en
entier et avec les fils du même ticket (section « Actualisation » du
skill) et les points à reprendre (voir ci-dessous) ; sans nouveau message
ni activité sur les tickets, Claude n'est pas lancé. Jira est lu en
parallèle de l'export Outlook. Compter 2 à 5 minutes avec Outlook ouvert
(un Outlook qui démarre impose 3 minutes d'attente de synchronisation).

**Consommation** : la liste des tickets Jira et leurs commentaires sont de la
recopie, confiée à un modèle léger (si un champ manque, le modèle standard
refait la lecture). L'analyse reçoit `suivi-analyse.json` et
`archives-analyse.json` : points ouverts, 3 derniers événements, points
réglés en résumé. `suivi.json` et `archives.json` restent complets.

**Annuaire** : l'onglet Annuaire liste les personnes rencontrées dans les tickets
et les mails (`annuaire.json`). Il grandit à chaque analyse et à chaque rédaction,
et complète les prises de contact à qui il manque une adresse ou un numéro. Une
fiche modifiée dans la page n'est plus touchée par les analyses, une fiche retirée
ne revient pas. Les Contacts des tiers des Réglages restent le canal officiel d'un
fournisseur.

**Chat** : l'avatar en bas à droite de la page ouvre un petit panneau pour poser
une question, vérifier une information ou faire rédiger un texte.
`Repondre-Chat.ps1` répond avec Claude Code, en Sonnet par défaut (`-Modele`).
Il lit les registres, les dossiers AnsamDiag et clients, Jira et Confluence, et
Outlook en direct par `Serveur-OutilsChat.ps1` (mails sur 180 jours par défaut,
agenda). Il ne peut ni écrire, ni lancer de commande, ni aller sur le web. Un
mail préparé devient un brouillon Outlook seulement si on clique sur le bouton.
Une tâche ouverte dans la page est transmise à Claude. Les conversations sont
dans `chat\` et s'effacent après 14 jours. Consignes :
`references/consigne-chat.md`.

**Actualisations automatiques** : aux heures du réglage du même nom (12:30 et
16:30 par défaut, vide pour aucune), la tâche « Synthese mails ANSAM -
actualisation » fait comme le bouton, sans mail. Elle s'abstient pendant une
absence, un jour férié, un jour sans envoi ou avant la synthèse du matin, et ne
réveille pas le PC. Réinstaller les tâches après avoir changé le réglage depuis
un fichier ; depuis la page, c'est automatique.

**Statuts fiables** : à chaque analyse, `Export-JiraActivite.ps1` relit les
commentaires des tickets dont la date a changé (`jira-activite.json`). Jira
ne m'envoie aucun mail pour mes commentaires, et ceux des autres arrivent
sans leur texte : c'est la seule façon de voir ce qui s'est passé dans un
ticket. Le lanceur fait ensuite, sans passer par l'IA :

- un ticket clos ferme son point, même marqué par moi ; un changement de
  statut Jira est repris, sauf sur un point marqué par moi ;
- tout point dont une source a bougé depuis sa dernière revue (commentaire,
  statut Jira, note AnsamDiag, mail) est listé dans `a-revoir.json` et doit
  être repris par l'analyse, quel que soit le mode.

Sur la page, chaque tâche ouverte montre son dernier commentaire Jira et la
date de sa dernière revue. Le badge **« à revérifier »** signale une tâche
dont une source a bougé sans que l'analyse l'ait reprise, et l'en-tête
signale une source non lue : un statut sans badge a été revu avec tout ce
qui était connu à cette date. Les tâches sans ticket ne sont revues que
d'après les mails.

**Synthèse incrémentale** : de même, la synthèse du matin ne relit en
entier que les fils qui ont du nouveau depuis la dernière analyse ; les
points ouverts sont revus depuis le registre et Jira, sans relire leurs
anciens mails, et la synthèse reste complète pour le lecteur. Une **analyse
complète** de toute la fenêtre a lieu le lundi (réglage
`analyseCompleteJour` de `config.json`), au retour d'absence et après plus
de 3 jours sans analyse ; elle rattrape un fil qui aurait été mal classé.
`Lancer-SyntheseMails.ps1 -Complet` la force, `-Essai` produit une
synthèse dans `syntheses\essais` sans l'envoyer.

Bouton **« Envoyer la synthèse »** (en haut de la page, deux clics pour
confirmer) : relance tout de suite la synthèse complète avec les statuts
à jour et l'envoie par mail, par exemple après avoir marqué les points
déjà réglés. La progression s'affiche sous l'en-tête. Un seul lancement à
la fois : si une synthèse tourne déjà (tâche du matin ou bouton), le
second lancement est refusé.

Onglet **Squad** : le formulaire de la réunion hebdomadaire de la squad
(participants, Clients A+, Mercato, RDV clients de la semaine écoulée et de
la semaine en cours, visites, revues, projets, problématiques, variabilités,
facturation, satisfactions, divers). Le jour de la réunion (lundi par
défaut), la synthèse du matin prépare **mes lignes** (`Preparer-Squad.ps1`,
règles dans `references/squad.md`) à partir de l'agenda Outlook
(`Export-OutlookAgenda.ps1`, de la semaine écoulée aux quatre semaines
suivantes), des mails, du suivi, des éléments à garder, de Jira et de la
fiche précédente ; chaque ligne porte sa source. Le mail du matin contient
le lien « Ouvrir la préparation ». Pas de préparation un jour d'absence ou
férié. Dans l'onglet, tout se modifie sur place et s'enregistre seul
(`%USERPROFILE%\SyntheseMails\squad\AAAA-MM-JJ.json`) ; « Mettre à jour »
relance la préparation sans toucher à ce que j'ai écrit ou modifié.
« Copier pour OneNote » copie tout le formulaire en tableaux, « Copier mes
lignes » seulement les miennes, par sujet.

Onglet **Réglages** :

- synthèse active, envoi par mail, heure, jours d'envoi ;
- **garder Outlook classique ouvert en arrière-plan** (activé par défaut) :
  la page permanente tient Outlook classique ouvert sans fenêtre ni icône
  (`OUTLOOK.EXE -Embedding`, environ 250 Mo), et le relance s'il s'arrête.
  On peut utiliser le nouvel Outlook au quotidien : les exports n'attendent
  plus les 3 minutes de démarrage d'Outlook classique ;
- **jours fériés** du canton (Vaud par défaut, Genève, ou aucun) : ni
  synthèse ni préparation de réunion ce jour-là ;
- **réunion squad** : jour, mes initiales et mon prénom, participants,
  initiales de l'équipe (rubriques RDV clients et visites) ;
- fenêtre relue (14 jours par défaut) ;
- **absences** (vacances, formation…) : pas de synthèse pendant
  l'absence, puis une synthèse de retour qui couvre toute la période ;
- **contacts des tiers** : canal officiel d'un opérateur ou d'un
  fournisseur (par défaut Swisscom Partner Management,
  partner.management@swisscom.com, +41 800 800 650). Une tâche qui demande
  de contacter ce tiers (reconnu par ses mots-clés) reçoit cette adresse et
  ce numéro ; un cas ServiceNow (CS…) reste suivi dans ServiceNow ;
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
- **Pas de mail du tout** : regarder le résultat de la tâche dans le
  Planificateur. En veille moderne, Windows ne lance aucune tâche : avec
  « Réveiller le PC à l'heure prévue » (Réglages, activé par défaut), la
  tâche réveille le PC les jours d'envoi, si le plan d'alimentation
  l'autorise (en général sur secteur seulement). Sinon (batterie, PC
  éteint), la synthèse part à l'ouverture de session ou au déverrouillage,
  environ 8 minutes plus tard. `0xC000013A` ou `0x8007042B` : la console
  du script a été tuée au démarrage ; les tâches passent par
  `conhost.exe --headless` pour l'éviter (réinstaller les tâches depuis la
  page si elles datent d'avant).
- **« Accès refusé : lien incomplet »** sur la page : le lien n'a pas le
  jeton ; utiliser le raccourci du menu Démarrer.
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
