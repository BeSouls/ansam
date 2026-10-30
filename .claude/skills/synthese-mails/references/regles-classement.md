# Règles de classement

## Établi ou supposé

Chaque élément de la synthèse est soit **établi** (écrit noir sur blanc dans
un mail ou dans le contexte AnsamDiag), soit **supposé** (déduit d'un
indice). Un élément supposé n'est **jamais écarté** et **jamais présenté
comme un fait** : il apparaît dans la synthèse, marqué « à vérifier », avec
l'indice qui l'a fait remonter et la façon de le vérifier. Mieux vaut une
suggestion de trop qu'une information manquée.

## Rattacher un mail à son client et à son ticket

Avant de classer, croiser chaque fil avec `contexte-ansam.json` (voir
SKILL.md) :

1. **Client** : domaine de l'expéditeur ou des destinataires contre les
   `domaines` des fiches clients, puis nom de société (`nom`), puis dossier
   Outlook (`…\MES CLIENTS\<CODE>`, `…\CLIENTS\<CODE>`). Sans
   correspondance : « client à identifier ».
2. **Ticket** : clé dans l'objet ou le corps (`ASD-27334`), sinon ticket
   Jira assigné dont le rapporteur ou le client correspond
   (`jira-assignes.json`), sinon ticket local ouvert du même client ou du
   même domaine (`tickets` du contexte), sinon alerte Jira « Incident
   raised » arrivée le même jour pour un ticket de ce client. Un
   rapprochement qui ne repose pas sur une clé explicite est **supposé** :
   l'écrire « lié à ASD-xxxx ? (même client, même jour) ».
3. Le **symptôme** et les **notes** du ticket local, et les
   **particularités** de la fiche client, complètent le mail : s'en servir
   pour comprendre le sujet, la priorité et ce qui est attendu.
4. **Point de suivi** : un mail rattaché à un ticket qui a déjà un point
   ouvert dans le registre met à jour **ce point** (une seule ligne par
   ticket). Il ne crée un nouveau point que s'il apporte une demande
   distincte.

### Information complémentaire ou nouvelle demande

Un mail sur un ticket déjà suivi est une **information complémentaire**
quand il ne demande rien de nouveau : photo, pièce jointe, précision sur le
symptôme, confirmation, remerciement, copie d'un échange. Il enrichit le
point existant (« info reçue le 28.10 : factures de l'installateur ») sans
créer de tâche.

C'est une **nouvelle demande** quand il pose une question, demande une
action ou une date, ou change le périmètre. Elle met à jour l'action
attendue du point existant ; un sujet différent sur le même ticket devient
un point séparé.

Une **réponse du client sur un ticket en attente du client** (Jira
« En attente du client », AnsamDiag `EnAttente`, ou point `attente_tiers`)
fait repasser le point en `a_faire` : c'est à moi de reprendre.

## Regrouper les points

Un point peut appartenir à un **groupe** : une tâche plus large dont il est
une sous-tâche. Le groupe se lit d'un coup d'œil (« Revues des tickets :
3/5 ») et s'ouvre pour voir ses sous-tâches.

Regrouper quand plusieurs points relèvent d'**une même activité** :

- **Revues récurrentes** : toutes les tâches « Revue des tickets » de mes
  clients (une par client, deux fois par semaine) vont dans le groupe
  **« Revues des tickets »** ; toutes les « Revue des variabilités » dans
  **« Revues des variabilités »**. Chaque client reste une sous-tâche avec
  son propre statut.
- **Même chantier** : plusieurs points qui font partie d'un même projet,
  d'une même migration ou d'une même demande de la direction (« 8 tickets à
  faire avancer » et les tickets qu'elle cite) vont dans un groupe au nom du
  chantier.

Ne pas regrouper deux points sans lien autre que le client : un client a
souvent plusieurs sujets indépendants.

Dans le registre, un groupe a `estGroupe` = vrai ; une sous-tâche a
`parent` (id du groupe) et `parentPar` (`moi` ou `synthese`). **Un point
rangé ou sorti d'un groupe par moi (`parentPar` = `moi`) ne se déplace
plus** : ne pas lui proposer d'autre groupe. Pour ranger un point, donner le
nom du groupe dans le champ `groupe` du bloc de suivi ; le lanceur retrouve
le groupe existant de même nom ou le crée.

## Statut d'un point

| Statut | Sens |
|---|---|
| `a_faire` | Une action m'attend |
| `attente_tiers` | Mon action est faite, j'attends le client, un collègue, un fournisseur |
| `traite` | Réglé, plus rien à faire |
| `ignore` | Sans objet pour moi (posé par moi seulement, jamais proposé par la synthèse) |

Sources, de la plus forte à la plus faible :

1. **Mon marquage** (`majPar` = `moi`) : il tient tant qu'aucun message
   n'est arrivé sur le point **après** `majLe`. Un nouveau message peut le
   rouvrir, en le disant (« rouvert : nouveau message de X le 30.10 »).
2. **AnsamDiag** : statut `Resolu` ou note « RESOLU » → `traite` ; statut
   `EnAttente` ou note qui dit que l'action est faite et que j'attends →
   `attente_tiers` ; clôture préparée (`cloture`) → `traite`.
3. **Jira** (`suivis` donne le statut réel de chaque ticket suivi, même
   clos ou réassigné) : statut de catégorie « Terminé » (« Résolu »,
   « Clos »…) → `traite` ; ticket réassigné à quelqu'un d'autre → `traite`
   avec le motif « réassigné à X » ; « En attente du client », « En
   attente du fournisseur », « En attente Interne », « Awaiting CAB »,
   « Planifié » → `attente_tiers` (préciser qui ou quoi) ; « En cours »,
   « Ouvert », « À traiter » → `a_faire`.
4. **Les mails** : ma réponse est le dernier message et elle n'appelle
   rien de ma part → `attente_tiers` ; elle clôt le sujet → `traite` ; un
   message plus récent que le mien me demande quelque chose → `a_faire`.

### Commentaire du client après la clôture

À la clôture d'un ticket, le client répond presque toujours (« accepté »,
« merci », questionnaire de satisfaction, notification « nouveau
commentaire sur le ticket clôturé »). **Ce n'est pas une tâche.** Quand il
refuse la résolution, Jira rouvre le ticket : il réapparaît alors ouvert
dans `suivis` ou dans mes tickets.

- Ticket **Terminé** dans `suivis` (ou clôture préparée dans AnsamDiag) et
  dernier événement = commentaire du client → `traite`, motif
  « commentaire après clôture, ticket toujours clos ». Ne jamais rouvrir le
  point pour ça, ni créer de tâche dérivée (« documenter… »).
- Exception : le commentaire, lisible dans le mail, **demande clairement
  autre chose** (nouveau problème, refus explicite) alors que le ticket est
  resté clos → `a_faire`, motif « nouvelle demande après clôture : … ».
- Ticket **rouvert** (statut ouvert dans `suivis` après une note « RESOLU »)
  → `a_faire`, motif « rouvert par le client ».

Quand les sources se contredisent, prendre la plus forte et signaler la
contradiction dans le motif (« Jira dit En cours, AnsamDiag dit Résolu »).
Un point `traite` par une source automatique est listé dans « Réglé depuis
la dernière synthèse » avec son motif : on doit pouvoir le rouvrir.

Un point `attente_tiers` depuis plus de 5 jours ouvrés est signalé
« à relancer ? », sans changer de statut.

Un document d'un tiers (facture d'installateur, bon de livraison, fiche
technique) transmis par un client sert de **contexte au ticket** : il ne
relève pas de la facturation ANSAM.

## A – En attente de moi

Un fil est **en attente de moi** si toutes ces conditions sont réunies :

1. Le dernier message du fil **n'est pas de moi** (vérifier aussi les
   Éléments envoyés : une réponse peut être partie hors du fil).
2. Je suis en **À**, ou je suis nommé·e, ou la demande s'adresse clairement
   à moi en tant que TAM du client. En **Cc** seul : uniquement si on me
   pose explicitement une question, ou s'il s'agit d'une **relance d'un
   client de mon portefeuille** (je dois m'assurer que la réponse part).
3. Le message attend quelque chose : une question, une demande de
   validation, un document, une date, un devis, une intervention, un
   arbitrage, un retour sur une variabilité ou sur une facture.

Cas particulier : j'ai écrit « je reviens vers vous », « je vérifie et je
vous dis »… et rien n'est parti depuis. Le fil est alors en attente de moi,
avec la mention « relance de mon engagement ».

Les rappels Jira qui me concernent directement sont aussi des attentes :
tâche **échue**, ticket **en retard**, **assignation** nouvelle,
**mention** de mon nom.

### Tickets Jira : tous dans le suivi

**Chaque ticket Jira ouvert qui m'est assigné a un point** : le lanceur
l'ajoute avant l'analyse (`origine` = `jira`, statut repris de Jira,
action générique). Ces points sont dans le registre que je reçois : les
**compléter** comme les autres (client, action concrète tirée des mails, des
notes AnsamDiag ou du titre, priorité, contact, groupe), sans les
supprimer ni les fusionner avec un autre ticket.

- Un ticket sans mail ni note récente reste un point ouvert : l'action dit
  ce qu'il faut faire pour le faire avancer (relancer, planifier,
  clôturer s'il est réglé), et un ticket sans activité depuis plus de
  10 jours ouvrés passe au moins en 🟠.
- Un ticket **encore ouvert dans Jira** (hors « Résolu ») n'est jamais
  `traite` : s'il est réglé d'après les mails ou AnsamDiag, le laisser
  `a_faire` avec l'action « Clôturer le ticket dans Jira ». Le lanceur
  rouvre de toute façon un point réglé dont le ticket reste ouvert, sauf
  si c'est moi qui l'ai marqué.
- « Résolu » dans Jira (résolu, en attente de clôture) vaut `traite`.

### Priorité

| Priorité | Critère |
|---|---|
| 🔴 Urgent | Échéance aujourd'hui ou dépassée ; SLA dépassé ; incident en production ; **alarme matérielle sur site client** (onduleur / UPS, batterie, disque, RAID, température, alimentation) ; sauvegarde en échec ; client mécontent ; relance (2e demande ou plus) ; direction en demande |
| 🟠 Cette semaine | Échéance sous 5 jours ouvrés, demande client standard, validation de facture ou de variabilité |
| 🟢 Quand possible | Pas d'échéance, information à confirmer, demande interne non bloquante |

Une **relance** remonte toujours d'un cran. Sans échéance, une attente de
plus de 3 jours ouvrés passe au moins en 🟠. En cas d'hésitation entre deux
niveaux, prendre le plus élevé et dire pourquoi.

### Prise de contact

Un point demande une **prise de contact** quand, pour avancer, je dois
moi-même appeler ou écrire à quelqu'un : relance restée sans réponse,
information à obtenir, rendez-vous ou date à fixer, décision ou accord à
demander, test à faire avec l'utilisateur, **information à transmettre à
un tiers** (opérateur, fournisseur, éditeur : « transmettre à Swisscom
l'adresse de livraison », « confirmer au fournisseur la date »). Le point
reçoit alors un objet `contact` (voir le modèle). Une action qui commence
par transmettre, relancer, appeler, contacter, écrire, demander, informer
ou confirmer à quelqu'un a toujours un `contact` ; à défaut, le lanceur en
crée un vide, à rédiger. Le `contact` :

- `mode` : `appel` quand c'est urgent, qu'un échange de mails a déjà
  échoué, que la personne a demandé à être appelée ou qu'il faut agir avec
  elle ; sinon `mail` ;
- `nom`, `role`, `email`, `telephone` : la bonne personne, pas la boîte
  générique, avec son adresse et son numéro tels qu'ils figurent dans les
  mails (signature), la fiche client (contacts) ou le ticket. Rien
  d'inventé : une donnée introuvable reste `null` ;
- pour un **tiers de l'annuaire** (donné dans la demande : nom, mots-clés,
  adresse, numéro, sujet), son canal officiel est la bonne adresse, même
  générique : `nom` = le nom de l'annuaire, `email` et `telephone` repris
  tels quels, sauf si le fil donne un interlocuteur nommé chez ce tiers pour
  ce dossier. L'annuaire vaut aussi chez un client BWS (ligne, raccordement,
  commande) ; seul un cas ServiceNow ouvert (CS…) se suit dans ServiceNow ;
- pour un mail : `objet` (avec la clé du ticket s'il y en a une) et
  `message`, le brouillon complet, prêt à envoyer : en français, vouvoiement
  pour les clients, court, qui rappelle le contexte en une phrase, pose la
  ou les questions précises, et se termine comme mes mails envoyés
  (formule et signature) ;
- pour un appel : `points`, les 3 à 6 points à aborder, dans l'ordre, et ce
  qu'il faut obtenir à la fin de l'appel.

Jamais de mot de passe, de lien de mot de passe ni de secret dans un
brouillon. Le brouillon est une proposition : rien n'est envoyé sans moi.

### Ce qu'on attend de moi

Une phrase à l'infinitif, concrète et vérifiable : « Confirmer à X le
nombre de postes à facturer pour novembre », pas « répondre au mail ».
S'il y a plusieurs demandes dans le fil, les lister toutes.

## B – À archiver

L'information doit être **durable** : utile dans un mois, lors d'une
variabilité, d'une facturation ou d'un litige. Une simple discussion ne
suffit pas, il faut un fait, un chiffre ou une décision.

### TAM – suivi des clients

- décisions prises avec le client, changements de périmètre ;
- incidents significatifs : cause, résolution, impact ;
- changements d'interlocuteurs côté client (arrivées, départs, rôles) ;
- comptes rendus de réunion, comités de pilotage, plans d'action ;
- engagements pris par ANSAM ou par le client, avec leurs dates ;
- documents de tiers utiles à un ticket (facture d'installateur, garantie).

### Variabilité – inventaire mensuel par client

En fin de mois, chaque TAM vérifie pour chacun de ses clients que
l'inventaire facturable est à jour. Éléments concernés :

- arrivées et départs d'utilisateurs, créations et suppressions de comptes ;
- postes de travail ajoutés, remplacés, retirés ou réattribués ;
- serveurs, VM et équipements réseau ajoutés ou décommissionnés ;
- licences (M365, antivirus, sauvegarde…) ajoutées ou résiliées ;
- ressources de VM (disque, mémoire) quand le contrat les facture ;
- toute confirmation ou correction d'inventaire envoyée par le client.

Deux tableaux distincts :

- **Établi** : le mail l'affirme (« Mme X quitte la société le 31.10 »,
  « le nouveau serveur est en production »). Noter le client, le type, le
  sens (+ / −), la quantité, la date d'effet et la source.
- **Suggestions à vérifier** : tout indice qu'un élément d'inventaire a pu
  changer, même indirect (départ mentionné en passant, ticket « Création
  compte » dont le mail ne dit rien de plus, équipement mis en service,
  migration avec anciens serveurs « conservés quelques jours », extension
  de disque proposée). Pour chacun : l'indice exact, et **comment le
  vérifier** (compte dans Entra, inventaire du client, ticket Jira,
  question au client).

### Facturation

Ce qu'ANSAM facture ou se fait facturer :

- factures émises ou reçues, avoirs, notes de crédit ;
- devis et offres, bons de commande, acceptations ;
- contestations, rappels, retards de paiement ;
- changements de tarif, renouvellements et résiliations de contrat ;
- toute donnée qui modifie le montant facturé d'un client.

### Autre information d'activité

Procédures, accès et documentation fournis par un client ou un
fournisseur, informations contractuelles, RH utiles (décompte de salaire,
part variable, notes de frais, congés).

### Dossier d'archivage proposé

Proposer un chemin du type `Clients/<Client>/<Catégorie>` (TAM,
Variabilité, Facturation, Contrats), ou `Interne/<Catégorie>`. On reprend la
structure des sous-dossiers existants de la boîte quand elle est connue.

## C – Bruit

Est du bruit **uniquement** ce qui entre dans une de ces catégories :

- newsletters, publicités, invitations à des événements externes ;
- publications Viva Engage / Yammer et synthèses Confluence sans mention
  de mon nom ;
- annonces collectives sans action (départs, arrivées, vie d'entreprise) ;
- notifications Jira **sans auteur humain externe à ANSAM et sans action
  pour moi** : « Automation for Jira a fait une mise à jour », « Issue …
  has not been updated in 5 days » sur une tâche interne non échue ;
- alertes d'opérateur **closes** (« de nouveau disponible »), accusés de
  lecture, invitations déjà acceptées ;
- mes propres synthèses « Synthèse mails – … ».

**N'est jamais du bruit**, même si l'expéditeur est automatique : une
mention de mon nom, une assignation, une tâche ou un ticket échu ou en
retard, un commentaire d'un client (adresse externe dans le nom de
l'auteur Jira), une alerte « Incident raised » ou « New alert », un ticket
clôturé qui reçoit un nouveau commentaire, un message Teams qui m'est
adressé, un mail avec drapeau de suivi.

**En cas de doute, ce n'est pas du bruit** : classer en A, en B ou dans les
points d'attention.

Le bruit est listé en annexe de la synthèse (voir le modèle) pour qu'on
puisse vérifier d'un coup d'œil que rien d'important n'y est tombé.
