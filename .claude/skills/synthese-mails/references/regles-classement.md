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
   local ouvert du même client ou du même domaine (`tickets`), sinon alerte
   Jira « Incident raised » arrivée le même jour pour un ticket de ce
   client. Un rapprochement qui ne repose pas sur une clé explicite est
   **supposé** : l'écrire « lié à ASD-xxxx ? (même client, même jour) ».
3. Le **symptôme** et les **notes** du ticket local, et les
   **particularités** de la fiche client, complètent le mail : s'en servir
   pour comprendre le sujet, la priorité et ce qui est attendu.

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

### Priorité

| Priorité | Critère |
|---|---|
| 🔴 Urgent | Échéance aujourd'hui ou dépassée ; SLA dépassé ; incident en production ; **alarme matérielle sur site client** (onduleur / UPS, batterie, disque, RAID, température, alimentation) ; sauvegarde en échec ; client mécontent ; relance (2e demande ou plus) ; direction en demande |
| 🟠 Cette semaine | Échéance sous 5 jours ouvrés, demande client standard, validation de facture ou de variabilité |
| 🟢 Quand possible | Pas d'échéance, information à confirmer, demande interne non bloquante |

Une **relance** remonte toujours d'un cran. Sans échéance, une attente de
plus de 3 jours ouvrés passe au moins en 🟠. En cas d'hésitation entre deux
niveaux, prendre le plus élevé et dire pourquoi.

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
