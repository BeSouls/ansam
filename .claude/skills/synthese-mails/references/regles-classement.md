# Règles de classement

## A – En attente de moi

Un fil est **en attente de moi** si toutes ces conditions sont réunies :

1. Le dernier message du fil **n'est pas de moi** (vérifier aussi les
   Éléments envoyés : une réponse peut être partie hors du fil).
2. Je suis en **À**, ou je suis nommé·e, ou la demande s'adresse clairement
   à moi en tant que TAM du client. En **Cc** seul : uniquement si on me
   pose explicitement une question.
3. Le message attend quelque chose : une question, une demande de
   validation, un document, une date, un devis, une intervention, un
   arbitrage, un retour sur une variabilité ou sur une facture.

Cas particulier : j'ai écrit « je reviens vers vous », « je vérifie et je
vous dis »… et rien n'est parti depuis. Le fil est alors en attente de moi,
avec la mention « relance de mon engagement ».

### Priorité

| Priorité | Critère |
|---|---|
| 🔴 Urgent | Échéance aujourd'hui ou dépassée, incident en production, client mécontent, relance (2e demande ou plus), direction en demande |
| 🟠 Cette semaine | Échéance sous 5 jours ouvrés, demande client standard, validation de facture ou de variabilité |
| 🟢 Quand possible | Pas d'échéance, information à confirmer, demande interne non bloquante |

Une **relance** remonte toujours d'un cran. Sans échéance, une attente de
plus de 3 jours ouvrés passe au moins en 🟠.

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
- engagements pris par ANSAM ou par le client, avec leurs dates.

### Variabilité – inventaire mensuel par client

En fin de mois, chaque TAM vérifie pour chacun de ses clients que
l'inventaire facturable est à jour. À archiver :

- arrivées et départs d'utilisateurs, créations et suppressions de comptes ;
- postes de travail ajoutés, remplacés, retirés ou réattribués ;
- serveurs, VM et équipements réseau ajoutés ou décommissionnés ;
- licences (M365, antivirus, sauvegarde…) ajoutées ou résiliées ;
- toute confirmation ou correction d'inventaire envoyée par le client.

Pour chaque élément, noter le **client**, le **type** (user, poste,
serveur, licence…), le **sens** (+ / −), la **quantité**, la **date
d'effet** et la source.

### Facturation

- factures émises ou reçues, avoirs, notes de crédit ;
- devis et offres, bons de commande, acceptations ;
- contestations, rappels, retards de paiement ;
- changements de tarif, renouvellements et résiliations de contrat ;
- toute donnée qui modifie le montant facturé d'un client.

### Autre information d'activité

Procédures, accès et documentation fournis par un client ou un
fournisseur, informations contractuelles, RH utiles (décompte de salaire,
part variable, notes de frais).

### Dossier d'archivage proposé

Proposer un chemin du type `Clients/<Client>/<Catégorie>` (TAM,
Variabilité, Facturation, Contrats), ou `Interne/<Catégorie>`. On reprend la
structure des sous-dossiers existants de la boîte quand elle est connue.
