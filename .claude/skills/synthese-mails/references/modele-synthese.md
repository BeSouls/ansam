# Modèle de synthèse

Respecter cette structure : la synthèse est envoyée telle quelle par mail
(la première ligne devient l'objet). Omettre une section vide en écrivant
« Rien aujourd'hui. » Chaque sujet est un lien vers le mail quand la source
en fournit un ; avec l'export local, indiquer à la place le dossier Outlook
et la date de réception (ex. `Boîte de réception\CLIENTS\ABC · 27.10 09:14`).
Les clés de ticket sont des liens Jira
(`https://ansamgroup.atlassian.net/browse/ASD-12345`).

Ce qui est supposé porte la mention **à vérifier** (voir
`regles-classement.md`).

~~~markdown
# Synthèse mails – <jour JJ.MM.AAAA>

**En bref** : <n> points à faire (<n> 🔴, dont <n> nouveaux), <n> en attente d'un tiers, <n> réglés depuis la dernière synthèse, <n> informations à archiver, <n> suggestions à vérifier, <n> mails écartés (bruit, liste en annexe).

## 1. En attente de moi

### À faire
Tous les points `a_faire` du registre et les nouveaux, par priorité. « Nouveau » marque un point apparu aujourd'hui ; « info » signale une information complémentaire reçue sur un point existant.

| Prio | Point | Client / émetteur | Sujet | Attend depuis | Ce qu'on attend de moi | Échéance |
|---|---|---|---|---|---|---|
| 🔴 | `ASD-12345` nouveau | ABC – Jean Dupont | Onduleur, [ASD-12345](https://ansamgroup.atlassian.net/browse/ASD-12345) (`Boîte de réception · 28.10 10:47`) | 0 j | Répondre sur l'onduleur (batterie en fin de vie) et retrouver l'installateur grâce à la facture jointe | aujourd'hui |

Un **groupe** tient sur une seule ligne : son nom, l'avancement et les
sous-tâches ouvertes en résumé, par exemple :

| 🟠 | `G-20261029-1` | 5 clients | Revues des tickets (0/5 faites) | 2 j | Revues ouvertes : STJAC, SWPI, SLV, PREST, ZENIT | 30.10 |

### En attente d'un tiers
| Point | Client | Sujet | J'attends | Depuis | Source du statut |
|---|---|---|---|---|---|
| `ASD-12346` | ABC | Création de compte | Le nom du manager (client) | 24.10 (4 j, à relancer ?) | Jira « En attente du client » |

### Réglé depuis la dernière synthèse
Points fermés automatiquement, avec la raison, pour pouvoir les rouvrir dans la page de suivi.
- `ASD-12347` ABC – Export gDoc : réglé (AnsamDiag « RESOLU 28.10 16:53 »).

## 2. À archiver

### TAM
- **Client SA** – [Sujet](lien) – <fait clé en une phrase>. → `Clients/Client SA/TAM`

### Variabilité (<mois>) – établie
| Client | Élément | Sens | Qté | Date d'effet | Source |
|---|---|---|---|---|---|
| Client SA | Utilisateur (M. Martin) | + | 1 | 01.11 | `Boîte de réception · 27.10 09:14` |

### Variabilité (<mois>) – suggestions à vérifier
| Client | Élément possible | Sens | Indice | Comment vérifier |
|---|---|---|---|---|
| Client SA | Départ de Mme X | − | « Mme X ne travaille plus chez nous » (mail du 26.10) | Compte et licence dans Entra |

### Facturation
- **Client SA** – [Facture 2026-1042](lien) – CHF 4 250.– HT, échéance 30.11. → `Clients/Client SA/Facturation`

### Autres informations
- …

## 3. Points d'attention
- Mails suspects, éléments incertains, clients non identifiés.
- Notifications Jira sans contenu à lire dans Jira (clé, auteur, ce qui a changé).
- Du 25 du mois au 3 du mois suivant : clients du portefeuille sans aucune
  information de variabilité reçue ce mois-ci.

## Annexe – Bruit écarté (<n>)
Regroupé par catégorie, avec le nombre et les objets distincts, sur toute la période des nouveautés :
- **Notifications Jira automatiques** (<n>) : ASD-20194 « Création compte » ×3, …
- **Viva Engage / Confluence** (<n>) : …
- **Opérateurs, alertes closes** (<n>) : …

<!-- SUIVI -->
```json
[
  {"id": "ASD-12345", "titre": "Onduleur en alarme", "client": "ABC", "ticket": "ASD-12345",
   "statut": "a_faire", "priorite": "rouge", "action": "Répondre sur l'onduleur et retrouver l'installateur",
   "echeance": "2026-10-28", "source": "Boîte de réception · 28.10 10:47",
   "dernierMessage": "2026-10-28T14:14:00+01:00", "motif": "Nouvelle demande du client"}
]
```
~~~

## Bloc de suivi

Après la synthèse, une ligne `<!-- SUIVI -->` puis un tableau JSON (dans un
bloc de code ou non) avec :

- **tous les points ouverts** du registre (`a_faire`, `attente_tiers`),
  même sans nouveauté, avec leur statut revu ;
- **tous les nouveaux points** ;
- les points que je dois voir passer à `traite` aujourd'hui.

Champs : `id` (celui du registre ; pour un nouveau point, la clé du ticket
s'il y en a un, sinon `null`), `titre`, `client`, `ticket` (ou `null`),
`statut` (`a_faire`, `attente_tiers`, `traite` ; jamais `ignore`),
`groupe` (nom du groupe où ranger le point, par exemple « Revues des
tickets », ou `null` : voir « Regrouper les points » dans les règles ; ne
jamais renvoyer les groupes eux-mêmes, leur état se calcule),
`priorite` (`rouge`, `orange`, `vert`), `action` (ce qu'on attend de moi, ou
ce que j'attends), `echeance` (`AAAA-MM-JJ` ou `null`), `source` (dossier
et date du mail le plus récent), `dernierMessage` (date ISO du dernier
message reçu sur le point, `null` sans mail), `motif` (pourquoi ce statut,
avec la source : « Jira En attente du client », « AnsamDiag RESOLU 28.10 »,
« ma réponse du 27.10 », « nouveau message du client le 29.10 »).

Ne pas proposer de changement de statut à un point marqué par moi
(`majPar` = `moi`) sans message reçu après `majLe` : le lanceur l'ignorerait.

## Bloc des éléments à garder

Après le bloc de suivi, une ligne `<!-- ARCHIVES -->` puis un tableau JSON
avec **chaque élément des sections « À archiver » et « Points d'attention »
de la synthèse du jour** : variabilité établie et suggestions à vérifier,
facturation, faits TAM à archiver, points d'attention. La page de suivi les
affiche dans ses vues Variabilité, Facturation et Attention, où je les
marque reportés, traités ou écartés.

Avant d'ajouter un élément, chercher dans `archives.json` s'il y est déjà
(même client, même élément) : reprendre alors son `id` pour le mettre à
jour au lieu d'en créer un doublon. Ne pas renvoyer un élément déjà
enregistré qui n'a pas changé.

Champs :

- `id` : celui du registre, ou `null` pour un nouvel élément ;
- `type` : `variabilite`, `facturation`, `tam` ou `attention` ;
- `certitude` : `etabli` (le mail l'affirme) ou `a_verifier` (indice) ;
- `client` (code) ;
- `element` : ce dont il s'agit, en quelques mots (« Départ de Mme X »,
  « Facture 2026-1042 », « SwissTransfer CBERNARD expire ») ;
- `sens` (`+`, `−` ou `null`), `quantite`, `montant` (avec la devise),
  `dateEffet` et `echeance` (`AAAA-MM-JJ` ou `null`) ;
- `mois` : mois de variabilité concerné (`AAAA-MM`), pour le type
  `variabilite` ;
- `source` : dossier et date du mail ;
- `indice` : la phrase ou le fait qui a fait remonter l'élément ;
- `verifier` : comment le vérifier (pour `a_verifier`) ;
- `dossier` : classement proposé (`Clients/ABC/Variabilité`) ;
- `ticket` : clé du ticket lié, ou `null`.

~~~
<!-- ARCHIVES -->
[
  {"id": null, "type": "variabilite", "certitude": "a_verifier", "client": "ABC", "element": "Départ de Mme X",
   "sens": "−", "quantite": 1, "montant": null, "dateEffet": null, "echeance": null, "mois": "2026-10",
   "source": "Boîte de réception\\CLIENTS\\ABC · 26.10 16:06", "indice": "« ne travaille plus chez ABC »",
   "verifier": "Compte et licence dans Entra", "dossier": "Clients/ABC/Variabilité", "ticket": null}
]
~~~

## Variante « retour d'absence »

Quand la demande indique un retour d'absence (du <début> au <fin>), la
synthèse couvre toute la période depuis la dernière synthèse et commence
par une section 0, avant la section 1 :

```markdown
# Synthèse mails – <jour JJ.MM.AAAA> – retour d'absence (<début> → <fin>)

**En bref** : …

## 0. Retour d'absence

### À faire en priorité
Les 5 à 10 actions les plus importantes de la période, dans l'ordre où
les traiter, une ligne chacune (client, action, échéance).

### Réglé pendant mon absence
Fils où quelqu'un d'autre a répondu ou traité la demande : une ligne
chacun, pour savoir que rien n'est à faire.

### Archivables sans action
Mails à classer directement, sans rien à faire : une ligne chacun avec le
dossier proposé (`Boîte de réception · 10.11 08:12` → `Clients/ABC/TAM`).
```

La suite (sections 1 à 3 et annexe) reste identique, sur toute la période.
