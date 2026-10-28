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

```markdown
# Synthèse mails – <jour JJ.MM.AAAA>

**En bref** : <n> échanges attendent une action de ma part (<n> 🔴), <n> informations à archiver, <n> suggestions à vérifier, <n> mails écartés (bruit, liste en annexe).

## 1. En attente de moi

| Prio | Client / émetteur | Sujet | Reçu | Attend depuis | Ce qu'on attend de moi | Échéance |
|---|---|---|---|---|---|---|
| 🔴 | ABC – Jean Dupont | Onduleur, lié à [ASD-12345](https://ansamgroup.atlassian.net/browse/ASD-12345) (`Boîte de réception · 28.10 10:47`) | 28.10 | 0 j | Répondre sur l'onduleur (batterie en fin de vie) et retrouver l'installateur grâce à la facture jointe | aujourd'hui |

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
Regroupé par catégorie, avec le nombre et les objets distincts :
- **Notifications Jira automatiques** (<n>) : ASD-20194 « Création compte » ×3, …
- **Viva Engage / Confluence** (<n>) : …
- **Opérateurs, alertes closes** (<n>) : …
```

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
