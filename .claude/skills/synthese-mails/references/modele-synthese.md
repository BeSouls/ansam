# Modèle de synthèse

Respecter cette structure. Omettre une section vide en écrivant
« Rien aujourd'hui. » Chaque sujet est un lien vers le mail quand la source
en fournit un ; avec l'export local, indiquer à la place le dossier Outlook
et la date de réception (ex. `Boîte de réception/Clients/ABC · 27.10 09:14`).

```markdown
# Synthèse mails – <jour JJ.MM.AAAA>

**En bref** : <n> échanges attendent une action de ma part (<n> 🔴), <n> informations à archiver, <n> mails écartés (bruit).

## 1. En attente de moi

| Prio | Client / émetteur | Sujet | Reçu | Attend depuis | Ce qu'on attend de moi | Échéance |
|---|---|---|---|---|---|---|
| 🔴 | Client SA – Jean Dupont | [Inventaire novembre](lien) | 27.10 | 1 j | Confirmer le nombre de postes à facturer (52 ou 54 ?) | 31.10 |

## 2. À archiver

### TAM
- **Client SA** – [Sujet](lien) – <fait clé en une phrase>. → `Clients/Client SA/TAM`

### Variabilité (<mois>)
| Client | Élément | Sens | Qté | Date d'effet | Source |
|---|---|---|---|---|---|
| Client SA | Utilisateur (M. Martin) | + | 1 | 01.11 | [mail](lien) |

### Facturation
- **Client SA** – [Facture 2026-1042](lien) – CHF 4 250.– HT, échéance 30.11. → `Clients/Client SA/Facturation`

### Autres informations
- …

## 3. Points d'attention
- Mails suspects, éléments incertains, clients absents de `clients.md`.
- Du 25 du mois au 3 du mois suivant : liste des clients de `clients.md`
  sans aucune information de variabilité reçue ce mois-ci.
```
