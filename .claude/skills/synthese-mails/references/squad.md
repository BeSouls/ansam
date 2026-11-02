# Préparation de la réunion squad

Chaque semaine, la squad (BC, SDM, responsable, TAM) passe en revue un
formulaire commun. Chacun y remplit ses lignes. Tu prépares **les miennes**
(mes initiales sont données dans la demande) à partir de ce que je
fournis : agenda Outlook, mails, registre de suivi, éléments à garder
(variabilité, facturation, attention), tickets Jira, contexte AnsamDiag
(fiches clients : code, nom, modèle ASW / BWS / SOM), fiche de la réunion
précédente.

## Règles

- **Que des faits sourcés.** Chaque ligne a une `source` courte et
  vérifiable : « agenda 03.11 », « mail de K. Lötscher du 27.10 »,
  « suivi ASD-27975 », « variabilité 10.2026 ». Rien de supposé, rien
  d'extrapolé. Sans source, pas de ligne.
- **Mon périmètre seulement** : mes clients, mes rendez-vous, mes projets,
  ce que je porte. Les lignes des autres membres ne te regardent pas.
- **Le style du formulaire** : télégraphique, une ligne par sujet, le code
  client d'abord, en majuscules : « CIMO : revue mensuelle avec FCF »,
  « AJE : attente retour du Conseil intercommunal ». Dates en JJ.MM.
  Pas de phrase d'introduction, pas de jargon interne (pas d'AnsamDiag,
  de Panoramix, de registre).
- `etape` : la prochaine étape concrète, avec sa date si elle est connue
  (« Installation par l'éditeur à planifier »), sinon vide.
- `porteur` : mes initiales par défaut ; les initiales ou le nom d'un autre
  quand c'est lui qui porte la suite (BC, SDM, OPS).
- Au plus 6 lignes par sujet : l'essentiel pour la squad, pas chaque ticket.
- Jamais de secret (mot de passe, lien d'accès, clé).
- Un sujet sans rien d'établi reste vide : ne pas remplir pour remplir.

## Sujets (champ `sujet`)

- `rdv_ecoulee` — mes rendez-vous **clients** de la semaine écoulée (du
  lundi au vendredi précédant la réunion), d'après l'agenda : rendez-vous
  avec au moins un participant externe ou un code client dans l'objet,
  accepté ou organisé par moi. Pas les réunions internes, les absences, ni
  la squad. « MCRIN : SOM BWS, point avec M. X (Teams) ».
- `rdv_en_cours` — mes rendez-vous clients de la semaine de la réunion
  (lundi au vendredi), même règle.
- `visites` — mes visites sur site (lieu = adresse d'un client, « visite »
  dans l'objet), faites la semaine écoulée ou prévues dans les quatre
  semaines : « MCRIN : visite du TAM le 03.11 à Vevey ».
- `revues` — mes revues opérationnelles ou mensuelles prévues dans les
  quatre semaines (agenda) ou à organiser (mails, suivi).
- `projets` — mes projets, commandes et offres en cours : où on en est et
  ce qui bloque (« AVE : VM cartes d'identification livrée le 02.11,
  installation par l'éditeur à planifier »). Sources : mails, suivi.
- `problematiques` — incidents ou sujets sensibles en cours chez mes
  clients : point urgent du suivi, escalade, incident récurrent, impact
  client. Pas les demandes courantes.
- `variabilites` — variabilité du mois à reporter (éléments à garder de
  type variabilité, à traiter), regroupée par client.
- `facturation` — points de facturation ouverts (éléments à garder de
  type facturation, à traiter), demandes de participation financière,
  contestations.
- `satisfaction` — satisfaction ou insatisfaction exprimée par un client
  dans les mails de la période (remerciement, plainte, relance agacée).
  Citer le fait, pas une interprétation.
- `mercato` — reprises ou transferts de clients qui me concernent
  (nouveau client repris, client transmis), d'après les mails.
- `divers` — ce que la squad doit savoir : mes absences prévues dans les
  quatre semaines (réglages ou agenda), changement d'interlocuteur chez un
  client, information d'équipe.
- `clients_a`, `rooster` — seulement s'il y a un élément explicite.

Pour un sujet qui continue depuis la réunion précédente (fiche fournie),
reprends-le s'il est toujours d'actualité, mis à jour avec les faits
nouveaux ; s'il est réglé, ne le reprends pas.

## Réponse

Uniquement un objet JSON :

```json
{
  "resume": "2 à 4 lignes : ce qui marque la semaine pour moi",
  "lignes": [
    { "sujet": "rdv_en_cours", "situation": "MCRIN : visite du TAM", "etape": "", "porteur": "GBR", "source": "agenda 03.11" }
  ]
}
```
