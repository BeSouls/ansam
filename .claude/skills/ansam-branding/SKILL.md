---
name: ansam-branding
description: Charte graphique ANSAM pour tout document Word (.docx) produit pour ANSAM ou ses clients – guides, procédures, documentations techniques, offres, rapports, décharges, comptes rendus. Utiliser dès qu'un document est rédigé « pour ANSAM », « au branding / aux couleurs ANSAM », à destination d'un client d'ANSAM, ou quand l'utilisateur·rice demande une doc, un guide, une procédure ou un livrable Word dans un contexte ANSAM. Fournit le template officiel (logos, en-têtes, pieds de page, polices Montserrat, rose #D7007F) et une bibliothèque Python qui génère le document directement à partir du template.
---

# Branding ANSAM – documents Word

Tout document ANSAM se construit **à partir du template officiel** `assets/ANSAM_Modele.docx`, jamais d'un document vierge. Le template porte l'identité visuelle qu'il ne faut pas recréer à la main : logos ANSAM group dans l'en-tête et la couverture, pied de page ANSAM Holding SA (adresse, agences, TVA, téléphone), polices et styles.

## Méthode

1. Rédiger le contenu (français, vouvoiement pour les clients).
2. Générer le .docx avec `scripts/ansam_doc.py` (bibliothèque standard Python, aucune dépendance) :

```python
import sys; sys.path.insert(0, "<dossier de cette skill>/scripts")
from ansam_doc import AnsamDoc

d = AnsamDoc()
d.cover("Titre du document", "Sous-titre",
        recipient=["Client SA", "Adresse", "NPA Localité"],      # bloc « À L’ATTENTION DE »
        meta=[("Objet", "…"), ("Version / date", "1.0 – JJ.MM.AAAA")])
d.h1("Section", page_break=True)   # Titre1 du template
d.h2("Sous-section")               # Titre2
d.p("Texte avec **gras** et `code`.")
d.bullets([...]); d.steps([...])  # puces du template / liste 1. 2. 3.
d.table(["Col A", "Col B"], rows, widths=[3000, 6072])   # somme = 9072
d.info("Titre", ["…"]); d.warn("Titre", ["…"])
d.placeholder("[Capture à insérer : …]")
d.image("capture.png", "Légende", width_cm=15)
d.qr_box()                         # cadre « COLLER ICI le QR code »
d.save("sortie.docx")
```

`scripts/example.py` produit un document de démonstration qui utilise tous les blocs.

3. Vérifier le résultat : si le skill `docx` est disponible, lancer son `validate.py` ; sinon, contrôler que le .docx s'ouvre (zip valide, XML bien formé).

## Charte (à respecter si on édite ou complète un document à la main)

| Élément | Règle |
|---|---|
| Couleur principale | Rose ANSAM `#D7007F` (titres, en-têtes de tableaux, accents) |
| Police | **Montserrat** 11 pt (texte, justifié) ; **Montserrat Black** pour les titres |
| Titre de couverture | Montserrat Black 28 pt, rose, centré, sous le logo ANSAM group |
| Titres de section | Style `Titre1` (Montserrat Black 14 pt rose), **sans numérotation** (`numId 0`) ; `Titre2` gras rose |
| Couverture | Logo ANSAM group centré → titre → sous-titre → « **À L’ATTENTION DE :** » + destinataire en gras centré → tableau récapitulatif (objet, version/date) |
| Tableaux | En-tête fond `#D7007F`, texte blanc gras ; bordures `#BFBFBF` ; largeur utile 9072 twips (A4, marges 2,5 cm) |
| Encadré info | Fond rose clair `#FBE5F2`, barre gauche rose `#D7007F` |
| Encadré alerte | Fond `#FFF2CC`, barre gauche `#BF8F00`, titre préfixé « ⚠ » |
| Emplacement à compléter | Cadre gris `#F2F2F2`, bordure pointillée, texte italique « [Capture / Photo à insérer : …] » |
| Légendes | Italique gris `#595959`, 9 pt, centrées |
| Puces | Liste à puces du template (`numId 14`) ; numérotation « 1. » (abstractNum 3) |
| En-têtes / pieds | Ceux du template, **ne jamais les modifier ni les supprimer** |
| Contact ANSAM | +41 58 911 88 00 (pied de page du template) |

## Bonnes pratiques de contenu

- Documents client : pas de jargon, étapes courtes, une action par étape, un tableau « En cas de problème » et un encadré de contact support à la fin.
- Ne jamais écrire de secret dans un document (mot de passe, token, code temporaire) : prévoir une ligne à compléter à la main ou indiquer « transmis séparément ».
- Pour les captures qu'on n'a pas encore : utiliser `placeholder()` avec une description précise de la capture attendue.
- Nom de fichier conseillé : `NN_Titre_Court_ANSAM.docx`.
