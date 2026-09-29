"""Exemple : génère un document de démonstration au branding ANSAM.

    python example.py [sortie.docx]
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ansam_doc import AnsamDoc  # noqa: E402

out = sys.argv[1] if len(sys.argv) > 1 else "Exemple_ANSAM.docx"

d = AnsamDoc()
d.cover(
    "Titre du document",
    "Sous-titre ou objet du document",
    recipient=["Nom du client SA", "Rue de l’Exemple 1", "1000 Lausanne"],
    meta=[("Objet", "Démonstration du branding ANSAM"), ("Version / date", "1.0 – 29.09.2026")],
)
d.h1("Introduction", page_break=True)
d.p("Paragraphe standard en Montserrat, avec du **gras** et du `code`.")
d.bullets(["Premier point", "Deuxième point"])
d.h2("Procédure")
d.steps(["Première étape", "Deuxième étape", "Troisième étape"])
d.table(["Élément", "Valeur"], [["Plateforme", "Microsoft Intune"], ["Statut", "Terminé"]], [3000, 6072])
d.info("Bon à savoir", ["Encadré d’information aux couleurs ANSAM."])
d.warn("Attention", ["Encadré d’avertissement."])
d.placeholder("[Capture à insérer : description de la capture]")
d.qr_box()
d.save(out)
print("OK", out)
